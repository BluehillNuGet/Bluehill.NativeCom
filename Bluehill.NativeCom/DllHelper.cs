using System.Collections.Concurrent;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Bluehill.NativeCom;

/// <summary>
/// Provides helper methods for working with dynamically linked libraries (DLLs) and COM interop.
/// This class includes utilities for creating COM instances and managing COM interfaces.
/// </summary>
public static unsafe class DllHelper {
    // ReSharper disable InconsistentNaming

    // HRESULT for "invalid pointer" — used when ppvObject is null.
    private const int E_POINTER = unchecked((int)0x80004003);

    // HRESULT indicating that COM aggregation is not supported for the requested class.
    private const int CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);

    // Generic unexpected failure HRESULT — used when class is not properly attributed or other unexpected conditions occur.
    private const int E_UNEXPECTED = unchecked((int)0x8000FFFF);

    // Success HRESULT.
    private const int S_OK = 0;
    // ReSharper restore InconsistentNaming

    // Cache to memoize whether a given TClass is a COM class (has GeneratedComClassAttribute).
    // This avoids repeated reflection checks across activations.
    private static readonly ConcurrentDictionary<Type, bool> ClassCache = new();

    // Central ComWrappers instance used to create and manage RCW/CCW for generated COM classes.
    // StrategyBasedComWrappers cooperates with source-generated COM marshalling in .NET.
    private static readonly StrategyBasedComWrappers Sbcw = new();

    /// <summary>
    /// Creates a COM instance of the specified class type and retrieves the requested interface pointer.
    /// </summary>
    /// <typeparam name="TClass">
    /// The type of the class to instantiate. This type must have the <see cref="GeneratedComClassAttribute"/> applied.
    /// </typeparam>
    /// <param name="pUnkOuter">
    /// A pointer to the outer unknown interface for aggregation. Must be <c>null</c> for non-aggregated objects.
    /// </param>
    /// <param name="riid">
    /// A pointer to the interface identifier (IID) of the interface being requested.
    /// </param>
    /// <param name="ppvObject">
    /// A pointer to the location where the interface pointer will be stored upon successful creation.
    /// </param>
    /// <returns>
    /// An HRESULT indicating success or failure. Possible values include:
    /// <list type="bullet">
    /// <item><description><c>0</c> (S_OK) if the operation was successful.</description></item>
    /// <item><description><c>CLASS_E_NOAGGREGATION</c> if <paramref name="pUnkOuter"/> is not <c>null</c>.</description></item>
    /// <item><description><c>CLASS_E_CLASSNOTAVAILABLE</c> if the requested interface is not supported.</description></item>
    /// <item><description><c>E_UNEXPECTED</c> if the interface or class is not properly marked or an unexpected error occurs.</description></item>
    /// </list>
    /// </returns>
    public static int CreateInstanceHelper<TClass>(void* pUnkOuter, Guid* riid, void** ppvObject) where TClass : class, new() {
        // Validate out parameter pointer; COM pattern: write null on failure path and return E_POINTER.
        if (ppvObject is null) {
            return E_POINTER;
        }

        // Initialize out parameter to null per COM conventions.
        *ppvObject = null;

        // COM aggregation is not supported; reject if caller passed an outer IUnknown.
        if (pUnkOuter is not null) {
            return CLASS_E_NOAGGREGATION;
        }

        // Ensure the class is a source-generated COM class (must be annotated with GeneratedComClassAttribute).
        if (!IsComClass<TClass>()) {
            return E_UNEXPECTED;
        }

        // Construct the managed instance that will back the COM callable wrapper (CCW).
        TClass instance = new();

        // Ask the ComWrappers to create a COM interface pointer for the managed instance.
        // This returns a primary IUnknown for the CCW.
        var ptr1 = Sbcw.GetOrCreateComInterfaceForObject(instance, CreateComInterfaceFlags.None);

        // Defensive check: if wrapper creation failed for some reason, report failure.
        if (ptr1 == IntPtr.Zero) {
            return E_UNEXPECTED;
        }

        // Query the CCW for the interface identified by riid.
        // This adds a reference to the returned interface pointer on success.
        var hr = Marshal.QueryInterface(ptr1, in *riid, out var ptr2);

        // On success, populate the caller's out pointer with the requested interface.
        if (hr == S_OK) {
            *ppvObject = ptr2.ToPointer();
        }

        // Release the temporary reference to the primary IUnknown we acquired from ComWrappers.
        Marshal.Release(ptr1);

        // Return the final HRESULT from QueryInterface.
        return hr;
    }

    // Checks whether TClass is decorated with GeneratedComClassAttribute, caching the result for performance.
    private static bool IsComClass<TClass>() where TClass : class, new()
        => ClassCache.GetOrAdd(typeof(TClass),
            static _ => Attribute.GetCustomAttribute(typeof(TClass), typeof(GeneratedComClassAttribute), false) is not null);
}
