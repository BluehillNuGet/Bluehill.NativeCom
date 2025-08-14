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
    private const int E_POINTER = unchecked((int)0x80004003);
    private const int CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
    private const int E_UNEXPECTED = unchecked((int)0x8000FFFF);
    private const int S_OK = 0;
    // ReSharper restore InconsistentNaming
    private static readonly ConcurrentDictionary<Type, bool> ClassCache = new();
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
        if (ppvObject is null) {
            return E_POINTER;
        }

        *ppvObject = null;

        if (pUnkOuter is not null) {
            return CLASS_E_NOAGGREGATION;
        }

        if (!IsComClass<TClass>()) {
            return E_UNEXPECTED;
        }

        TClass instance = new();
        var ptr1 = Sbcw.GetOrCreateComInterfaceForObject(instance, CreateComInterfaceFlags.None);

        if (ptr1 == IntPtr.Zero) {
            return E_UNEXPECTED;
        }

        var hr = Marshal.QueryInterface(ptr1, in *riid, out var ptr2);

        if (hr == S_OK) {
            *ppvObject = ptr2.ToPointer();
        }

        Marshal.Release(ptr1);

        return hr;
    }

    private static bool IsComClass<TClass>() where TClass : class, new()
        => ClassCache.GetOrAdd(typeof(TClass),
            _ => Attribute.GetCustomAttribute(typeof(TClass), typeof(GeneratedComClassAttribute), false) is not null);
}
