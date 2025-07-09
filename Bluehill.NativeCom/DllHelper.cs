using JetBrains.Annotations;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Bluehill.NativeCom;

/// <summary>
/// Contains useful methods for implementing the <c>DllGetClassObject</c> function and the <see cref="IClassFactory"/> interface.
/// </summary>
public static unsafe class DllHelper {
    private const int E_UNEXPECTED = unchecked((int)0x8000FFFF);
    private const int CLASS_E_NOAGGREGATION = unchecked((int)0x80040110);
    private const int CLASS_E_CLASSNOTAVAILABLE = unchecked((int)0x80040111);
    private static readonly StrategyBasedComWrappers Sbcw = new();

    /// <summary>
    /// Creates an instance of a COM object that implements the specified interface and class.
    /// </summary>
    /// <typeparam name="TInterface">
    /// The type of the interface that the created object should implement. This must be an interface
    /// marked with the <see cref="GeneratedComInterfaceAttribute"/>.
    /// </typeparam>
    /// <typeparam name="TClass">
    /// The type of the class that implements the specified interface. This must be a class
    /// marked with the <see cref="GeneratedComClassAttribute"/> and must have a parameterless constructor.
    /// </typeparam>
    /// <param name="pUnkOuter">
    /// A pointer to the outer unknown interface for aggregation. This must be <c>null</c> as aggregation is not supported.
    /// </param>
    /// <param name="riid">
    /// A pointer to the interface identifier (IID) of the interface to be used to communicate with the object.
    /// </param>
    /// <param name="ppvObject">
    /// A pointer to a variable that receives the interface pointer requested in <paramref name="riid"/>.
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
    /// <remarks>
    /// This method is typically used to implement the <see cref="IClassFactory.CreateInstance"/> method.
    /// </remarks>
    [UsedImplicitly]
    public static int CreateInstanceHelper<TInterface, TClass>(void* pUnkOuter, Guid* riid, void** ppvObject) where TClass : class, new() {
        *ppvObject = null;

        if (pUnkOuter is not null) {
            return CLASS_E_NOAGGREGATION;
        }

        var interfaceType = typeof(TInterface);

        if (!interfaceType.IsInterface
            || interfaceType.CustomAttributes.All(a => a.AttributeType != typeof(GeneratedComInterfaceAttribute))
            || typeof(TClass).CustomAttributes.All(a => a.AttributeType != typeof(GeneratedComClassAttribute))) {
            return E_UNEXPECTED;
        }

        if (*riid != interfaceType.GUID) {
            return CLASS_E_CLASSNOTAVAILABLE;
        }

        TClass instance = new();
        var ptr1 = Sbcw.GetOrCreateComInterfaceForObject(instance, CreateComInterfaceFlags.None);

        if (ptr1 == IntPtr.Zero) {
            return E_UNEXPECTED;
        }

        var hr = Marshal.QueryInterface(ptr1, in *riid, out var ptr2);
        *ppvObject = hr == 0 ? ptr2.ToPointer() : null;

        Marshal.Release(ptr1);

        return hr;
    }

    /// <summary>
    /// Retrieves a class factory for the specified CLSID and interface ID.
    /// </summary>
    /// <typeparam name="T">
    /// The type of the class factory that implements the <see cref="IClassFactory"/> interface.
    /// This type must be a class and have a parameterless constructor.
    /// </typeparam>
    /// <param name="rclsid">
    /// A pointer to the CLSID that identifies the class object whose class factory is to be retrieved.
    /// </param>
    /// <param name="riid">
    /// A pointer to the interface ID of the class factory to retrieve. Typically, this is the IID of <see cref="IClassFactory"/>.
    /// </param>
    /// <param name="ppv">
    /// A pointer to a variable that receives the interface pointer requested in <paramref name="riid"/>.
    /// If the method fails, this value is set to <c>null</c>.
    /// </param>
    /// <returns>
    /// An HRESULT indicating success or failure. A value of <c>0</c> indicates success, while a non-zero value indicates an error.
    /// </returns>
    /// <remarks>
    /// This method is typically used in COM server implementations to provide a class factory for a specific CLSID.
    /// It internally uses <see cref="CreateInstanceHelper{TInterface, TClass}"/> to create the class factory instance.
    /// </remarks>
    [UsedImplicitly]
    public static int GetClassFactory<T>(Guid* rclsid, Guid* riid, void** ppv) where T : class, IClassFactory, new() => CreateInstanceHelper<IClassFactory, T>(null, riid, ppv);
}
