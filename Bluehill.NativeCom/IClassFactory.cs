using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Bluehill.NativeCom;

/// <summary>
/// Represents a COM class factory interface that provides methods for creating instances of COM objects
/// and managing their lifetime. This interface is typically used in conjunction with the
/// <see cref="ClassFactoryAttribute"/> and <see cref="DllHelper"/> utility methods to implement
/// COM class factories in .NET.
/// </summary>
/// <remarks>
/// The <c>IClassFactory</c> interface is a standard COM interface that allows clients to create
/// instances of COM objects and control their locking state. It is identified by the GUID
/// "00000001-0000-0000-C000-000000000046" and uses the <c>InterfaceIsIUnknown</c> interface type.
/// </remarks>
[GeneratedComInterface]
[Guid("00000001-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public partial interface IClassFactory {
    /// <summary>
    /// Creates an instance of a COM object that supports the specified interface.
    /// </summary>
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
    /// <item><description><c>E_UNEXPECTED</c> if an unexpected error occurs.</description></item>
    /// </list>
    /// </returns>
    /// <remarks>
    /// This method is typically used to create instances of COM objects that implement a specific interface.
    /// </remarks>
    [PreserveSig]
    unsafe int CreateInstance(void* pUnkOuter, Guid* riid, void** ppvObject);

    /// <summary>
    /// Increments or decrements the lock count of the server. This method is used to control
    /// whether the server can be unloaded from memory. When the lock count is greater than zero,
    /// the server remains loaded in memory.
    /// </summary>
    /// <param name="fLock">
    /// A boolean value indicating whether to increment or decrement the lock count.
    /// Pass <c>true</c> to increment the lock count, or <c>false</c> to decrement it.
    /// </param>
    /// <returns>
    /// An HRESULT value indicating the success or failure of the operation. A value of <c>S_OK</c>
    /// indicates success, while other values indicate an error.
    /// </returns>
    [PreserveSig]
    int LockServer([MarshalAs(UnmanagedType.Bool)] bool fLock);
}
