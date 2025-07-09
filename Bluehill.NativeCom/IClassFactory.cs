using System.Runtime.InteropServices;
using System.Runtime.InteropServices.Marshalling;

namespace Bluehill.NativeCom;

/// <summary>
/// Represents a COM interface for creating and managing class objects.
/// This interface provides methods to create instances of classes and
/// to control the locking of the server.
/// </summary>
[GeneratedComInterface]
[Guid("00000001-0000-0000-C000-000000000046")]
[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
public partial interface IClassFactory {
    /// <summary>
    /// Creates an instance of a class object that supports the specified interface.
    /// </summary>
    /// <param name="pUnkOuter">
    /// A pointer to the controlling IUnknown interface of the aggregate, or <c>null</c> if the object is not part of an aggregate.
    /// </param>
    /// <param name="riid">
    /// A pointer to the interface identifier (IID) of the interface to be used to communicate with the newly created object.
    /// </param>
    /// <param name="ppvObject">
    /// A pointer to a location that receives the interface pointer requested in <paramref name="riid"/>.
    /// Upon successful completion, this parameter contains the requested interface pointer.
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
    [PreserveSig]
    unsafe int CreateInstance(void* pUnkOuter, Guid* riid, void** ppvObject);

    /// <summary>
    /// Increments or decrements the lock count of the server.
    /// This method is used to control whether the server can be unloaded from memory.
    /// </summary>
    /// <param name="fLock">
    /// A boolean value indicating whether to increment or decrement the lock count.
    /// Pass <c>true</c> to increment the lock count, preventing the server from being unloaded.
    /// Pass <c>false</c> to decrement the lock count, allowing the server to be unloaded if no other locks exist.
    /// </param>
    /// <returns>
    /// An HRESULT value indicating success or failure. A value of <c>S_OK</c> indicates success,
    /// while other values indicate an error.
    /// </returns>
    [PreserveSig]
    int LockServer([MarshalAs(UnmanagedType.Bool)] bool fLock);
}
