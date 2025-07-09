using JetBrains.Annotations;

namespace Bluehill.NativeCom;

/// <summary>
/// Specifies a COM class factory for a class.
/// </summary>
/// <remarks>
/// Applying this attribute to a class enables the automatic generation of the
/// <c>DllGetClassObject</c> and <c>DllCanUnloadNow</c> functions, which are
/// critical for the implementation of In-process COM servers.
/// </remarks>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = false)]
public sealed class ClassFactoryAttribute : Attribute {
    /// <summary>
    /// Initializes a new instance of the <see cref="ClassFactoryAttribute"/> class with the specified COM class type.
    /// </summary>
    /// <param name="classType">
    /// The <see cref="Type"/> of the COM class for which the class factory is specified.
    /// </param>
    public ClassFactoryAttribute(Type classType) => ClassType = classType;

    /// <summary>
    /// Gets the <see cref="Type"/> of the COM class associated with this class factory.
    /// </summary>
    /// <value>
    /// The <see cref="Type"/> representing the COM class for which the class factory is specified.
    /// </value>
    [UsedImplicitly]
    public Type ClassType { get; }
}
