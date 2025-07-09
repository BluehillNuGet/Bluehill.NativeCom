using JetBrains.Annotations;

namespace Bluehill.NativeCom;

/// <summary>
/// An attribute that specifies the interface and implementation types for a COM class factory.
/// </summary>
/// <remarks>
/// Attaching this attribute will automatically generate <c>DllGetClassObject</c> and <c>DllCanUnloadNow</c> functions,
/// which are essential for In-proc COM servers.
/// </remarks>
[AttributeUsage(AttributeTargets.Class, AllowMultiple = true, Inherited = false)]
[UsedImplicitly]
public sealed class ClassFactoryAttribute : Attribute {
    /// <summary>
    /// Initializes a new instance of the <see cref="ClassFactoryAttribute"/> class,
    /// specifying the interface and implementation types for the factory.
    /// </summary>
    /// <param name="interfaceType">The interface type that the factory is responsible for creating.</param>
    /// <param name="classType">The concrete implementation type that the factory will instantiate.</param>
    public ClassFactoryAttribute(Type interfaceType, Type classType) {
        InterfaceType = interfaceType;
        ClassType = classType;
    }

    /// <summary>
    /// The interface type that the factory is responsible for creating.
    /// </summary>
    /// <value>
    /// The interface type that the factory is responsible for creating.
    /// </value>
    [UsedImplicitly]
    public Type InterfaceType { get; }

    /// <summary>
    /// The concrete implementation type that the factory will instantiate.
    /// </summary>
    /// <value>
    /// The concrete implementation type that the factory will instantiate.
    /// </value>
    [UsedImplicitly]
    public Type ClassType { get; }
}
