using JetBrains.Annotations;

namespace Bluehill.NativeCom;

/// <summary>
/// Specifies a COM class factory for a class.
/// </summary>
/// <typeparam name="TClass">
/// The type of the class for which the COM class factory is specified.
/// The type must be a reference type and have a parameterless constructor.
/// </typeparam>
/// <remarks>
/// Applying this attribute to a class enables the automatic generation of the
/// <c>DllGetClassObject</c> and <c>DllCanUnloadNow</c> functions, which are
/// critical for the implementation of In-process COM servers.
/// </remarks>
[AttributeUsage(AttributeTargets.Class, Inherited = false)]
[UsedImplicitly]
public sealed class ClassFactoryAttribute<TClass> : Attribute where TClass : class, new();
