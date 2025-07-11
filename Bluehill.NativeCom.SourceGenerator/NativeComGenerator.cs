using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;
using System.Collections.Immutable;
using System.Text;

namespace Bluehill.NativeCom.SourceGenerator;

[Generator(LanguageNames.CSharp)]
public sealed class NativeComGenerator : IIncrementalGenerator {
    private static readonly SymbolDisplayFormat FqnFormat
        = new(typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces,
            memberOptions: SymbolDisplayMemberOptions.IncludeContainingType);

    public void Initialize(IncrementalGeneratorInitializationContext context) {
        var value = context.SyntaxProvider.ForAttributeWithMetadataName("Bluehill.NativeCom.ClassFactoryAttribute`1",
            static (n, _) => n is ClassDeclarationSyntax,
            static (c, _) => ((CSharpCompilation)c.SemanticModel.Compilation).LanguageVersion >= LanguageVersion.CSharp11
                ? ((INamedTypeSymbol)c.TargetSymbol, (INamedTypeSymbol?)c.Attributes.Select(ad => ad.AttributeClass!.TypeArguments.Single()).Single())
                : ((INamedTypeSymbol, INamedTypeSymbol?)?)null).Collect();

        context.RegisterSourceOutput(value, GenerateSource);
    }

    private static void GenerateSource(SourceProductionContext context, ImmutableArray<(INamedTypeSymbol, INamedTypeSymbol?)?> array) {
        if (array.Any(e => e?.Item2 is null)) {
            context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0001", "NativeCOM", "Requires C# 11 or higher", DiagnosticSeverity.Error,
                DiagnosticSeverity.Error, true, 0));

            return;
        }

        StringBuilder outerSb = new();

        outerSb.AppendLine("#if !NO_GENERATE_DLL")
            .AppendLine("using Bluehill.NativeCom;").AppendLine()
            .AppendLine("[System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]")
            .AppendLine("internal static unsafe class Dll {")
            .AppendLine("    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]")
            .AppendLine("    public static volatile int Locks;").AppendLine()
            .AppendLine("    [System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllGetClassObject))]")
            .AppendLine("    private static int DllGetClassObject(Guid* rclsid, Guid* riid, void** ppv) {")
            .Append("        ");

        foreach (var tuple in array) {
            var (factory, target) = tuple!.Value;

            if (!factory.Interfaces.Any(ad => ad.ToDisplayString(FqnFormat) == "Bluehill.NativeCom.IClassFactory")) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0002", "NativeCOM", "Class does not implement IClassFactory", DiagnosticSeverity.Error,
                    DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                continue;
            }

            if (!factory.GetAttributes().Any(ad => ad.AttributeClass?.ToDisplayString(FqnFormat) == "System.Runtime.InteropServices.Marshalling.GeneratedComClassAttribute")) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0003", "NativeCOM", "Class does not have GeneratedComClassAttribute", DiagnosticSeverity.Error,
                    DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                continue;
            }

            if (target is null) {
                // Something went Wrong
                throw new InvalidOperationException("This exception shouldn't be thrown");
            }

            if (target.AllInterfaces.Any(i => i.GetAttributes().Any(ad => ad.AttributeClass?.ToDisplayString(FqnFormat) == "System.Runtime.InteropServices.Marshalling.GeneratedComInterfaceAttribute"))) {
                StringBuilder innerSb = new();

                innerSb.Append("namespace ").Append(factory.ContainingNamespace.ToDisplayString(FqnFormat)).AppendLine(";").AppendLine()
                    .Append("partial class ").Append(factory.Name).AppendLine(" {")
                    .Append("    public unsafe int CreateInstance(void* pUnkOuter, Guid* riid, void** ppvObject) => Bluehill.NativeCom.DllHelper.CreateInstanceHelper<")
                    .Append(target.ToDisplayString(FqnFormat)).AppendLine(">(pUnkOuter, riid, ppvObject);").AppendLine()
                    .AppendLine("    public int LockServer(bool fLock) {")
                    .AppendLine("        if (fLock) {")
                    .AppendLine("            Interlocked.Increment(ref Dll.Locks);")
                    .AppendLine("        } else {")
                    .AppendLine("            Interlocked.Decrement(ref Dll.Locks);")
                    .AppendLine("        }").AppendLine()
                    .AppendLine("        return 0;")
                    .AppendLine("    }")
                    .AppendLine("}");

                context.AddSource($"{factory.ToDisplayString(FqnFormat)}.BHNC.g.cs", SourceText.From(innerSb.ToString(), Encoding.UTF8));

                outerSb.Append("if (*rclsid == typeof(").Append(target.ToDisplayString(FqnFormat)).AppendLine(").GUID) {")
                    .Append("            return DllHelper.CreateInstanceHelper<").Append(factory.ToDisplayString(FqnFormat)).AppendLine(">(null, riid, ppv);")
                    .Append("        } else ");
            }
        }

        outerSb.AppendLine(" {")
            .AppendLine("            return -2147221231;")
            .AppendLine("        }")
            .AppendLine("    }").AppendLine()
            .AppendLine("    [System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]")
            .AppendLine("    private static int DllCanUnloadNow() => Locks <= 0 ? 0 : 1;")
            .AppendLine("}")
            .AppendLine("#endif");

        context.AddSource("Dll.g.cs", SourceText.From(outerSb.ToString(), Encoding.UTF8));
    }
}
