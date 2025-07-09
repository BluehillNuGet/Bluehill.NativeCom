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
        var value = context.SyntaxProvider.ForAttributeWithMetadataName("Bluehill.NativeCom.ClassFactoryAttribute",
            static (n, _) => n is ClassDeclarationSyntax,
            static (c, _) => ((CSharpCompilation)c.SemanticModel.Compilation).LanguageVersion >= LanguageVersion.CSharp10
                ? ((INamedTypeSymbol)c.TargetSymbol, c.Attributes.Select(ad => ((INamedTypeSymbol?)ad.ConstructorArguments[0].Value, (INamedTypeSymbol?)ad.ConstructorArguments[1].Value)).ToArray())
                : ((INamedTypeSymbol, (INamedTypeSymbol?, INamedTypeSymbol?)[])?)null).Collect();

        context.RegisterSourceOutput(value, GenerateSource);
    }

    private static void GenerateSource(SourceProductionContext context, ImmutableArray<(INamedTypeSymbol AttributeSymbol, (INamedTypeSymbol? InterfaceType, INamedTypeSymbol? ClassType)[] Types)?> array) {
        if (array.All(e => e is null)) {
            context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0001", "NativeCOM", "Requires C# 10 or higher", DiagnosticSeverity.Error,
                DiagnosticSeverity.Error, true, 0));

            return;
        }

        if (array.Length > 1) {
            context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0002", "NativeCOM", "ClassFactoryAttribute attached to multiple types", DiagnosticSeverity.Error,
                DiagnosticSeverity.Error, true, 0));

            return;
        }

        var (targetType, typesArray) = array.Single()!.Value;

        if (!targetType.Interfaces.Any(i => i.ToDisplayString(FqnFormat) == "Bluehill.NativeCom.IClassFactory")) {
            context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0003", "NativeCOM", "Class does not implement IClassFactory", DiagnosticSeverity.Error,
                DiagnosticSeverity.Error, true, 0, location: targetType.Locations[0]));

            return;
        }

        if (!targetType.GetAttributes().Any(ad => ad.AttributeClass!.ToDisplayString(FqnFormat) == "System.Runtime.InteropServices.Marshalling.GeneratedComClassAttribute")) {
            context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0004", "NativeCOM", "Class does not have GeneratedComClassAttribute", DiagnosticSeverity.Error,
                DiagnosticSeverity.Error, true, 0, location: targetType.Locations[0]));

            return;
        }

        StringBuilder sb = new();

        if (typesArray.All(t => t.InterfaceType is not null && t.ClassType is not null)) {
            sb.Append("namespace ").Append(targetType.ContainingNamespace.ToDisplayString(FqnFormat)).AppendLine(";").AppendLine()
                .Append("partial class ").Append(targetType.Name).AppendLine(" {")
                .AppendLine("    private static volatile int serverLocks;").AppendLine()
                .AppendLine("    public static bool IsLocked => serverLocks > 0;").AppendLine()
                .AppendLine("    public unsafe int CreateInstance(void* pUnkOuter, Guid* riid, void** ppvObject) {")
                .Append("        ");

            foreach (var tuple in typesArray) {
                var (i, c) = tuple;

                sb.Append("if (*riid == typeof(").Append(i!.ToDisplayString(FqnFormat)).AppendLine(").GUID) {")
                    .Append("            return Bluehill.NativeCom.DllHelper.CreateInstanceHelper<").Append(i.ToDisplayString(FqnFormat)).Append(", ").Append(c!.ToDisplayString(FqnFormat)).AppendLine(">(pUnkOuter, riid, ppvObject);")
                    .Append("        } else ");
            }

            sb.AppendLine("{")
                .AppendLine("            return -2147221231;")
                .AppendLine("        }")
                .AppendLine("    }").AppendLine()
                .AppendLine("    public int LockServer(bool fLock) {")
                .AppendLine("        if (fLock) {")
                .AppendLine("            Interlocked.Increment(ref serverLocks);")
                .AppendLine("        } else {")
                .AppendLine("            Interlocked.Decrement(ref serverLocks);")
                .AppendLine("        }").AppendLine()
                .AppendLine("        return 0;")
                .AppendLine("    }")
                .AppendLine("}");

            context.AddSource($"{targetType.ToDisplayString(FqnFormat)}.ClassFactory.g.cs", SourceText.From(sb.ToString(), Encoding.UTF8));

            context.AddSource("Dll.g.cs", SourceText.From($$"""
                                                            [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]
                                                            internal static unsafe class Dll {
                                                                [System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllGetClassObject))]
                                                                private static int DllGetClassObject(Guid* rclsid, Guid* riid, void** ppv) => Bluehill.NativeCom.DllHelper.GetClassFactory<{{targetType.ToDisplayString(FqnFormat)}}>(rclsid, riid, ppv);

                                                                [System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]
                                                                private static int DllCanUnloadNow() => {{targetType.ToDisplayString(FqnFormat)}}.IsLocked.GetHashCode();
                                                            }
                                                            """, Encoding.UTF8));
        }
    }
}
