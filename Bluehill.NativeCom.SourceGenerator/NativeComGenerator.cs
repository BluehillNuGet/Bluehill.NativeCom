using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;
using System.Collections.Immutable;
using System.Text;

namespace Bluehill.NativeCom.SourceGenerator;

[Generator(LanguageNames.CSharp)]
public sealed class NativeComGenerator : IIncrementalGenerator {
    private static readonly SymbolDisplayFormat FqnFormat = new(
        typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces,
        memberOptions: SymbolDisplayMemberOptions.IncludeContainingType);

    private static readonly DiagnosticDescriptor NativeCom0001 = new("NATIVECOM0001", "Couldn't be found required type", "'{0}' couldn't be found",
        "NativeCOM", DiagnosticSeverity.Error, true);

    public void Initialize(IncrementalGeneratorInitializationContext context) {
        var value = context.SyntaxProvider.ForAttributeWithMetadataName("Bluehill.NativeCom.ClassFactoryAttribute`1",
                static (n, _) => n is ClassDeclarationSyntax,
                static (c, _) => ((INamedTypeSymbol)c.TargetSymbol,
                                  (INamedTypeSymbol)c.Attributes.Select(static ad => ad.AttributeClass!.TypeArguments.Single()).Single()))
            .Combine(context.CompilationProvider).Collect();

        context.RegisterSourceOutput(value, GenerateSource);
    }

    private static void GenerateSource(SourceProductionContext context, ImmutableArray<((INamedTypeSymbol, INamedTypeSymbol), Compilation)> array) {
        StringBuilder clsidSb = new();
        StringBuilder helperSb = new();
        INamedTypeSymbol? icfSymbol = null;
        INamedTypeSymbol? gccaSymbol = null;
        INamedTypeSymbol? gaaSymbol = null;

        for (var index = 0; index < array.Length; index++) {
            var (innerTuple, compilation) = array[index];
            var (factory, target) = innerTuple;

            const string icf = "Bluehill.NativeCom.IClassFactory";
            icfSymbol ??= compilation.GetTypeByMetadataName(icf);

            if (icfSymbol is null) {
                context.ReportDiagnostic(Diagnostic.Create(NativeCom0001, factory.Locations[0], icf));

                return;
            }

            if (!factory.Interfaces.Any(i => SymbolEqualityComparer.Default.Equals(i, icfSymbol))) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0002", "NativeCOM", "Factory class does not implement IClassFactory",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            const string gcca = "System.Runtime.InteropServices.Marshalling.GeneratedComClassAttribute";
            gccaSymbol ??= compilation.GetTypeByMetadataName(gcca);

            if (gccaSymbol is null) {
                context.ReportDiagnostic(Diagnostic.Create(NativeCom0001, factory.Locations[0], gcca));

                return;
            }

            if (!factory.GetAttributes().Any(ad => SymbolEqualityComparer.Default.Equals(ad.AttributeClass, gccaSymbol))) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0003", "NativeCOM", "Factory class does not have GeneratedComClassAttribute",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            var ta = target.GetAttributes();

            if (!ta.Any(ad => SymbolEqualityComparer.Default.Equals(ad.AttributeClass, gccaSymbol))) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0004", "NativeCOM", "Target class does not have GeneratedComClassAttribute",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            const string gaa = "System.Runtime.InteropServices.GuidAttribute";
            gaaSymbol ??= compilation.GetTypeByMetadataName(gaa);

            if (gaaSymbol is null) {
                context.ReportDiagnostic(Diagnostic.Create(NativeCom0001, factory.Locations[0], gaa));

                return;
            }

            var ga = ta.SingleOrDefault(ad => SymbolEqualityComparer.Default.Equals(ad.AttributeClass, gaaSymbol));

            if (ga is null) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0005", "NativeCOM", "Target class does not have GuidAttribute",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            var guidArray = ga.ConstructorArguments.Single().Value!
                .ToString().Split('-');

            var guid4 = guidArray[0];
            var guid6 = guidArray[1];
            var guid8 = guidArray[2];
            var guid9 = guidArray[3].Substring(0, 2);
            var guid10 = guidArray[3].Substring(2, 2);
            var guid11 = guidArray[4].Substring(0, 2);
            var guid12 = guidArray[4].Substring(2, 2);
            var guid13 = guidArray[4].Substring(4, 2);
            var guid14 = guidArray[4].Substring(6, 2);
            var guid15 = guidArray[4].Substring(8, 2);
            var guid16 = guidArray[4].Substring(10, 2);

            const string commaZeroX = ", 0x";

            clsidSb.Append("        { new(0x").Append(guid4).Append(commaZeroX).Append(guid6).Append(commaZeroX).Append(guid8).Append(commaZeroX)
                .Append(guid9).Append(commaZeroX).Append(guid10).Append(commaZeroX).Append(guid11).Append(commaZeroX).Append(guid12)
                .Append(commaZeroX).Append(guid13).Append(commaZeroX).Append(guid14).Append(commaZeroX).Append(guid15).Append(commaZeroX)
                .Append(guid16).Append("), ").Append(index).AppendLine(" },");

            helperSb.Append("        &global::Bluehill.NativeCom.DllHelper.CreateInstanceHelper<global::").Append(factory.ToDisplayString(FqnFormat))
                .AppendLine(">,");

            StringBuilder cfSb = new();

            cfSb.Append("namespace ").Append(factory.ContainingNamespace.ToDisplayString(FqnFormat)).AppendLine(";").AppendLine()
                .Append("partial class ").Append(factory.Name).AppendLine(" {")
                .AppendLine("    public unsafe int CreateInstance(void* pUnkOuter, global::System.Guid* riid, void** ppvObject)")
                .Append("        => global::Bluehill.NativeCom.DllHelper.CreateInstanceHelper<global::").Append(target.ToDisplayString(FqnFormat))
                .AppendLine(">(pUnkOuter, riid, ppvObject);").AppendLine()
                .AppendLine("#if !NO_GENERATE_DLL_ENTRYPOINT_AND_LOCKSERVER")
                .AppendLine("    public int LockServer(bool fLock) {")
                .AppendLine("        if (fLock) {")
                .AppendLine("            global::System.Threading.Interlocked.Increment(ref Dll.Locks);")
                .AppendLine("        } else {")
                .AppendLine("            global::System.Threading.Interlocked.Decrement(ref Dll.Locks);")
                .AppendLine("        }").AppendLine()
                .AppendLine("        return 0;")
                .AppendLine("    }")
                .AppendLine("#endif")
                .AppendLine("}");

            context.AddSource($"{factory.ToDisplayString(FqnFormat)}.BHNC.g.cs", SourceText.From(cfSb.ToString(), Encoding.UTF8));
        }

        StringBuilder dllSb = new();

        const string editorBrowsableNever
            = "[global::System.ComponentModel.EditorBrowsable(global::System.ComponentModel.EditorBrowsableState.Never)]";

        dllSb.AppendLine("#if !NO_GENERATE_DLL_ENTRYPOINT_AND_LOCKSERVER")
            .AppendLine(editorBrowsableNever)
            .AppendLine("internal static unsafe class Dll {")
            .AppendLine($"    {editorBrowsableNever}")
            .AppendLine("    public static volatile int Locks;").AppendLine()
            .AppendLine("    private static readonly global::System.Collections.Generic.Dictionary<global::System.Guid, int> Clsids = new() {")
            .Append(clsidSb)
            .AppendLine("    };").AppendLine()
            .AppendLine("    private static readonly delegate*<void*, global::System.Guid*, void**, int>[] Helpers = [")
            .Append(helperSb)
            .AppendLine("    ];").AppendLine()
            .AppendLine("    [global::System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllGetClassObject))]")
            .AppendLine("    private static int DllGetClassObject(global::System.Guid* rclsid, global::System.Guid* riid, void** ppv)")
            .AppendLine("        => Clsids.TryGetValue(*rclsid, out var index) ? Helpers[index](null, riid, ppv) : -2147221231;").AppendLine()
            .AppendLine("    [global::System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]")
            .AppendLine("    private static int DllCanUnloadNow() => Locks <= 0 ? 0 : 1;")
            .AppendLine("};")
            .AppendLine("#endif");

        context.AddSource("Dll.g.cs", SourceText.From(dllSb.ToString(), Encoding.UTF8));
    }
}
