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

    public void Initialize(IncrementalGeneratorInitializationContext context) {
        var value = context.SyntaxProvider.ForAttributeWithMetadataName("Bluehill.NativeCom.ClassFactoryAttribute`1",
            static (n, _) => n is ClassDeclarationSyntax,
            static (c, _) => ((INamedTypeSymbol)c.TargetSymbol,
                              (INamedTypeSymbol)c.Attributes.Select(static ad => ad.AttributeClass!.TypeArguments.Single()).Single())).Collect();

        context.RegisterSourceOutput(value, GenerateSource);
    }

    private static void GenerateSource(SourceProductionContext context, ImmutableArray<(INamedTypeSymbol, INamedTypeSymbol)> array) {
        StringBuilder dllSb = new();

        dllSb.AppendLine("#if !NO_GENERATE_DLL")
            .AppendLine("using Bluehill.NativeCom;").AppendLine()
            .AppendLine("[System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]")
            .AppendLine("internal static unsafe class Dll {")
            .AppendLine("    [System.ComponentModel.EditorBrowsable(System.ComponentModel.EditorBrowsableState.Never)]")
            .AppendLine("    public static volatile int Locks;").AppendLine();

        StringBuilder clsidSb = new();
        StringBuilder helperSb = new();

        for (var i = 0; i < array.Length; i++) {
            var (factory, target) = array[i];

            if (!factory.Interfaces.Any(static ad => ad.ToDisplayString(FqnFormat) == "Bluehill.NativeCom.IClassFactory")) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0001", "NativeCOM", "Class does not implement IClassFactory",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            if (!factory.GetAttributes().Any(static ad
                    => ad.AttributeClass?.ToDisplayString(FqnFormat) == "System.Runtime.InteropServices.Marshalling.GeneratedComClassAttribute")) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0002", "NativeCOM", "Factory class does not have GeneratedComClassAttribute",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            var ta = target.GetAttributes();

            if (!ta.Any(static ad
                    => ad.AttributeClass?.ToDisplayString(FqnFormat) == "System.Runtime.InteropServices.Marshalling.GeneratedComClassAttribute")) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0003", "NativeCOM", "Target class does not have GeneratedComClassAttribute",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            var guidArray = ta.Single(ad => ad.AttributeClass?.ToDisplayString(FqnFormat) == "System.Runtime.InteropServices.GuidAttribute")
                .ConstructorArguments.Single().Value!.ToString().Split('-');

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
                .Append(guid16).Append("), ").Append(i).AppendLine(" },");

            helperSb.Append("        &DllHelper.CreateInstanceHelper<").Append(factory.ToDisplayString(FqnFormat)).AppendLine(">,");

            StringBuilder cfSb = new();

            cfSb.Append("namespace ").Append(factory.ContainingNamespace.ToDisplayString(FqnFormat)).AppendLine(";").AppendLine()
                .Append("partial class ").Append(factory.Name).AppendLine(" {")
                .AppendLine("    public unsafe int CreateInstance(void* pUnkOuter, Guid* riid, void** ppvObject)")
                .Append("        => Bluehill.NativeCom.DllHelper.CreateInstanceHelper<").Append(target.ToDisplayString(FqnFormat))
                .AppendLine(">(pUnkOuter, riid, ppvObject);").AppendLine()
                .AppendLine("    public int LockServer(bool fLock) {")
                .AppendLine("        if (fLock) {")
                .AppendLine("            Interlocked.Increment(ref Dll.Locks);")
                .AppendLine("        } else {")
                .AppendLine("            Interlocked.Decrement(ref Dll.Locks);")
                .AppendLine("        }").AppendLine()
                .AppendLine("        return 0;")
                .AppendLine("    }")
                .AppendLine("}");

            context.AddSource($"{factory.ToDisplayString(FqnFormat)}.BHNC.g.cs", SourceText.From(cfSb.ToString(), Encoding.UTF8));
        }

        dllSb.AppendLine("    private static readonly Dictionary<Guid, int> Clsids = new() {")
            .Append(clsidSb)
            .AppendLine("    };").AppendLine()
            .AppendLine("    private static readonly delegate*<void*, Guid*, void**, int>[] Helpers = [")
            .Append(helperSb)
            .AppendLine("    ];").AppendLine()
            .AppendLine("    [System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllGetClassObject))]")
            .AppendLine("    private static int DllGetClassObject(Guid* rclsid, Guid* riid, void** ppv)")
            .AppendLine("        => Clsids.TryGetValue(*rclsid, out var index) ? Helpers[index](null, riid, ppv) : -2147221231;").AppendLine()
            .AppendLine("    [System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]")
            .AppendLine("    private static int DllCanUnloadNow() => Locks <= 0 ? 0 : 1;")
            .AppendLine("};")
            .AppendLine("#endif");

        context.AddSource("Dll.g.cs", SourceText.From(dllSb.ToString(), Encoding.UTF8));
    }
}
