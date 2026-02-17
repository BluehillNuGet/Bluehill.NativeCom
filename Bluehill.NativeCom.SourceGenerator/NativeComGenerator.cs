using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;
using System.Collections.Immutable;
using System.Text;

namespace Bluehill.NativeCom.SourceGenerator;

/// <summary>
/// Incremental source generator that discovers factory classes marked with ClassFactoryAttribute&lt;T&gt;
/// and produces COM activation glue:
/// - Validates that the factory implements IClassFactory and is annotated with GeneratedComClassAttribute.
/// - Validates that the target type is annotated with GeneratedComClassAttribute and GuidAttribute.
/// - Emits a partial factory implementation delegating to DllHelper and, optionally, DLL entry points.
/// </summary>
[Generator(LanguageNames.CSharp)]
public sealed class NativeComGenerator : IIncrementalGenerator {
    // Fully-qualified name format (namespace + containing types) for embedding type names into generated code.
    private static readonly SymbolDisplayFormat FqnFormat = new(
        typeQualificationStyle: SymbolDisplayTypeQualificationStyle.NameAndContainingTypesAndNamespaces,
        memberOptions: SymbolDisplayMemberOptions.IncludeContainingType);

    // Diagnostic emitted when a required type cannot be found by metadata name in the current compilation.
    private static readonly DiagnosticDescriptor NativeCom0001 = new("NATIVECOM0001", "Couldn't be found required type", "'{0}' couldn't be found",
        "NativeCOM", DiagnosticSeverity.Error, true);

    /// <summary>
    /// Configures the incremental pipeline that tracks classes with ClassFactoryAttribute&lt;T&gt; and the Compilation.
    /// </summary>
    /// <param name="context">The context for the incremental generator.</param>
    public void Initialize(IncrementalGeneratorInitializationContext context) {
        // Query class declarations annotated with Bluehill.NativeCom.ClassFactoryAttribute<T>.
        // We capture (factory type, target type) by reading the attribute's generic argument.
        var value = context.SyntaxProvider.ForAttributeWithMetadataName("Bluehill.NativeCom.ClassFactoryAttribute`1",
                static (n, _) => n is ClassDeclarationSyntax,
                static (c, _) => ((INamedTypeSymbol)c.TargetSymbol,
                                  (INamedTypeSymbol)c.Attributes.Select(static ad => ad.AttributeClass!.TypeArguments.Single()).Single()))
            .Collect();

        // Register source output callback invoked when tracked inputs change.
        // Pair with the Compilation to resolve well-known types by metadata name during generation.
        context.RegisterSourceOutput(context.CompilationProvider.Combine(value), GenerateSource);
    }

    /// <summary>
    /// Generates partial factory implementations and an optional hidden Dll helper with COM exports.
    /// </summary>
    private static void GenerateSource(SourceProductionContext context, (Compilation, ImmutableArray<(INamedTypeSymbol, INamedTypeSymbol)>) tuple) {
        var (compilation, array) = tuple;

        // Resolve IClassFactory by metadata name once per-compilation.
        const string icf = "Bluehill.NativeCom.IClassFactory";
        var icfSymbol = compilation.GetTypeByMetadataName(icf);

        if (icfSymbol is null) {
            // Cannot proceed without IClassFactory.
            context.ReportDiagnostic(Diagnostic.Create(NativeCom0001, null, icf));

            return;
        }

        // Resolve GeneratedComClassAttribute for validation of both factory and target types.
        const string gcca = "System.Runtime.InteropServices.Marshalling.GeneratedComClassAttribute";
        var gccaSymbol = compilation.GetTypeByMetadataName(gcca);

        if (gccaSymbol is null) {
            context.ReportDiagnostic(Diagnostic.Create(NativeCom0001, null, gcca));

            return;
        }

        // Resolve GuidAttribute to extract the target's CLSID fields.
        const string gaa = "System.Runtime.InteropServices.GuidAttribute";
        var gaaSymbol = compilation.GetTypeByMetadataName(gaa); // System.Runtime.InteropServices.GuidAttribute

        if (gaaSymbol is null) {
            context.ReportDiagnostic(Diagnostic.Create(NativeCom0001, null, gaa));

            return;
        }

        // Accumulators for generated initializer code:
        // - clsidSb: lines mapping CLSIDs to helper indices,
        // - helperSb: function pointers aligned to those indices.
        StringBuilder clsidSb = new();
        StringBuilder helperSb = new();

        for (var index = 0; index < array.Length; index++) {
            var (factory, target) = array[index];

            // Ensure the factory type implements IClassFactory.
            if (!factory.Interfaces.Any(i => SymbolEqualityComparer.Default.Equals(i, icfSymbol))) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0002", "NativeCOM", "Factory class does not implement IClassFactory",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            // Factory must be annotated with GeneratedComClassAttribute.
            if (!factory.GetAttributes().Any(ad => SymbolEqualityComparer.Default.Equals(ad.AttributeClass, gccaSymbol))) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0003", "NativeCOM", "Factory class does not have GeneratedComClassAttribute",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            var ta = target.GetAttributes();

            // Target must also be annotated with GeneratedComClassAttribute.
            if (!ta.Any(ad => SymbolEqualityComparer.Default.Equals(ad.AttributeClass, gccaSymbol))) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0004", "NativeCOM", "Target class does not have GeneratedComClassAttribute",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            // Fetch GuidAttribute from the target type (required for CLSID mapping).
            var ga = ta.SingleOrDefault(ad => SymbolEqualityComparer.Default.Equals(ad.AttributeClass, gaaSymbol));

            if (ga is null) {
                context.ReportDiagnostic(Diagnostic.Create("NATIVECOM0005", "NativeCOM", "Target class does not have GuidAttribute",
                    DiagnosticSeverity.Error, DiagnosticSeverity.Error, true, 0, location: factory.Locations[0]));

                return;
            }

            // Split the canonical GUID string to build Guid literal: new(0xA, 0xB, 0xC, 0xD0, 0xD1, ..., 0xD7).
            var guidArray = ga.ConstructorArguments.Single().Value!
                .ToString().Split('-');

            // Field slices correspond to {a}-{b}-{c}-{d0d1}-{d2d3d4d5d6d7}.
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

            // Add a CLSID -> helper index mapping entry.
            clsidSb.Append("        { new(0x").Append(guid4).Append(commaZeroX).Append(guid6).Append(commaZeroX).Append(guid8).Append(commaZeroX)
                .Append(guid9).Append(commaZeroX).Append(guid10).Append(commaZeroX).Append(guid11).Append(commaZeroX).Append(guid12)
                .Append(commaZeroX).Append(guid13).Append(commaZeroX).Append(guid14).Append(commaZeroX).Append(guid15).Append(commaZeroX)
                .Append(guid16).Append("), ").Append(index).AppendLine(" },");

            // Add a function pointer to the CreateInstance helper for this factory type.
            helperSb.Append("        &global::Bluehill.NativeCom.DllHelper.CreateInstanceHelper<global::").Append(factory.ToDisplayString(FqnFormat))
                .AppendLine(">,");

            // Generate a partial class for the factory that forwards calls to DllHelper<TTarget>.
            StringBuilder cfSb = new();

            cfSb.Append("namespace ").Append(factory.ContainingNamespace.ToDisplayString(FqnFormat)).AppendLine(";").AppendLine()
                .Append("partial class ").Append(factory.Name).AppendLine(" {")
                .AppendLine("    public unsafe int CreateInstance(void* pUnkOuter, global::System.Guid* riid, void** ppvObject)")
                .Append("        => global::Bluehill.NativeCom.DllHelper.CreateInstanceHelper<global::").Append(target.ToDisplayString(FqnFormat))
                .AppendLine(">(pUnkOuter, riid, ppvObject);").AppendLine()
                .AppendLine("    public int LockServer(bool fLock) => 0;")
                .AppendLine("}");

            // Emit the partial factory source for the current factory type.
            context.AddSource($"{factory.ToDisplayString(FqnFormat)}.BHNC.g.cs", SourceText.From(cfSb.ToString(), Encoding.UTF8));
        }

        // Build the Dll static class that exposes COM exports and holds activation helpers.
        StringBuilder dllSb = new();

        // Hide generated members from IntelliSense to keep user projects tidy.
        const string editorBrowsableNever
            = "[global::System.ComponentModel.EditorBrowsable(global::System.ComponentModel.EditorBrowsableState.Never)]";

        // Guard generation with NO_GENERATE_DLLGCO so consumers can opt out of DLL exports and LockServer.
        dllSb.AppendLine(editorBrowsableNever)
            .AppendLine("internal static unsafe class Dll {")
            .AppendLine("#if !NO_GENERATE_DLLGCO")
            // Maps CLSID to index into Helpers[].
            .AppendLine("    private static readonly global::System.Collections.Generic.Dictionary<global::System.Guid, int> Clsids = new() {")
            .Append(clsidSb)
            .AppendLine("    };").AppendLine()
            // Function pointer table to CreateInstance helpers; index must match entries in Clsids.
            .AppendLine("    private static readonly delegate*<void*, global::System.Guid*, void**, int>[] Helpers = [")
            .Append(helperSb)
            .AppendLine("    ];").AppendLine()
            // Classic COM export for class factory resolution.
            .AppendLine("    [global::System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllGetClassObject))]")
            .AppendLine("    private static int DllGetClassObject(global::System.Guid* rclsid, global::System.Guid* riid, void** ppv)")
            // If CLSID not found, return CLASS_E_CLASSNOTAVAILABLE (0x80040111 -> -2147221231).
            .AppendLine("        => Clsids.TryGetValue(*rclsid, out var index) ? Helpers[index](null, riid, ppv) : -2147221231;").AppendLine()
            .AppendLine("#endif")
            // Classic COM export indicating whether the DLL can be unloaded (S_OK if no locks, S_FALSE otherwise).
            .AppendLine("    [global::System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]")
            .AppendLine("    private static int DllCanUnloadNow() => 1;")
            .AppendLine("};");

        // Emit the Dll helper source.
        context.AddSource("Dll.g.cs", SourceText.From(dllSb.ToString(), Encoding.UTF8));
    }
}
