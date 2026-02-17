using Microsoft.CodeAnalysis.Testing;

namespace Bluehill.NativeCom.SourceGenerator.Tests;

public sealed class NativeComGeneratorTests {
    [Fact]
    public async Task ValidSetup_GeneratesSource() {
        const string source = """
                              using Bluehill.NativeCom;
                              using System;
                              using System.Runtime.InteropServices;
                              using System.Runtime.InteropServices.Marshalling;

                              namespace TestNamespace;

                              [GeneratedComInterface]
                              [Guid("00000002-0000-0000-C000-000000000046")]
                              public partial interface IExplorerCommand { }

                              [GeneratedComClass]
                              [Guid("E10F1111-2222-3333-4444-555566667777")]
                              public partial class ExplorerCommand : IExplorerCommand { }

                              [GeneratedComClass]
                              [ClassFactory<ExplorerCommand>]
                              public partial class ExplorerCommandFactory : IClassFactory { }

                              """;

        const string expectedFactorySource = """
                                             namespace TestNamespace;

                                             partial class ExplorerCommandFactory {
                                                 public unsafe int CreateInstance(void* pUnkOuter, global::System.Guid* riid, void** ppvObject)
                                                     => global::Bluehill.NativeCom.DllHelper.CreateInstanceHelper<global::TestNamespace.ExplorerCommand>(pUnkOuter, riid, ppvObject);

                                                 public int LockServer(bool fLock) => 0;
                                             }

                                             """;

        const string expectedDllSource = """
                                         [global::System.ComponentModel.EditorBrowsable(global::System.ComponentModel.EditorBrowsableState.Never)]
                                         internal static unsafe class Dll {
                                         #if !NO_GENERATE_DLLGCO
                                             private static readonly global::System.Collections.Generic.Dictionary<global::System.Guid, int> Clsids = new() {
                                                 { new(0xE10F1111, 0x2222, 0x3333, 0x44, 0x44, 0x55, 0x55, 0x66, 0x66, 0x77, 0x77), 0 },
                                             };

                                             private static readonly delegate*<void*, global::System.Guid*, void**, int>[] Helpers = [
                                                 &global::Bluehill.NativeCom.DllHelper.CreateInstanceHelper<global::TestNamespace.ExplorerCommandFactory>,
                                             ];

                                             [global::System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllGetClassObject))]
                                             private static int DllGetClassObject(global::System.Guid* rclsid, global::System.Guid* riid, void** ppv)
                                                 => Clsids.TryGetValue(*rclsid, out var index) ? Helpers[index](null, riid, ppv) : -2147221231;

                                         #endif
                                             [global::System.Runtime.InteropServices.UnmanagedCallersOnly(EntryPoint = nameof(DllCanUnloadNow))]
                                             private static int DllCanUnloadNow() => 1;
                                         };

                                         """;

        CSharpSourceGeneratorVerifier<NativeComGenerator>.Test test = new() {
            TestState = {
                Sources = {
                    source
                },
                GeneratedSources = {
                    (typeof(NativeComGenerator), "TestNamespace.ExplorerCommandFactory.BHNC.g.cs", expectedFactorySource),
                    (typeof(NativeComGenerator), "Dll.g.cs", expectedDllSource),
                },
            },
        };

        test.TestState.AdditionalReferences.Add(typeof(IClassFactory).Assembly);

        await test.RunAsync(TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task MissingIClassFactory_ReportsDiagnostic() {
        const string source = """
                              using Bluehill.NativeCom;
                              using System.Runtime.InteropServices;
                              using System.Runtime.InteropServices.Marshalling;

                              namespace TestNamespace;

                              [GeneratedComClass]
                              [Guid("E10F1111-2222-3333-4444-555566667777")]
                              public partial class ExplorerCommand { }

                              [GeneratedComClass]
                              [ClassFactory<ExplorerCommand>]
                              public partial class ExplorerCommandFactory { }

                              """;

        CSharpSourceGeneratorVerifier<NativeComGenerator>.Test test = new() {
            TestState = {
                Sources = {
                    source
                },
                ExpectedDiagnostics = {
                    new DiagnosticResult("NATIVECOM0002", Microsoft.CodeAnalysis.DiagnosticSeverity.Error).WithSpan(13, 22, 13, 44)
                        .WithMessage("Factory class does not implement IClassFactory"),
                },
            },
        };

        test.TestState.AdditionalReferences.Add(typeof(IClassFactory).Assembly);

        await test.RunAsync(TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task MissingGuidAttribute_ReportsDiagnostic() {
        const string source = """
                              using Bluehill.NativeCom;
                              using System.Runtime.InteropServices;
                              using System.Runtime.InteropServices.Marshalling;

                              namespace TestNamespace;

                              [GeneratedComClass]
                              public partial class ExplorerCommand { }

                              [GeneratedComClass]
                              [ClassFactory<ExplorerCommand>]
                              public partial class ExplorerCommandFactory : IClassFactory { }

                              """;

        CSharpSourceGeneratorVerifier<NativeComGenerator>.Test test = new() {
            TestState = {
                Sources = {
                    source
                },
                ExpectedDiagnostics = {
                    new DiagnosticResult("NATIVECOM0005", Microsoft.CodeAnalysis.DiagnosticSeverity.Error).WithSpan(12, 22, 12, 44)
                        .WithMessage("Target class does not have GuidAttribute"),
                    // 인터페이스 미구현 오류 무시 (생성기가 실패하여 코드를 생성하지 않으므로 발생)
                    DiagnosticResult.CompilerError("CS0535").WithSpan(12, 47, 12, 60).WithArguments("TestNamespace.ExplorerCommandFactory",
                        "Bluehill.NativeCom.IClassFactory.CreateInstance(void*, System.Guid*, void**)"),
                    DiagnosticResult.CompilerError("CS0535").WithSpan(12, 47, 12, 60).WithArguments("TestNamespace.ExplorerCommandFactory",
                        "Bluehill.NativeCom.IClassFactory.LockServer(bool)"),
                },
            },
        };

        test.TestState.AdditionalReferences.Add(typeof(IClassFactory).Assembly);

        await test.RunAsync(TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task TargetMissingGeneratedComClass_ReportsDiagnostic() {
        const string source = """
                              using Bluehill.NativeCom;
                              using System.Runtime.InteropServices;
                              using System.Runtime.InteropServices.Marshalling;

                              namespace TestNamespace;

                              [Guid("E10F1111-2222-3333-4444-555566667777")]
                              public partial class ExplorerCommand { }

                              [GeneratedComClass]
                              [ClassFactory<ExplorerCommand>]
                              public partial class ExplorerCommandFactory : IClassFactory { }

                              """;

        CSharpSourceGeneratorVerifier<NativeComGenerator>.Test test = new() {
            TestState = {
                Sources = {
                    source
                },
                ExpectedDiagnostics = {
                    new DiagnosticResult("NATIVECOM0004", Microsoft.CodeAnalysis.DiagnosticSeverity.Error).WithSpan(12, 22, 12, 44)
                        .WithMessage("Target class does not have GeneratedComClassAttribute"),
                    DiagnosticResult.CompilerError("CS0535").WithSpan(12, 47, 12, 60).WithArguments("TestNamespace.ExplorerCommandFactory",
                        "Bluehill.NativeCom.IClassFactory.CreateInstance(void*, System.Guid*, void**)"),
                    DiagnosticResult.CompilerError("CS0535").WithSpan(12, 47, 12, 60).WithArguments("TestNamespace.ExplorerCommandFactory",
                        "Bluehill.NativeCom.IClassFactory.LockServer(bool)"),
                },
            },
        };

        test.TestState.AdditionalReferences.Add(typeof(IClassFactory).Assembly);

        await test.RunAsync(TestContext.Current.CancellationToken);
    }

    [Fact]
    public async Task FactoryMissingGeneratedComClass_ReportsDiagnostic() {
        const string source = """
                              using Bluehill.NativeCom;
                              using System.Runtime.InteropServices;
                              using System.Runtime.InteropServices.Marshalling;

                              namespace TestNamespace;

                              [GeneratedComClass]
                              [Guid("E10F1111-2222-3333-4444-555566667777")]
                              public partial class ExplorerCommand { }

                              [ClassFactory<ExplorerCommand>]
                              public partial class ExplorerCommandFactory : IClassFactory { }

                              """;

        CSharpSourceGeneratorVerifier<NativeComGenerator>.Test test = new() {
            TestState = {
                Sources = {
                    source
                },
                ExpectedDiagnostics = {
                    new DiagnosticResult("NATIVECOM0003", Microsoft.CodeAnalysis.DiagnosticSeverity.Error).WithSpan(12, 22, 12, 44)
                        .WithMessage("Factory class does not have GeneratedComClassAttribute"),
                    DiagnosticResult.CompilerError("CS0535").WithSpan(12, 47, 12, 60).WithArguments("TestNamespace.ExplorerCommandFactory",
                        "Bluehill.NativeCom.IClassFactory.CreateInstance(void*, System.Guid*, void**)"),
                    DiagnosticResult.CompilerError("CS0535").WithSpan(12, 47, 12, 60).WithArguments("TestNamespace.ExplorerCommandFactory",
                        "Bluehill.NativeCom.IClassFactory.LockServer(bool)"),
                },
            },
        };

        test.TestState.AdditionalReferences.Add(typeof(IClassFactory).Assembly);

        await test.RunAsync(TestContext.Current.CancellationToken);
    }
}
