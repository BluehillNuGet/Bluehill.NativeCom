# Bluehill.NativeCom
Build native in-process COM servers (DLLs) on .NET with Native AOT. This library + source generator removes most of the boilerplate for COM activation, class factories, and DLL exports, making it practical to ship shell extensions and other in-proc COM components written in C#.

> Windows-only. Requires .NET 8 or later and NativeAOT publishing.

## What’s inside
- `Bluehill.NativeCom.IClassFactory` — COM `IClassFactory` (`[GeneratedComInterface]`) with `CreateInstance` and `LockServer`.
- `Bluehill.NativeCom.ClassFactoryAttribute<T>` — annotate a factory class to wire it to a target COM class.
- `Bluehill.NativeCom.DllHelper` — helper that implements CreateInstance for your target class using `StrategyBasedComWrappers`.
- Source generator (`Bluehill.NativeCom.SourceGenerator`) that:
  - Validates your setup and reports diagnostics.
  - Generates partial implementation of your factory (method bodies for `CreateInstance` and `LockServer`).
  - Emits `UnmanagedCallersOnly` exports: `DllGetClassObject` and `DllCanUnloadNow`.

## Requirements
- Windows
- .NET 8+ (recommended TargetFramework: net8.0-windows)
- Native AOT publishing
- C# 12 (default in .NET 8)

## Installation
Add the following packages to your class library project:
- Bluehill.NativeCom
- Bluehill.NativeCom.SourceGenerator

You can add them via NuGet Package Manager or dotnet CLI.

## Quick start
1) Enable NativeAOT and unsafe code in your .csproj:
    ```xml
    <PropertyGroup>
      <TargetFramework>net8.0-windows</TargetFramework>
      <PublishAot>true</PublishAot>
      <AllowUnsafeBlocks>true</AllowUnsafeBlocks>
    </PropertyGroup>
    ```

2) Define COM interface(s) using the new source-generated COM attributes:
    ```csharp
    using System;
    using System.Runtime.InteropServices;
    using System.Runtime.InteropServices.Marshalling;
    
    [GeneratedComInterface]
    [Guid("a08ce4d0-fa25-44ab-b57c-c7b1c323e0b9")]
    [InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
    public partial interface IExplorerCommand {
    {
        ...
    }
    ```

3) Implement the COM class that provides the behavior. It must be annotated with `[GeneratedComClass]` and have a CLSID via `[Guid]`:
    ```csharp
    using System;
    using System.Runtime.InteropServices; // for [Guid]
    using System.Runtime.InteropServices.Marshalling;
    
    [GeneratedComClass]
    [Guid("22222222-2222-2222-2222-222222222222")] // CLSID used to activate this class
    public partial class ExplorerCommand : IExplorerCommand
    {
        ...
    }
    ```

4) Declare a factory class and connect it to your target via `[ClassFactory<TTarget>]`. Do not implement the methods yourself — the generator emits the method bodies.
    ```csharp
    using Bluehill.NativeCom;
    using System.Runtime.InteropServices.Marshalling;
    
    [GeneratedComClass]
    [ClassFactory<ExplorerCommand>]
    public partial class ExplorerCommandFactory : IClassFactory
    {
        // The generator provides CreateInstance(...) and LockServer(...).
    }
    ```

5) Publish with NativeAOT:
    ```powershell
    # From your project directory
    dotnet publish -c Release
    ```

The native DLL (your COM server) will be in the publish\native folder.

## How it works
- The source generator scans for classes annotated with `[ClassFactory<TTarget>]`. It validates that:
  - The factory class implements `Bluehill.NativeCom.IClassFactory`.
  - Both the factory and the target class have `[GeneratedComClass]`.
  - The target class has a `[Guid]` attribute (its CLSID).
- For each valid pair, the generator produces:
  - A partial factory implementation that forwards CreateInstance to `DllHelper.CreateInstanceHelper<TTarget>()`.
  - A partial LockServer method that do nothing.
- It also generates an internal static Dll class with `UnmanagedCallersOnly` exports:
  - `DllGetClassObject`: routes the requested CLSID to the correct factory `CreateInstance`.
  - `DllCanUnloadNow`: returns `S_FALSE` (1), since .NET Native AOT DLLs cannot be unloaded.

### Opting out of DLL exports
If you need to provide your own exports implementation, define the symbol `NO_GENERATE_DLLGCO` in your project. Example:
```xml
<PropertyGroup>
  <DefineConstants>$(DefineConstants);NO_GENERATE_DLLGCO</DefineConstants>
</PropertyGroup>
```
With this symbol defined, the generator will NOT emit `DllGetClassObject`.

## Diagnostics
The generator reports clear diagnostics to help you fix setup issues:
- `NATIVECOM0001`: Required type '{0}' couldn't be found.
- `NATIVECOM0002`: Factory class does not implement IClassFactory.
- `NATIVECOM0003`: Factory class does not have GeneratedComClassAttribute.
- `NATIVECOM0004`: Target class does not have GeneratedComClassAttribute.
- `NATIVECOM0005`: Target class does not have GuidAttribute.

## COM registration (in-proc)
This package generates `DllGetClassObject` and `DllCanUnloadNow`, but it does NOT generate `DllRegisterServer`/`DllUnregisterServer`. You are responsible for registration (e.g., via an installer, script, or custom code). At minimum, in-proc registration requires keys like:
- HKCR\CLSID\{Your-Target-Class-CLSID}\ (default: your class name or description)
- HKCR\CLSID\{Your-Target-Class-CLSID}\InprocServer32\ (default: full path to your compiled DLL, ThreadingModel=Both)
Ensure you register the DLL for the correct bitness (x64 vs x86) and that your process matches. Shell extensions typically require additional shell-specific registration.

## Notes & limitations
- In-process servers only (DLL). No out-of-proc server support.
- Aggregation is not supported (`CreateInstance` rejects non-null pUnkOuter).
- NativeAOT is required to export `UnmanagedCallersOnly` entry points.
- This library relies on .NET source-generated COM (`GeneratedComInterface`/`GeneratedComClass`) available in .NET 8+.

## License
MIT License — see LICENSE.
