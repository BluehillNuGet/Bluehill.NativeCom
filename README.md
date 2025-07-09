# Bluehill.NativeCom
A package to help create native in-process COM servers (DLLs) using native AOT. This can be used to create in-process COM servers, such as shell extensions.

## What's included?
* `Bluehill.NativeCom.ClassFactoryAttribute`: Attribute indicating that a class implements `IClassFactory`.
* `Bluehill.NativeCom.DllHelper`: A class that helps implement `DllGetClassObject` and `IClassFactory`.
* `Bluehill.NativeCom.IClassFactory`: The COM `IClassFactory` interface.
* Source generator that automatically generates class factory class, `DllGetClassObject` and `DllCanUnloadNow`

## How to use
1. Create a class library project targeting .NET 8 or higher.
1. Make sure you are using C# 10 or higher. .NET 8 and higher use C# 12 and higher by default.
1. Enable native AOT publishing and unsafe code by modifying your project file as follows:
    ```
    <PropertyGroup>
      <PublishAot>true</PublishAot>
      <AllowUnsafeBlocks>true</AllowUnsafeBlocks>
    </PropertyGroup>
    ```
1. Write the COM interfaces you want and the classes that implement them. Don't forget to use `[GeneratedComInterface]` and `[GeneratedComClass]` instead of `[ComImport]` and `[ComVisible(true)]`.
1. Automatically generate a factory and `DllGetClassObject` and `DllCanUnloadNow` using the `ClassFactory` attribute as follows:
    ```
    [GeneratedComClass]
    [Guid("GUID")]
    [ClassFactory(typeof(IInterface), typeof(Implementation))]
    [ClassFactory(typeof(IOtherInterface), typeof(OtherImplementation))]
    public partial class MyClassFactory : IClassFactory;
    ```
1. Run `dotnet publish`. You will find the COM server DLL in the `native` folder of your publish folder!
