# ComAutoWrapper

Simple AutoWrap-style COM method/property invoker for .NET (C#).  
Useful for Excel, Word, and other COM automation tasks — without Interop DLLs or `dynamic`.

## Introspect COM members

You can list all callable members (methods and properties) of any IDispatch-based COM object:

```csharp
var (methods, propsGet, propsSet) = ComTypeInspector.ListMembers(comObject);

Console.WriteLine("Methods:");
methods.ForEach(Console.WriteLine);

Console.WriteLine("PropertyGet:");
propsGet.ForEach(Console.WriteLine);

Console.WriteLine("PropertySet:");
propsSet.ForEach(Console.WriteLine);

You can also get the COM type name:

var typeName = ComTypeInspector.GetTypeName(comObject);
Console.WriteLine($"COM type: {typeName}");


## Install

```bash
dotnet add package ComAutoWrapper --version 1.1.0

Example (Excel Automation):
var excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);
ComInvoker.SetProperty(excel, "Visible", true);
var workbooks = ComInvoker.GetProperty<object>(excel, "Workbooks");
ComInvoker.CallMethod(workbooks, "Add");
ComInvoker.CallMethod(excel, "Quit");

Features:
Typed GetProperty<T>()
Safe SetProperty(...)
CallMethod(...) with params
COM HRESULT error highlighting

No Interop DLL dependency

Works on .NET 6/7/8/9

License:
MIT