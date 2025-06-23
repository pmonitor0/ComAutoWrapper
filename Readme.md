# ComAutoWrapper

Simple AutoWrap-style COM method/property invoker for .NET (C#).  
Useful for Excel, Word, and other COM automation tasks — without Interop DLLs or `dynamic`.

## Install

```bash
dotnet add package ComAutoWrapper --version 1.0.0

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