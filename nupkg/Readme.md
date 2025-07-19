# ComAutoWrapper

**ComAutoWrapper** is a lightweight, zero-Interop, fluent C# helper library for automating COM objects such as **Excel** and **Word** — without relying on bulky Primary Interop Assemblies (PIAs).

✔️ Fully dynamic  
✔️ Typed property/method access  
✔️ Introspectable  
✔️ Ideal for WPF / Console / WinForms projects  
✔️ Just **~30 KB** compiled DLL

---

## 🚀 Features

- **No Interop DLLs needed**
- Lightweight COM helper for C#
- Elegant dynamic wrappers:
  - `GetProperty<T>()`, `SetProperty()`
  - `CallMethod<T>()`
- COM introspection (`ComTypeInspector`)
- Excel selection utilities (`ComSelectionHelper`)
- Safe release of COM objects
- Compatible with: .NET 6, 7, 8, 9+

---

## 🧠 Examples

### Get/Set COM Properties

```csharp
var excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);
ComInvoker.SetProperty(excel, "Visible", true);

var workbooks = ComInvoker.GetProperty<object>(excel, "Workbooks");
var workbook = ComInvoker.CallMethod<object>(workbooks, "Add");
```

### Invoke COM Methods

```csharp
var sheet = ComInvoker.GetProperty<object>(workbook, "ActiveSheet");
var cell = ComInvoker.GetProperty<object>(sheet, "Cells");
ComInvoker.SetProperty(cell, "Item", new object[] { 1, 1 }, "Hello");
```

### Introspect COM Object

```csharp
var (methods, propsGet, propsSet) = ComTypeInspector.ListMembers(sheet);
Console.WriteLine("Available methods:");
methods.ForEach(Console.WriteLine);
```

---

## ✨ Excel-Specific Helpers (Optional)

Provided via the built-in `ComSelectionHelper`:

| Method | Description |
|--------|-------------|
| `SelectCells(excel, sheet, "A1", "B3", "C5")` | Selects non-contiguous Excel cells |
| `GetSelectedCellCoordinates(excel)` | Returns `(row, column)` for each selected cell |
| `HighlightUsedRange(sheet)` | Highlights the used range with color |

These helpers abstract away the quirks of Excel's COM object model.

---

## 📦 NuGet Package

Install via CLI:

```bash
dotnet add package ComAutoWrapper
```

Or via Visual Studio NuGet UI.

---

## 💻 Requirements

- Windows OS (COM-based)
- .NET 6 / 7 / 8 / 9
- Microsoft Excel/Word must be installed

> The library **does not embed Interop DLLs**. It uses late binding with proper error handling.

---

## 🔗 Related Project

- [ComAutoWrapperDemo (GitHub)](https://github.com/pmonitor0/ComAutoWrapperDemo)  
  WPF demo showcasing full Excel and Word automation using this wrapper.

---

## 📊 Comparison: OpenXML vs COM Automation

| Feature | OpenXML SDK | ComAutoWrapper |
|--------|-------------|----------------|
| Requires Excel Installed | ❌ | ✅ |
| Works on Locked/Password Files | ❌ | ✅ |
| Manipulate Active Excel Instance | ❌ | ✅ |
| Word Automation | ❌ | ✅ |
| File Size (DLL) | >10 MB | ~30 KB |
| API Simplicity | Moderate | High (fluent & dynamic) |
| Cell Selection / UI Interaction | ❌ | ✅ |
| UsedRange / Borders / Colors | ❌ | ✅ |

---

## 🙏 Acknowledgment

This library is the result of an iterative collaboration between the author and ChatGPT.  
Special thanks to all early testers and contributors who shaped the API.

---

## 📄 License

MIT