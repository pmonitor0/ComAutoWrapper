🚀 Quick Example (Excel automation)

```csharp
var excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);
ComAutoHelper.SetProperty(excel, "Visible", true);

var workbooks = ComAutoHelper.GetProperty<object>(excel, "Workbooks");
ComAutoHelper.CallMethod(workbooks, "Add");

ComAutoHelper.CallMethod(excel, "Quit"); ```

🔍 COM Member Introspection
You can list all callable members of any IDispatch COM object:

```csharp
var (methods, propsGet, propsSet) = ComTypeInspector.ListMembers(comObject);

methods.ForEach(m => Console.WriteLine("Method: " + m));
propsGet.ForEach(p => Console.WriteLine("PropertyGet: " + p));
propsSet.ForEach(p => Console.WriteLine("PropertySet: " + p)); ```


Get the COM type name:

```csharp
string typeName = ComTypeInspector.GetTypeName(comObject);
Console.WriteLine($"COM type: {typeName}"); ```


🧰 Property Access

Set properties
```csharp
ComAutoHelper.SetProperty(app, "DisplayAlerts", false);
ComAutoHelper.SetProperty(sheet, "Name", "Summary");
ComAutoHelper.SetProperty(rng, "Value", new object[,] { ... }); ```

Get properties (typed or untyped)
```csharp
bool visible = ComAutoHelper.GetProperty<bool>(excel, "Visible");
object sheets = ComAutoHelper.GetProperty<object>(workbook, "Sheets"); ```

⚙️ Method Invocation
With return type:
```csharp
int count = ComAutoHelper.CallMethod<int>(workbooks, "Count");
object sheet = ComAutoHelper.CallMethod<object>(sheets, "Item", 1); ```
Or generic/untyped:
```csharp

object result = ComAutoHelper.CallMethod(sheet, "Calculate"); ```

### 🔍 Check if a COM property exists

You can safely check if a property exists on a COM object:

```csharp
bool exists = ComAutoHelper.PropertyExists(excel, "DisplayAlerts");
if (exists)
    Console.WriteLine("Property exists."); ```


✅ TryGetProperty: safely get a COM property
You can try getting a property without catching exceptions:
if (ComAutoHelper.TryGetProperty(excel, "Version", out string? version))
{
    Console.WriteLine($"Excel version: {version}");
}
else
{
    Console.WriteLine("Property not found or failed.");
}

## Full Excel + Word automation demo

This WPF app runs both Excel and Word COM automation examples without any Interop DLLs:

- Writes data into Excel
- Formats Word paragraph
- Inspects COM members via `ComTypeInspector`

Source: [ComAutoWrapperDemo](https://github.com/pmonitor0/ComAutoWrapperDemo)


## 📊 Comparison: OpenXML vs COM Automation

This section compares two popular approaches for automating Office documents in C#.

| Feature / Capability                              | OpenXML SDK           | COM Automation (`ComAutoWrapper`) |
|--------------------------------------------------|------------------------|------------------------------------|
| File-based read/write                            | ✅ Yes                | ❌ No                              |
| Live Office application control (Excel/Word)     | ❌ No                 | ✅ Yes                             |
| Handles password-protected files                 | ❌ No support         | ✅ Yes (if Office can open it)     |
| Supports running VBA macros                      | ❌ No                 | ✅ Yes                             |
| Reads current user selection                     | ❌ No                 | ✅ Yes                             |
| Formatting (color, styles, font size, etc.)      | ⚠️ Limited            | ✅ Full                            |
| Chart and graphic manipulation                   | ❌ No                 | ✅ Yes                             |
| Interactive editing of running instance          | ❌ No                 | ✅ Yes                             |
| Requires Interop DLLs                            | ❌ No                 | ❌ No (via ComAutoWrapper)         |
| Can be used without Office installed             | ✅ Yes                | ❌ No                              |
| Dependency size                                  | ✅ Small              | ✅ Small (via wrapper)             |

> ⚠️ Note: OpenXML is best for static document generation and server-side manipulation.  
> ✅ COM Automation is best for real-time document interaction and full feature access.

---

Using `ComAutoWrapper`, you get the **full power of Office** with the **ease of a lightweight, interop-free helper**, suitable for Excel and Word automation alike.

🙏 Köszönetnyilvánítás
A ChatGPT által nyújtott nagyon sok segítségért, amely hozzájárult a projekt egyes részeinek megvalósításához.

📄 License
MIT

