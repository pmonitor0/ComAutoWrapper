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

🙏 Köszönetnyilvánítás
A ChatGPT által nyújtott segítségért, amely hozzájárult a projekt egyes részeinek megvalósításához.

📄 License
MIT
