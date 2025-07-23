# ComAutoWrapper

**ComAutoWrapper** egy minimalista és oktató jellegű C# wrapper könyvtár, amely megkönnyíti a **COM-objektumok késői kötésű** (`late binding`) használatát `IDispatch` alapon – külső interop DLL nélkül.

A cél: egyszerűen és biztonságosan vezérelhessük az Excel, Word vagy más COM-alapú alkalmazásokat .NET-ből.

---

## 🎯 Fő előnyök

- ✅ **Könnyű használat** – magas szintű metódusokkal
- ✅ **Interop DLL-mentes** – nem kell Microsoft.Office.Interop referenciát hozzáadni
- ✅ **Hibakezelés és felszabadítás** beépítve
- ✅ **Excel & Word példák** dokumentáltan

---

## 📦 Telepítés

A NuGet csomag hamarosan elérhető:

```bash
dotnet add package ComAutoWrapper
Fejlesztés alatt, lokális .nupkg is használható addig.

🔧 Fő komponensek
Osztály	Szerepe
ComInvoker	Property/metódus elérés late binding-gel
ComReleaseHelper	COM-objektumok nyomon követése és felszabadítása (FinalReleaseComObject)
ComValueConverter	.NET típusok → COM-kompatibilis (pl. Color → OLE_COLOR)
ComRotHelper	Excel példányok listázása a Running Object Table-ből
ExcelHelper	Workbook / Worksheet / Range lekérdezés
ExcelSelectionHelper	Kijelölt tartomány kezelése, koordináta lekérdezés
ExcelStyleHelper	Cella háttérszínezés
WordHelper	Teljes minta Word táblázat beszúrására
WordStyleHelper	Word Range formázása (pl. félkövér + háttérszín)
ComTypeInspector	COM tagok introspektív lekérdezése ITypeInfo alapján

🧪 Példák
📘 Excel – cellák formázása
csharp
var app = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
ComInvoker.SetProperty(app!, "Visible", true);

var workbooks = ComInvoker.GetProperty<object>(app!, "Workbooks");
ComInvoker.CallMethod(workbooks!, "Add");

var apps = ComRotHelper.GetExcelApplications();

foreach (var excel in apps)
{
    foreach (var wb in ExcelHelper.GetWorkbooks(excel))
    {
        foreach (var sheet in ExcelHelper.GetWorksheets(wb))
        {
            var range = ExcelHelper.GetRange(sheet, "B2:D2");
            ComInvoker.SetProperty(range, "Value", "Teszt");

            var interior = ComInvoker.GetProperty<object>(range, "Interior");
            int szin = ComValueConverter.ToOleColor(System.Drawing.Color.LightGreen);
            ComInvoker.SetProperty(interior!, "Color", szin);

            ComReleaseHelper.Track(range);
            ComReleaseHelper.Track(interior);
        }
        ComInvoker.SetProperty(wb, "Saved", ComValueConverter.ToComBool(true));
        ComInvoker.CallMethod(wb, "Close", ComValueConverter.ToComBool(true));
        ComReleaseHelper.Track(wb);
    }
    ComInvoker.CallMethod(excel, "Quit");
    ComReleaseHelper.Track(excel);
}
ComReleaseHelper.ReleaseAll();
📝 Word – táblázat beszúrása és formázása
csharp
var wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
ComInvoker.SetProperty(wordApp!, "Visible", true);
ComInvoker.SetProperty(wordApp!, "DisplayAlerts", false);

var documents = ComInvoker.GetProperty<object>(wordApp!, "Documents");
var doc = ComInvoker.CallMethod<object>(documents!, "Add");

var range = ComInvoker.GetProperty<object>(doc!, "Content");
var tables = ComInvoker.GetProperty<object>(doc!, "Tables");
var table = ComInvoker.CallMethod<object>(tables!, "Add", range, 3, 3);

for (int row = 1; row <= 3; row++)
{
    for (int col = 1; col <= 3; col++)
    {
        var cell = ComInvoker.CallMethod<object>(table, "Cell", row, col);
        var cellRange = ComInvoker.GetProperty<object>(cell, "Range");
        ComInvoker.SetProperty(cellRange, "Text", $"R{row}C{col}");

        if (row == 1)
        {
            WordStyleHelper.ApplyStyle(
                cellRange,
                fontColor: ComValueConverter.ToOleColor(Color.White),
                backgroundColor: ComValueConverter.ToOleColor(Color.DarkRed),
                bold: true
            );
        }

        ComReleaseHelper.Track(cell);
        ComReleaseHelper.Track(cellRange);
    }
}

ComInvoker.SetProperty(doc, "Saved", ComValueConverter.ToComBool(true));
ComInvoker.CallMethod(doc, "Close", ComValueConverter.ToComBool(false));
ComInvoker.CallMethod(wordApp!, "Quit");

ComReleaseHelper.Track(table);
ComReleaseHelper.Track(tables);
ComReleaseHelper.Track(doc);
ComReleaseHelper.Track(documents);
ComReleaseHelper.Track(wordApp);
ComReleaseHelper.ReleaseAll();
🔐 License
MIT License
Szabadon használható oktatási és üzleti célra is.
Lásd: LICENSE

🙋‍♂️ Kinek ajánlott?
.NET fejlesztőknek, akik nem akarnak Office Interop DLL-t használni

Oktatóknak, akik bemutatnák a IDispatch-alapú elérést

Haladó automatizálóknak, akik minimalista, de stabil COM API-t keresnek

