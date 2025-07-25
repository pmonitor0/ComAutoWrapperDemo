# ComAutoWrapper

A lightweight and educational C# wrapper around COM automation. Focused on Excel and Word interop examples without primary interop assemblies.

## Features

- 📊 Excel automation (launching, writing values, formatting, reading)
- 📄 Word automation (insert tables, style ranges)
- 🧠 Introspection (list available COM methods/properties)
- 🧹 Centralized COM object cleanup
- 🔎 Running Object Table (ROT) based detection
- ✅ Cross-version compatibility (no PIA dependencies)

---

## 🔧 Examples

### Excel: Insert values and manipulate cells

```csharp
var excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
ComInvoker.SetProperty(excel, "Visible", true);

var workbooks = ComInvoker.GetProperty<object>(excel, "Workbooks");
ComInvoker.CallMethod(workbooks, "Add");

var workbook = ComInvoker.GetProperty<object>(excel, "Workbooks", new object[] { 1 });
var sheet = ComInvoker.GetProperty<object>(workbook, "Worksheets", new object[] { 1 });
ComInvoker.SetProperty(sheet, "Name", "Summary");

// Fill a 15x15 multiplication table
int[,] data = new int[15, 15];
for (int i = 0; i < 15; i++)
    for (int j = 0; j < 15; j++)
        data[i, j] = (i + 1) * (j + 1);

var range = ComReleaseHelper.Track(ComInvoker.GetProperty<object>(sheet, "Range", new object[] { "A1:O15" }));
ComInvoker.SetProperty(range, "Value", data);

// Close Excel
ComInvoker.CallMethod(workbook, "Close", false);
ComInvoker.CallMethod(excel, "Quit");
ComReleaseHelper.ReleaseAll();
```

---

### Word: Insert and format a table

```csharp
var wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application"));
ComInvoker.SetProperty(wordApp, "Visible", true);

var documents = ComInvoker.GetProperty<object>(wordApp, "Documents");
var doc = ComInvoker.CallMethod<object>(documents, "Add");

var range = ComInvoker.GetProperty<object>(doc, "Content");
var tables = ComInvoker.GetProperty<object>(doc, "Tables");
var table = ComInvoker.CallMethod<object>(tables, "Add", range, 3, 3);

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

        ComReleaseHelper.Track(cellRange);
        ComReleaseHelper.Track(cell);
    }
}

ComInvoker.CallMethod(doc, "Close", false);
ComInvoker.CallMethod(wordApp, "Quit");
ComReleaseHelper.ReleaseAll();
```

---


---

### Excel: Advanced demo with selection, styling and type inspection

```csharp
var excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
ComInvoker.SetProperty(excel, "Visible", true);
ComInvoker.SetProperty(excel, "DisplayAlerts", true);
var workbooks = ComInvoker.GetProperty<object>(excel, "Workbooks");
ComInvoker.CallMethod(workbooks, "Add");
var workbook = ComInvoker.GetProperty<object>(excel, "Workbooks", new object[] { 1 });
var sheet = ComInvoker.GetProperty<object>(workbook, "Worksheets", new object[] { 1 });

int[,] data = new int[15, 15];
for (int i = 0; i < 15; i++)
    for (int j = 0; j < 15; j++)
        data[i, j] = (i + 1) * (j + 1);

var range = ComReleaseHelper.Track(ComInvoker.GetProperty<object>(sheet, "Range", "A1:O15"));
ComInvoker.SetProperty(range, "Value", data);

// Type inspection
var (methods, propsGet, propsSet) = ComTypeInspector.ListMembers(workbook);
methods.ForEach(Console.WriteLine);

// Highlight selected cell range
var selected = ExcelSelectionHelper.GetSelectedCellObjects(excel);
foreach (var (row, col, cell) in selected)
    ComReleaseHelper.Track(cell);

var coords = ExcelSelectionHelper.GetSelectedCellCoordinates(excel);
foreach (var (row, col) in coords)
    Console.WriteLine($"Selected: Row={row}, Col={col}");

// Clean up
ComInvoker.CallMethod(workbook, "Close", false);
ComInvoker.CallMethod(excel, "Quit");
ComReleaseHelper.ReleaseAll();
```


## 🧰 Utility Classes

| Helper             | Description                              |
|--------------------|------------------------------------------|
| `ComInvoker`       | Wrapper around `InvokeMember` for get/set/call |
| `ComReleaseHelper` | Tracks and releases COM references       |
| `ComAutoHelper`    | Introspection, safe get, property check  |
| `ComValueConverter`| Converts colors, booleans to OLE types   |
| `ComRotHelper`     | Finds running COM instances (e.g., Excel)|
| `ExcelHelper`      | High-level Excel-specific helpers        |
| `WordStyleHelper`  | Range formatting for Word tables/ranges  |

---

## 💡 Notes

- Always use `ComReleaseHelper.Track(...)` when keeping temporary references.
- Always call `ComReleaseHelper.ReleaseAll()` once at the end.
- Avoid using Marshal.ReleaseComObject directly unless debugging.

---

## 🔍 Goal

This project serves as a minimal, but flexible educational base to understand how to interact with COM objects from C# without PIA or interop DLLs.


---

## 🔒 Safe Property Access with `ComAutoHelper`

When working with late-bound COM objects, property access may fail or throw exceptions.  
To make this safer, you can use `ComAutoHelper.TryGetProperty(...)`:

csharp
if (ComAutoHelper.TryGetProperty(excelApp, "Version", out string? version))
    Console.WriteLine("Excel version: " + version);
This avoids try/catch clutter and ensures proper type-checking.
You can also pass index parameters (e.g. Sheets[1]) like this:

csharp
ComAutoHelper.TryGetProperty(workbook, "Worksheets", out object? sheet, 1);
You can check if a property exists before trying to access it:

csharp
if (ComAutoHelper.PropertyExists(excelApp, "DisplayAlerts"))
    Console.WriteLine("Property is available.");
To retrieve the process ID of an Excel instance (useful for process monitoring or diagnostics):

csharp
var proc = ComAutoHelper.GetProcessByExcelHandle(excelApp);
Console.WriteLine("Excel PID: " + proc.Id);
These helpers work together with ComReleaseHelper.Track(...) and ComInvoker to provide a clean, exception-safe COM automation experience.

---

### 📌 How to apply


MIT License.