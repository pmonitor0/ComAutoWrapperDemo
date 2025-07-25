# ComAutoWrapperDemo

This project is a **WPF-based demo** for using the [`ComAutoWrapper`](https://www.nuget.org/packages/ComAutoWrapper) NuGet package.  
It demonstrates how to automate COM objects (like Excel or Word) **without any Interop DLLs**, in a **type-safe and simplified** way.

---

## 🚀 Features Demonstrated

- Excel and Word automation from WPF
- COM property access:
  - `GetProperty<T>(...)`
  - `SetProperty(...)`
- COM method invocation:
  - `CallMethod<T>(...)`
- COM introspection (via `ComTypeInspector`)
  - List of readable, writable properties and methods
- Excel selection helpers:
  - Selecting specific cells
  - Querying selected cell positions (row/column)
  - Highlighting `UsedRange`

---

## 🔧 How the Demo Works

When launched, the WPF app:

1. **Console-mode automation (auto-run)**
   - Launches Excel
   - Fills it with a multiplication table
   - Inspects and prints `Worksheet` members
   - Launches Word
   - Inserts a styled paragraph
   - Quits Word
2. **Then** shows the main WPF window (`MainWindow`)

---

## 📦 Excel Selection Helpers (from `ComAutoWrapper`)

These helper methods are built into the `ComAutoWrapper` NuGet package under the `ComSelectionHelper` class:

| Method | Description |
|--------|-------------|
| `SelectCells(excel, sheet, "A1", "C3", "F5")` | Selects multiple (non-contiguous) cells |
| `GetSelectedCellCoordinates(excel)` | Returns list of `(row, column)` for user-selected cells |
| `HighlightUsedRange(sheet)` | Highlights the entire `UsedRange` in Excel |

You can reuse these from **any .NET app** (console, WPF, WinForms) — without any Interop reference.

---

## 💻 Prerequisites

- Windows OS (due to COM)
- Installed Microsoft Excel and/or Word
- .NET 6 / 7 / 8 / 9
- No Interop DLLs required

---

## 📥 Installation

1. Clone the repository:
   ```bash
   git clone https://github.com/pmonitor0/ComAutoWrapperDemo.git
Open the .sln file in Visual Studio

Ensure the ComAutoWrapper NuGet package is installed

🔗 Related Projects
ComAutoWrapper (NuGet)

ComAutoWrapper (GitHub)

🙏 Acknowledgment
This project was created as a collaborative effort with the support of ChatGPT.
Many of the ideas, refactorings, and enhancements came from an ongoing back-and-forth conversation.

📄 License
MIT