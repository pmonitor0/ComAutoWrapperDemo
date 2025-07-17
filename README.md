# ComAutoWrapperDemo

This project is a **WPF-based demo** for using the [ComAutoWrapper](https://www.nuget.org/packages/ComAutoWrapper) NuGet package.  
It demonstrates how to work with COM objects (e.g., Excel) without Interop DLLs in a simple, typed, and minimalistic way.

## Key Features

- Excel automation from WPF
- COM property access (`GetProperty<T>`, `SetProperty`)
- COM method calls (`CallMethod<T>`)
- COM member introspection:
  - Methods (`Method`)
  - Readable (`PropertyGet`)
  - Writable (`PropertySet`) properties

## How the demo works

After startup, the WPF application:

1. Runs in console mode:
   - Launches Excel
   - Fills cells with a multiplication table
   - Inspects `Worksheet` COM members
   - Prints results to the console
   - Launches Word
   - Inserts formatted text into Word
   - Closes Word
2. Then it shows the `MainWindow`.

## Installation

1. Clone the repo:
```bash
git clone https://github.com/pmonitor0/ComAutoWrapperDemo.git
2. Open the .sln file in Visual Studio.

3. Make sure the NuGet package (ComAutoWrapper) is installed.

Full Excel + Word automation demo
This WPF app runs both Excel and Word COM automation examples without any Interop DLLs:

Writes data into Excel

Formats a Word paragraph

Inspects COM members via ComTypeInspector

Source: ComAutoWrapperDemo

Requirements
Windows (due to COM)

.NET 6/7/8/9

Installed Microsoft Excel

Installed Microsoft Word

Related Projects
ComAutoWrapper (NuGet)

ComAutoWrapper (GitHub)

Acknowledgements
The idea and development of this project were a collaborative effort with ChatGPT. ChatGPT provided extensive support.

License
MIT