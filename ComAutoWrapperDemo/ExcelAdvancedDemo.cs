using System;
using System.Diagnostics;
using System.Drawing;

namespace ComAutoWrapper
{
    public static class ExcelAdvancedDemo
    {
        public static void Run()
        {
			var excel = CreateExcelApplication();

			Process? proc = ComAutoHelper.GetProcessByExcelHandle(excel);
			Console.WriteLine("Excel PID: " + proc.Id);

			// Új munkafüzet és munkalap
			object workbook = AddWorkbook(excel);
            var sheet = AddWorksheet(workbook!, "Summary");
            

            // Szorzótábla beírása
            int[,] data = new int[15, 15];
            for (int i = 0; i < 15; i++)
                for (int j = 0; j < 15; j++)
                    data[i, j] = (i + 1) * (j + 1);

            var range = ComReleaseHelper.Track(ComInvoker.GetProperty<object>(sheet!, "Range", new object[] { "A1:O15" }));
            ComInvoker.SetProperty(range!, "Value", data);

            // Workbook metaadatok
            Console.WriteLine("\nWorkbook: " + ComInvoker.GetProperty<string>(workbook!, "Name"));
            if (ComAutoHelper.TryGetProperty(excel!, "Version", out string? ver))
                Console.WriteLine("Excel version: " + ver);

            // Típuslistázás
            var (methods, getters, setters) = ComTypeInspector.ListMembers(workbook!);
            Console.WriteLine("\nWorkbook methods:");
            methods.ForEach(Console.WriteLine);

            // Több alkalmazás és munkafüzet bejárása
            var apps = ComRotHelper.GetExcelApplications();
            foreach (var app in apps)
            {
                foreach (var wb in ExcelHelper.GetWorkbooks(app))
                {
                    foreach (var sh in ExcelHelper.GetWorksheets(wb))
                    {
                        var r = ComReleaseHelper.Track(ExcelHelper.GetRange(sh, "B2:D2"));
                        if (r != null)
                        {
                            ComInvoker.SetProperty(r, "Value", "Teszt");
                            var color = ComValueConverter.ToOleColor(Color.Blue);
                            var r2 = ExcelHelper.GetRange(sheet!, "A1:B1");
                            var interior = ComInvoker.GetProperty<object>(r2!, "Interior");
                            ComInvoker.SetProperty(interior!, "Color", color);
                        }
                    }
                }
            }

            // Kijelölt cellák objektumai és koordinátái
            var cells = ExcelSelectionHelper.GetSelectedCellObjects(excel);
            foreach (var (row, col, cell) in cells)
            {
				Console.WriteLine($"Row: {cell}, Column: {col}");
                ComReleaseHelper.Track(cell);
            }

            var coords = ExcelSelectionHelper.GetSelectedCellCoordinates(excel);
            foreach (var (row, col) in coords)
                Console.WriteLine($"Row={row}, Column={col}");

			// ComAutoHelper használatának bemutatása
			if (ComAutoHelper.TryGetProperty<string>(excel!, "Version", out var version))
				Console.WriteLine($"\n[ComAutoHelper] Excel verzió: {version}");
			else
				Console.WriteLine("\n[ComAutoHelper] Nem sikerült lekérdezni az Excel verzióját.");

			bool exists = ComAutoHelper.PropertyExists(excel!, "DisplayAlerts");
			Console.WriteLine($"[ComAutoHelper] 'DisplayAlerts' property létezik: {exists}");

			bool fakeExists = ComAutoHelper.PropertyExists(excel!, "NemLetezoProperty123");
			Console.WriteLine($"[ComAutoHelper] 'NemLetezoProperty123' property létezik: {fakeExists}");

			Console.WriteLine("\nDemo vége. Enter után bezárunk mindent.");
            Console.ReadKey(true);

            // Bezárás és takarítás
            ComInvoker.CallMethod(workbook!, "Close", false);
            ComInvoker.CallMethod(excel!, "Quit");
            ComReleaseHelper.Track(excel);
            ComReleaseHelper.Track(workbook);
            ComReleaseHelper.ReleaseAll();
        }

        static object CreateExcelApplication()
        {
			var excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));
			ComInvoker.SetProperty(excel!, "Visible", true);
			ComInvoker.SetProperty(excel!, "DisplayAlerts", true);
            return excel!;
		}

        static object AddWorkbook(object excelApp)
        {
			var workbooks = ComInvoker.GetProperty<object>(excelApp!, "Workbooks");
			ComInvoker.CallMethod(workbooks!, "Add");
            var workbook = ComInvoker.GetProperty<object>(excelApp!, "ActiveWorkbook");
            return workbook!;
		}

		static object AddWorksheet(object workbook,string sheetName)
        {
            var sheets = ComInvoker.GetProperty<object>(workbook!, "Worksheets");
            var sheet = ComInvoker.CallMethod(sheets!, "Add");
			//var sheet = ComInvoker.GetProperty<object>(workbook!, "ActiveSheet");
			ComInvoker.SetProperty(sheet!, "Name", sheetName);
            return sheet!;
		}
    }
}