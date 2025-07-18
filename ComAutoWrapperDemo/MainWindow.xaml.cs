using System;
using System.Windows;
using System.Runtime.InteropServices;
using ComAutoWrapper;
using System.Runtime.InteropServices.Marshalling;
using System.CodeDom;
using System.Reflection;
using System.Security.Cryptography;
using System.Drawing;
using System.Text.RegularExpressions;
using System.Diagnostics;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

namespace ComAutoWrapperDemo
{
	public partial class MainWindow : Window
	{
		[System.Runtime.InteropServices.DllImport("kernel32.dll")]
		public static extern Boolean AllocConsole();

		private object? excel;
		private object? workbooks;
		private object? workbook;
		private object? WorkSheet;
		private object? rng;
		Process? _excel = null;

		public MainWindow()
		{
			AllocConsole();
			InitializeComponent();
		}

		private void StartExcel_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (_excel != null)
					return;
				excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);

				nint processID;
				int Hwnd = ComInvoker.GetProperty<int>(excel!, "Hwnd", null);
				GetWindowThreadProcessId(Hwnd, out processID);
				_excel = Process.GetProcessById(processID.ToInt32());


				ComInvoker.SetProperty(excel!, "Visible", true);
				Log("Excel started.");
				workbooks = ComInvoker.GetProperty<object>(excel!, "Workbooks");
				Log("Workbooks object acquired.");
			}
			catch (Exception ex)
			{
				Log($"ERROR: {ex.Message}");
			}
		}

		private void AddWorkbook_Click(object sender, RoutedEventArgs e)
		{
			if (_excel == null)
				return;
			try
			{
				ComInvoker.SetProperty(excel!, "DisplayAlerts", true);
				ComInvoker.CallMethod(workbooks!, "Add");
				Log("Workbook added.");
				workbook = ComInvoker.GetProperty<object>(excel!, "Workbooks", new object[] { 1 });

				WorkSheet = ComInvoker.GetProperty<object>(workbook!, "WorkSheets", new object[] { 1 });
				ComInvoker.SetProperty(WorkSheet!, "Name", "Summary");

				rng = ComInvoker.GetProperty<object>(WorkSheet!, "Range", new object[] { "A1:C2" });
				ComInvoker.SetProperty(rng!, "Value", new object[,] { { 1, 2, "=SUM(A1:B1)" }, { 3, 4, "=SUM(A2:B2)" } });  // Excel array
				ComInvoker.SetProperty(workbook!, "Saved", true);
			}
			catch (Exception ex)
			{
				Log($"ERROR: {ex.Message}");
			}
		}

		[DllImport("user32.dll", SetLastError = true)]
		static extern IntPtr GetWindowThreadProcessId(int hWnd, out IntPtr lpdwProcessId);

		private void QuitExcel_Click(object sender, RoutedEventArgs e)
		{
			if (_excel == null)
				return;
				try
			{

				var (methods, propsGet, propsSet) = ComTypeInspector.ListMembers(workbook!);

				Console.WriteLine("Methods:");
				methods.ForEach(Console.WriteLine);

				Console.WriteLine("PropertyGet:");
				propsGet.ForEach(Console.WriteLine);

				Console.WriteLine("PropertySet:");
				propsSet.ForEach(Console.WriteLine);

				var typeName = ComTypeInspector.GetTypeName(workbook!);
				Console.WriteLine($"COM type: {typeName}");

				string? name = ComInvoker.GetProperty<string>(workbook!, "Name");
				Console.WriteLine(name);

				ComInvoker.CallMethod(workbook!, "Close", (object)false);
				Log("Workbook closed.");
				ComInvoker.CallMethod(excel!, "Quit");
				Log("Excel closed.");
				
				if (_excel != null)
				{
					excel = null;
					_excel.Kill();
					_excel = null;
				}

			}
			catch (Exception ex)
			{
				Log($"ERROR: {ex.Message}");
			}
		}

		private void Log(string message)
		{
			LogBox.AppendText(message + Environment.NewLine);
			LogBox.ScrollToEnd();
		}

		private void Window_Initialized(object sender, EventArgs e)
		{
			Task.Run(() =>
			{
				RunExcelDemo();
				RunWordDemo();

				Dispatcher.Invoke(() =>
				{
					MessageBox.Show("A demók lefutottak.");
					this.Show(); // vagy akár megjeleníthetsz adatot is a UI-n
				});
			});
			this.Hide(); // amíg nem fut le a demo
		}

		

		private void RunExcelDemo()
		{
			if (_excel != null)
				return;
			excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application"));

			nint processID;
			int Hwnd = ComInvoker.GetProperty<int>(excel!, "Hwnd", null);
			GetWindowThreadProcessId(Hwnd, out processID);
			_excel = Process.GetProcessById(processID.ToInt32());

			ComInvoker.SetProperty(excel!, "Visible", true);
			ComInvoker.SetProperty(excel!, "DisplayAlerts", true);
			workbooks = ComInvoker.GetProperty<object>(excel!, "Workbooks");

			ComInvoker.CallMethod(workbooks!, "Add");
			workbook = ComInvoker.GetProperty<object>(excel!, "Workbooks", new object[] { 1 });

			WorkSheet = ComInvoker.GetProperty<object>(workbook!, "WorkSheets", new object[] { 1 });
			ComInvoker.SetProperty(WorkSheet!, "Name", "Summary");

			int[,] arr = new int[15, 15];
			for (int i = 0; i < 15; i++)
			{
				for (int j = 0; j < 15; j++)
				{
					arr[i, j] = (i + 1) * (j + 1);
				}
			}
			rng = ComInvoker.GetProperty<object>(WorkSheet!, "Range", new object[] { "A1:O15" });
			ComInvoker.SetProperty(rng!, "Value", arr);

			var (methods, propsGet, propsSet) = ComTypeInspector.ListMembers(workbook!);

			Console.WriteLine("Excel WorkBook Methods:");
			methods.ForEach(Console.WriteLine);

			Console.WriteLine("\nExcel WorkBook PropertyGet:");
			propsGet.ForEach(Console.WriteLine);

			Console.WriteLine("\nExcel WorkBook PropertySet:");
			propsSet.ForEach(Console.WriteLine);

			var typeName = ComTypeInspector.GetTypeName(workbook!);
			Console.WriteLine($"\nCOM type: {typeName}");

			string? name = ComInvoker.GetProperty<string>(workbook!, "Name");
			Console.WriteLine($"\nWorkbook Name: {name}");
			
			if (ComAutoHelper.TryGetProperty(excel!, "Version", out string? version))
				Console.WriteLine($"\nExcel version: {version}");
			else
				Console.WriteLine("\nProperty not found or failed.");
			bool exists = ComAutoHelper.PropertyExists(excel!, "DisplayAlerts0");
			if (exists)
				Console.WriteLine("\nProperty exists.");
			else
				Console.WriteLine("\nProperty not exists.");
			//Console.WriteLine("Select cells in the workbook, then press a key");
			//Console.ReadKey(true);

			ExcelSelectionHelper.SelectCells(WorkSheet!, new string[] { "A1", "B2", "C3", "D4" });

			var cells = ExcelSelectionHelper.GetSelectedCellObjects(excel);

			foreach (var (row, col, cell) in cells)
			{
				Console.WriteLine($"Cell at Row={row}, Column={col}");
				// Példa: háttérszín sárgára állítása
				ExcelStyleHelper.SetCellBackground(cell, Color.Yellow);
			}


			var selectedCells = ExcelSelectionHelper.GetSelectedCellCoordinates(excel);
			foreach (var (row, col) in selectedCells)
				Console.WriteLine($"Row={row}, Column={col}");

			Console.WriteLine("\nAfter pressing a key, we close Excel and then open Word");
			Console.ReadKey(true);
			ComInvoker.CallMethod(workbook!, "Close", (object)false);
			workbook = null;
			ComInvoker.CallMethod(excel!, "Quit");
			if (_excel != null)
			{
				excel = null;
				_excel.Kill();
				_excel = null;
			}
			
		}

		private void RunWordDemo()
		{
			var word = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!);
			ComInvoker.SetProperty(word!, "Visible", true);

			var docs = ComInvoker.GetProperty<object>(word!, "Documents");
			var doc = ComInvoker.CallMethod<object>(docs!, "Add");

			var content = ComInvoker.GetProperty<object>(doc!, "Content");

			var para = ComInvoker.GetProperty<object>(content!, "Paragraphs");
			var first = ComInvoker.GetProperty<object>(para!, "First");
			var range = ComInvoker.GetProperty<object>(first!, "Range");
			ComInvoker.SetProperty(range!, "Text", "Ez egy stílusos bekezdés.");
			WordStyleHelper.ApplyStyle(
				range!,
				fontColor: Color.Red,
				backgroundColor: Color.LightGreen,
				fontSize: 14,
				bold: true,
				italic: true,
				underline: true);

			var borders = ComInvoker.GetProperty<object>(range!, "Borders");
			ComInvoker.SetProperty(borders!, "OutsideLineStyle", 1); // wdLineStyleSingle*/

			//WordStyleHelper.ApplyBoldColoredBackground(range!, Color.Red, Color.Green, 16);

			var (methods, propsGet, propsSet) = ComTypeInspector.ListMembers(content!);
			Console.WriteLine("\nWord Methods:");
			methods.ForEach(Console.WriteLine);
			Console.WriteLine("\nWord PropertyGet:");
			propsGet.ForEach(Console.WriteLine);
			Console.WriteLine("\nWord PropertySet:");
			propsSet.ForEach(Console.WriteLine);




			ComInvoker.CallMethod(doc!, "SaveAs", "D:\\Temp\\DemoWord.docx");
			Console.WriteLine("Egy billentyű leütése után bezárjuk a word-ot.");
			Console.ReadKey(true);
			ComInvoker.CallMethod(word!, "Quit");
		}

		private object? wordApp;
		private object? wordDoc;

		private void StartWord_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				wordApp = Activator.CreateInstance(Type.GetTypeFromProgID("Word.Application")!);
				ComInvoker.SetProperty(wordApp!, "Visible", true);
				LogBox.AppendText("Word elindítva.\n");
			}
			catch (Exception ex)
			{
				LogBox.AppendText($"Hiba a Word indításakor: {ex.Message}\n");
			}
		}

		private void AddParagraph_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (wordApp == null)
				{
					LogBox.AppendText("A Word nincs elindítva.\n");
					return;
				}

				var docs = ComInvoker.GetProperty<object>(wordApp, "Documents");
				wordDoc = ComInvoker.CallMethod<object>(docs!, "Add");
				var content = ComInvoker.GetProperty<object>(wordDoc!, "Content");
				var paras = ComInvoker.GetProperty<object>(content!, "Paragraphs");
				var firstPara = ComInvoker.GetProperty<object>(paras!, "First");
				var range = ComInvoker.GetProperty<object>(firstPara!, "Range");

				ComInvoker.SetProperty(range!, "Text", "Ez egy formázott bekezdés.");
				ComInvoker.SetProperty(range!, "Bold", 1);
				var font = ComInvoker.GetProperty<object>(range!, "Font");
				ComInvoker.SetProperty(font!, "Size", 16);

				LogBox.AppendText("Formázott bekezdés létrehozva a Word dokumentumban.\n");
			}
			catch (Exception ex)
			{
				LogBox.AppendText($"Hiba a bekezdés létrehozásakor: {ex.Message}\n");
			}
		}

		private void QuitWord_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				if (wordApp != null)
				{
					ComInvoker.SetProperty(wordDoc!, "Saved", true);
					ComInvoker.CallMethod(wordApp, "Quit");
					Marshal.ReleaseComObject(wordApp);
					wordApp = null;
					LogBox.AppendText("Word bezárva.\n");
				}
			}
			catch (Exception ex)
			{
				LogBox.AppendText($"Hiba a Word bezárásakor: {ex.Message}\n");
			}

			GC.Collect();
			GC.WaitForPendingFinalizers();
		}

	}
}
