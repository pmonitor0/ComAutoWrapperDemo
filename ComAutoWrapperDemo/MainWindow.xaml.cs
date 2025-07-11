using System;
using System.Windows;
using System.Runtime.InteropServices;
using ComAutoWrapper;
using System.Runtime.InteropServices.Marshalling;
using System.CodeDom;
using System.Reflection;

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

		public MainWindow()
		{
			AllocConsole();
			InitializeComponent();
		}

		private void StartExcel_Click(object sender, RoutedEventArgs e)
		{
			try
			{
				excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);
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

		private void QuitExcel_Click(object sender, RoutedEventArgs e)
		{
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
				workbook = null;
				ComInvoker.CallMethod(excel!, "Quit");
				Log("Excel closed.");
				if (rng != null) Marshal.ReleaseComObject(rng!);
				if (WorkSheet != null) Marshal.ReleaseComObject(WorkSheet);
				if (workbooks != null) Marshal.ReleaseComObject(workbooks);


				if (excel != null) Marshal.ReleaseComObject(excel);
				workbooks = null;
				workbook = null;
				WorkSheet = null;
				rng = null;
				excel = null;

				GC.Collect();
				GC.WaitForPendingFinalizers();
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
			excel = Activator.CreateInstance(Type.GetTypeFromProgID("Excel.Application")!);
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
			rng = ComInvoker.GetProperty<object>(WorkSheet!, "Range", new object[] {"A1:O15"} );
			ComInvoker.SetProperty(rng!, "Value", arr);

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
			Console.WriteLine("Egy billentyű leütése után bezárjuk az excel-t.");
			Console.ReadKey(true);
			ComInvoker.CallMethod(workbook!, "Close", (object)false);
			workbook = null;
			ComInvoker.CallMethod(excel!, "Quit");
			if (rng != null) Marshal.ReleaseComObject(rng!);
			if (WorkSheet != null) Marshal.ReleaseComObject(WorkSheet);
			if (workbooks != null) Marshal.ReleaseComObject(workbooks);


			if (excel != null) Marshal.ReleaseComObject(excel);
			workbooks = null;
			workbook = null;
			WorkSheet = null;
			rng = null;
			excel = null;

			GC.Collect();
			GC.WaitForPendingFinalizers();
		}
	}
}
