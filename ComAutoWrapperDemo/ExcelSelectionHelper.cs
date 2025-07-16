using ComAutoWrapperDemo;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace ComAutoWrapper
{
	public class ExcelSelectionHelper
	{
		public static void SelectCells(object sheet, params string[] addresses)
		{
			if (addresses.Length == 0)
				return;

			// 1. Get Application from sheet
			var app = ComInvoker.GetProperty<object>(sheet, "Application");

			// 2. Get individual ranges from addresses
			var ranges = addresses
				.Select(addr => ComInvoker.GetProperty<object>(sheet, "Range", new object[] { addr }))
				.ToArray();

			// 3. Combine with Union if needed
			object combined = ranges[0];
			for (int i = 1; i < ranges.Length; i++)
			{
				combined = ComInvoker.CallMethod<object>(app, "Union", combined, ranges[i]);
			}

			// 4. Select the final range
			ComInvoker.CallMethod(combined, "Select");
		}

		public static List<(int Row, int Column)> GetSelectedCellCoordinates(object excel)
		{
			var coordinates = new List<(int Row, int Column)>();

			var selection = ComInvoker.GetProperty<object>(excel, "Selection");
			var areas = ComInvoker.GetProperty<object>(selection!, "Areas");
			int areaCount = ComInvoker.GetProperty<int>(areas!, "Count");

			for (int a = 1; a <= areaCount; a++)
			{
				var area = ComInvoker.GetProperty<object>(areas!, "Item", new object[] { a });
				var cellsInArea = ComInvoker.GetProperty<object>(area!, "Cells");
				int count = ComInvoker.GetProperty<int>(cellsInArea!, "Count");

				for (int i = 1; i <= count; i++)
				{
					var cell = ComInvoker.GetProperty<object>(cellsInArea!, "Item", new object[] { i });
					string address = ComInvoker.GetProperty<string>(cell!, "Address");

					var match = Regex.Match(address, @"\$([A-Z]+)\$(\d+)");
					if (match.Success)
					{
						string colLetter = match.Groups[1].Value;
						int row = int.Parse(match.Groups[2].Value);
						int col = ColumnLetterToNumber(colLetter);
						coordinates.Add((row, col));
					}
				}
			}

			return coordinates;
		}

		public static List<(int Row, int Column, object Cell)> GetSelectedCellObjects(object excel)
		{
			var result = new List<(int Row, int Column, object Cell)>();

			var selection = ComInvoker.GetProperty<object>(excel, "Selection");
			var areas = ComInvoker.GetProperty<object>(selection!, "Areas");
			int areaCount = ComInvoker.GetProperty<int>(areas!, "Count");

			for (int a = 1; a <= areaCount; a++)
			{
				var area = ComInvoker.GetProperty<object>(areas!, "Item", new object[] { a });
				var cellsInArea = ComInvoker.GetProperty<object>(area!, "Cells");
				int count = ComInvoker.GetProperty<int>(cellsInArea!, "Count");

				for (int i = 1; i <= count; i++)
				{
					var cell = ComInvoker.GetProperty<object>(cellsInArea!, "Item", new object[] { i });
					string address = ComInvoker.GetProperty<string>(cell!, "Address");

					var match = Regex.Match(address, @"\$([A-Z]+)\$(\d+)");
					if (match.Success)
					{
						string colLetter = match.Groups[1].Value;
						int row = int.Parse(match.Groups[2].Value);
						int col = ColumnLetterToNumber(colLetter);
						result.Add((row, col, cell));
					}
				}
			}

			return result;
		}

		public static int ColumnLetterToNumber(string col)
		{
			int sum = 0;
			foreach (char c in col)
			{
				sum *= 26;
				sum += (char.ToUpper(c) - 'A' + 1);
			}
			return sum;
		}
	}
}
