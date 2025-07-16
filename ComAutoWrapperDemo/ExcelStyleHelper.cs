using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace ComAutoWrapper
{
	public class ExcelStyleHelper
	{
		public static void SetCellBackground(object cell, Color color)
		{
			var interior = ComInvoker.GetProperty<object>(cell, "Interior");
			ComInvoker.SetProperty(interior!, "Color", color);
		}
	}
}
