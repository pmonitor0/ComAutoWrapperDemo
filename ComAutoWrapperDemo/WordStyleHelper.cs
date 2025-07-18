using ComAutoWrapper;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ComAutoWrapperDemo
{
    class WordStyleHelper
    {
		public static void ApplyStyle(
			object range,
			Color? fontColor = null,
			Color? backgroundColor = null,
			float? fontSize = null,
			bool bold = false,
			bool italic = false,
			bool underline = false)
		{
			var font = ComInvoker.GetProperty<object>(range, "Font");
			var shading = ComInvoker.GetProperty<object>(range, "Shading");

			if (bold)
				ComInvoker.SetProperty(range, "Bold", 1);
			if (italic)
				ComInvoker.SetProperty(range, "Italic", 1);
			if (underline)
				ComInvoker.SetProperty(range, "Underline", 1);

			if (fontColor.HasValue)
				ComInvoker.SetProperty(font!, "Color", fontColor.Value);
			if (fontSize.HasValue)
				ComInvoker.SetProperty(font!, "Size", fontSize.Value);
			if (backgroundColor.HasValue)
				ComInvoker.SetProperty(shading!, "BackgroundPatternColor", backgroundColor.Value);

			if (font != null) Marshal.ReleaseComObject(font);
			if (shading != null) Marshal.ReleaseComObject(shading);
		}


		public static void ApplyBoldColoredBackground(object range, Color fontColor, Color backgroundColor, float fontSize = 12f)
		{
			var font = ComInvoker.GetProperty<object>(range, "Font");
			var shading = ComInvoker.GetProperty<object>(range, "Shading");

			ComInvoker.SetProperty(range, "Bold", 1);
			ComInvoker.SetProperty(font!, "Color", fontColor);
			ComInvoker.SetProperty(font!, "Size", fontSize);
			ComInvoker.SetProperty(shading!, "BackgroundPatternColor", backgroundColor);

			if (font != null) Marshal.ReleaseComObject(font);
			if (shading != null) Marshal.ReleaseComObject(shading);
		}

	}
}
