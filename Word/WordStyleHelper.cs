using System;
using System.Collections.Generic;
using System.Linq;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ComAutoWrapper
{
	/// <summary>
	/// Segédosztály Word tartományok (<c>Range</c>) stílusainak beállításához COM automatizálással.
	/// </summary>
	public class WordStyleHelper
	{
		/// <summary>
		/// Általános stílusbeállító metódus, amely egy Word <c>Range</c> objektumra alkalmaz stílusokat: betűszín, háttérszín, méret, félkövérség stb.
		/// </summary>
		/// <param name="range">A formázandó <c>Range</c> COM objektum.</param>
		/// <param name="fontColor">Szövegszín OLE_COLOR (BGR int) formátumban. Ha <c>null</c>, nem módosul.</param>
		/// <param name="backgroundColor">Háttérszín OLE_COLOR (BGR int) formátumban. Ha <c>null</c>, nem módosul.</param>
		/// <param name="fontSize">A betűméret pontban (pl. 12.0). Ha <c>null</c>, nem módosul.</param>
		/// <param name="bold"><c>true</c>, ha félkövérre szeretnéd állítani.</param>
		/// <param name="italic"><c>true</c>, ha dőltre szeretnéd állítani.</param>
		/// <param name="underline"><c>true</c>, ha aláhúzás szükséges.</param>
		public static void ApplyStyle(
			object range,
			int? fontColor = null,
			int? backgroundColor = null,
			float? fontSize = null,
			bool bold = false,
			bool italic = false,
			bool underline = false)
		{
			var font = ComReleaseHelper.Track(ComInvoker.GetProperty<object>(range, "Font"));
			var shading = ComReleaseHelper.Track(ComInvoker.GetProperty<object>(range, "Shading"));

			if (bold)
				ComInvoker.SetProperty(font!, "Bold", 1);
			if (italic)
				ComInvoker.SetProperty(font!, "Italic", 1);
			if (underline)
				ComInvoker.SetProperty(font!, "Underline", 1);

			if (fontColor.HasValue)
				ComInvoker.SetProperty(font!, "Color", fontColor);
			if (fontSize.HasValue)
				ComInvoker.SetProperty(font!, "Size", fontSize.Value);
			if (backgroundColor.HasValue)
				ComInvoker.SetProperty(shading!, "BackgroundPatternColor", backgroundColor);
		}

		/// <summary>
		/// Gyors formázás: félkövér szöveg, háttér- és betűszín, megadott betűmérettel.
		/// </summary>
		/// <param name="range">A formázandó <c>Range</c> COM objektum.</param>
		/// <param name="fontColor">Szövegszín OLE_COLOR (BGR int) formátumban.</param>
		/// <param name="backgroundColor">Háttérszín OLE_COLOR (BGR int) formátumban.</param>
		/// <param name="fontSize">A betűméret pontban (alapértelmezett: 12.0).</param>
		public static void ApplyBoldColoredBackground(object range, int fontColor, int backgroundColor, float fontSize = 12f)
		{
			var font = ComReleaseHelper.Track(ComInvoker.GetProperty<object>(range, "Font"));
			var shading = ComReleaseHelper.Track(ComInvoker.GetProperty<object>(range, "Shading"));

			ComInvoker.SetProperty(range, "Bold", 1);
			ComInvoker.SetProperty(font!, "Color", fontColor);
			ComInvoker.SetProperty(font!, "Size", fontSize);
			ComInvoker.SetProperty(shading!, "BackgroundPatternColor", backgroundColor);
		}
	}
}
