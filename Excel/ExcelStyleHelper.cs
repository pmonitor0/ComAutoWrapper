using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Drawing;

namespace ComAutoWrapper
{
	/// <summary>
	/// Segédosztály Excel cellák stílusának módosításához (pl. háttérszín).
	/// </summary>
	public class ExcelStyleHelper
	{
		/// <summary>
		/// Beállítja egy Excel cella háttérszínét (interior color) a megadott <see cref="Color"/> érték alapján.
		/// </summary>
		/// <param name="cell">A cél cella COM objektum (típus: <c>Range</c>).</param>
		/// <param name="color">A kívánt háttérszín .NET <see cref="Color"/> típusban.</param>
		public static void SetCellBackground(object cell, Color color)
		{
			var interior = ComInvoker.GetProperty<object>(cell, "Interior");
			ComInvoker.SetProperty(interior!, "Color", color);
		}
	}
}
