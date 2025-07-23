using System;
using System.Drawing;

namespace ComAutoWrapper
{
	/// <summary>
	/// Segédosztály a .NET típusok és COM-kompatibilis értékek közötti konverzióhoz.
	/// Hasznos például Excel vagy Word automatizálás során, ahol OLE_COLOR vagy OLE_DATE típusokkal kell dolgozni.
	/// </summary>
	public static class ComValueConverter
	{
		/// <summary>
		/// Átalakít egy .NET <see cref="Color"/> színt OLE_COLOR formátumra (24 bites BGR egész szám).
		/// </summary>
		/// <param name="color">A .NET szín, amelyet konvertálni szeretnénk.</param>
		/// <returns>Az OLE_COLOR érték, amelyet a COM-kompatibilis API-k használnak (BGR sorrendű int).</returns>
		public static int ToOleColor(Color color)
		{
			return (color.B << 16) | (color.G << 8) | color.R;
		}

		/// <summary>
		/// Átalakít egy logikai értéket (bool) COM kompatibilis egész értékre: 1 (true) vagy 0 (false).
		/// </summary>
		/// <param name="value">A logikai érték.</param>
		/// <returns>1, ha true; 0, ha false.</returns>
		public static int ToComBool(bool value)
		{
			return value ? 1 : 0;
		}

		/// <summary>
		/// Átalakít egy .NET <see cref="DateTime"/> értéket OLE Automation Date formátumra (pl. Excel dátummező).
		/// </summary>
		/// <param name="value">A konvertálandó időpont.</param>
		/// <returns>Az OLE Automation Date formátumú dátum (double).</returns>
		public static double ToOleDate(DateTime value)
		{
			return value.ToOADate();
		}

		/// <summary>
		/// Átalakít egy OLE Automation Date értéket (double) .NET <see cref="DateTime"/> formátumra.
		/// </summary>
		/// <param name="value">Az OLE formátumú dátum (általában Excel vagy COM visszatérési érték).</param>
		/// <returns>A megfelelő <see cref="DateTime"/> példány.</returns>
		public static DateTime FromOleDate(double value)
		{
			return DateTime.FromOADate(value);
		}
	}
}
