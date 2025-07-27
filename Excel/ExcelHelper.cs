using System;
using System.Collections.Generic;

namespace ComAutoWrapper
{
	/// <summary>
	/// Magas szintű segédosztály Excel COM automatizáláshoz.
	/// Lehetővé teszi a munkafüzetek, munkalapok és cellatartományok egyszerű elérését.
	/// </summary>
	public static class ExcelHelper
	{
		/// <summary>
		/// Lekéri az összes megnyitott Excel munkafüzetet egy Excel Application COM objektumból.
		/// </summary>
		/// <param name="excelApplication">Az Excel Application COM objektum (pl. <c>Excel.Application</c>).</param>
		/// <returns>A munkafüzetek listája (<c>Workbook</c> COM objektumokként).</returns>
		public static List<object> GetWorkbooks(object excelApplication)
		{
			var result = new List<object>();

			try
			{
				// Workbooks gyűjtemény lekérése
				var workbooks = ComInvoker.GetProperty<object>(excelApplication, "Workbooks");
				if (workbooks == null)
					return result;

				int count = ComInvoker.GetProperty<int>(workbooks, "Count");

				for (int i = 1; i <= count; i++) // Excel COM indexelés: 1-alapú
				{
					var wb = ComInvoker.GetProperty<object>(workbooks, "Item", new object[] { i });
					if (wb != null)
						result.Add(wb);
				}
			}
			catch
			{
				// hibakezelés opcionális: log, kivétel dobás, stb.
			}

			return result;
		}

		/// <summary>
		/// Lekéri az összes munkalapot egy adott Excel munkafüzetből.
		/// </summary>
		/// <param name="workbook">A <c>Workbook</c> COM objektum.</param>
		/// <returns>A munkalapok listája (<c>Worksheet</c> COM objektumokként).</returns>
		public static List<object> GetWorksheets(object workbook)
		{
			var result = new List<object>();

			try
			{
				var sheets = ComInvoker.GetProperty<object>(workbook, "Sheets");
				if (sheets == null)
					return result;

				int count = ComInvoker.GetProperty<int>(sheets, "Count");

				for (int i = 1; i <= count; i++) // Excel indexelés: 1-alapú
				{
					var sheet = ComInvoker.GetProperty<object>(sheets, "Item", new object[] { i });
					if (sheet != null)
						result.Add(sheet);
				}
			}
			catch
			{
				// opcionális: log vagy kivétel továbbdobása
			}

			return result;
		}

		/// <summary>
		/// Lekér egy cellatartományt (range) a megadott munkalapról Excel-címezés alapján (pl. "B2:D5").
		/// </summary>
		/// <param name="worksheet">A <c>Worksheet</c> COM objektum, ahonnan a tartományt le szeretnénk kérni.</param>
		/// <param name="address">A tartomány címe Excel-formátumban (pl. "A1", "B2:C3").</param>
		/// <returns>A tartomány (<c>Range</c> COM objektum), vagy <c>null</c>, ha hiba történt.</returns>
		public static object? GetRange(object worksheet, string address)
		{
			try
			{
				return ComInvoker.GetProperty<object>(worksheet, "Range", new object[] { address });
			}
			catch
			{
				return null;
			}
		}
	}
}
