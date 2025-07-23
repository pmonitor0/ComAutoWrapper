using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace ComAutoWrapper
{
	/// <summary>
	/// Segédosztály Excel COM tartományok kijelöléséhez, színezéséhez, valamint cellák koordinátáinak lekérdezéséhez.
	/// </summary>
	public class ExcelSelectionHelper
	{
		/// <summary>
		/// Kijelöli az aktív munkalap használt tartományát (<c>UsedRange</c>).
		/// </summary>
		/// <param name="worksheet">A <c>Worksheet</c> COM objektum, amelyen a kijelölést végezzük.</param>
		public static void SelectUsedRange(object worksheet)
		{
			var usedRange = ComReleaseHelper.Track(ComInvoker.GetProperty<object>(worksheet, "UsedRange"));
			ComInvoker.CallMethod(usedRange!, "Select");
		}

		/// <summary>
		/// Kijelöli és háttérszínnel kiemeli az aktív munkalap használt tartományát.
		/// </summary>
		/// <param name="worksheet">A <c>Worksheet</c> COM objektum.</param>
		/// <param name="color">A kívánt háttérszín OLE_COLOR formátumban (pl. BGR int).</param>
		public static void HighlightUsedRange(object worksheet, int color)
		{
			SelectUsedRange(worksheet);
			var usedRange = ComInvoker.GetProperty<object>(worksheet, "UsedRange");
			var interior = ComInvoker.GetProperty<object>(usedRange, "Interior");
			ComInvoker.SetProperty(interior!, "Color", color);
		}

		/// <summary>
		/// Kijelöli a megadott cellacímek által meghatározott tartományokat (pl. "A1", "B2:D4").
		/// Több cím esetén automatikusan összevonja őket (<c>Union</c>).
		/// </summary>
		/// <param name="sheet">A <c>Worksheet</c> COM objektum.</param>
		/// <param name="addresses">A kijelölendő tartományok címei Excel formátumban.</param>
		public static void SelectCells(object sheet, params string[] addresses)
		{
			if (addresses.Length == 0)
				return;

			var app = ComInvoker.GetProperty<object>(sheet, "Application");

			var ranges = addresses
				.Select(addr => ComInvoker.GetProperty<object>(sheet, "Range", new object[] { addr }))
				.ToArray();

			object combined = ranges[0];
			for (int i = 1; i < ranges.Length; i++)
			{
				combined = ComInvoker.CallMethod<object>(app, "Union", combined, ranges[i]);
			}

			ComInvoker.CallMethod(combined, "Select");
		}

		/// <summary>
		/// Lekéri az aktuálisan kijelölt cellák (akár több tartományból) koordinátáit.
		/// </summary>
		/// <param name="excel">Az Excel Application vagy Window COM objektum, amelyből a kijelölt tartomány elérhető.</param>
		/// <returns>A kiválasztott cellák listája sor és oszlop szerint (<c>Row</c>, <c>Column</c>).</returns>
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

		/// <summary>
		/// Lekéri az aktuálisan kijelölt cellák koordinátáit és COM objektumait is.
		/// </summary>
		/// <param name="excel">Az Excel Application vagy Window COM objektum, amelyből a kijelölt tartomány elérhető.</param>
		/// <returns>Lista a kijelölt cellák koordinátáival és COM objektumaival: (<c>Row</c>, <c>Column</c>, <c>Cell</c>).</returns>
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

		/// <summary>
		/// Excel oszlopbetű (pl. "A", "AB") átalakítása sorszámmá (pl. 1, 28).
		/// </summary>
		/// <param name="col">Az oszlop betűjele.</param>
		/// <returns>A numerikus sorszám (1-alapú).</returns>
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
