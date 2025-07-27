using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ComAutoWrapper
{
	/// <summary>
	/// Mintaosztály, amely bemutatja, hogyan lehet Word dokumentumba táblázatot beszúrni és formázni COM automatizálással.
	/// </summary>
	public class WordHelper
	{
		/// <summary>
		/// Word alkalmazás indítása → új dokumentum létrehozása → 3x3-as táblázat beszúrása → cellák kitöltése és fejléc formázása.
		/// A végén a dokumentum bezáródik mentés nélkül, az alkalmazás pedig kilép.
		/// </summary>
		public static void RunWordInsertTableDemo()
		{
#pragma warning disable CA1416
			Type? type = Type.GetTypeFromProgID("Word.Application");
			object? wordApp = Activator.CreateInstance(type!);
			ComInvoker.SetProperty(wordApp!, "Visible", true);
			ComInvoker.SetProperty(wordApp!, "DisplayAlerts", false);

			var documents = ComInvoker.GetProperty<object>(wordApp!, "Documents");
			var doc = ComInvoker.CallMethod<object>(documents!, "Add");

			var range = ComInvoker.GetProperty<object>(doc!, "Content");

			// Táblázat beszúrása: 3 sor, 3 oszlop
			var tables = ComInvoker.GetProperty<object>(doc!, "Tables");
			var table = ComInvoker.CallMethod<object>(tables!, "Add", range!, 3, 3);

			// Cellák feltöltése
			for (int row = 1; row <= 3; row++)
			{
				for (int col = 1; col <= 3; col++)
				{
					var cell = ComInvoker.CallMethod<object>(table!, "Cell", row, col);
					var cellRange = ComInvoker.GetProperty<object>(cell!, "Range");
					ComInvoker.SetProperty(cellRange!, "Text", $"R{row}C{col}");

					if (row == 1)
					{
						// Fejléc formázása
						WordStyleHelper.ApplyStyle(
							cellRange!,
							fontColor: ComValueConverter.ToOleColor(Color.White),
							backgroundColor: ComValueConverter.ToOleColor(Color.DarkRed),
							bold: true
						);
					}

					ComReleaseHelper.Track(cellRange);
					ComReleaseHelper.Track(cell);
				}
			}

			Console.WriteLine("Táblázat beszúrva és formázva.");
			Console.ReadKey(true);

			// Bezárás mentés nélkül
			ComInvoker.SetProperty(doc!, "Saved", ComValueConverter.ToComBool(true));
			ComInvoker.CallMethod(doc!, "Close", ComValueConverter.ToComBool(false));
			ComInvoker.CallMethod(wordApp!, "Quit");

			ComReleaseHelper.Track(table);
			ComReleaseHelper.Track(tables);
			ComReleaseHelper.Track(range);
			ComReleaseHelper.Track(doc);
			ComReleaseHelper.Track(documents);
			ComReleaseHelper.Track(wordApp);
			ComReleaseHelper.ReleaseAll();
		}
	}
}
