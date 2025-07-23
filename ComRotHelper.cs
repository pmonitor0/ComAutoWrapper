using Microsoft.VisualBasic; // kell a Information.TypeName-hez
using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ComAutoWrapper
{
	/// <summary>
	/// Segédosztály a Running Object Table (ROT) vizsgálatához és Excel alkalmazások detektálásához.
	/// </summary>
	public static class ComRotHelper
	{
		/// <summary>
		/// Lekéri a rendszerben futó Excel alkalmazás példányokat a Running Object Table (ROT) alapján.
		/// A metódus olyan Workbook objektumokat keres, amelyek COM interfészen keresztül elérhetőek,
		/// majd ezekből kinyeri a hozzájuk tartozó Application objektumot.
		/// </summary>
		/// <returns>A detektált Excel Application COM objektumok listája.</returns>
		public static List<object> GetExcelApplications()
		{
			var result = new List<object>();

			if (GetRunningObjectTable(0, out IRunningObjectTable? rot) != 0 || rot == null)
				return result;

			rot.EnumRunning(out IEnumMoniker? enumMoniker);
			enumMoniker.Reset();

			var monikers = new IMoniker[1];
			IntPtr fetched = IntPtr.Zero;

			while (enumMoniker.Next(1, monikers, fetched) == 0)
			{
				CreateBindCtx(0, out IBindCtx? bindCtx);

				try
				{
					rot.GetObject(monikers[0], out object? comObject);
					if (comObject == null)
						continue;

					// Csak Workbook típusokra figyelünk
					if (Information.TypeName(comObject) == "Workbook")
					{
						var workbook = comObject;

						// Lekérjük a parent Application objektumot
						var app = workbook.GetType().InvokeMember(
							"Parent",
							System.Reflection.BindingFlags.GetProperty,
							null,
							workbook,
							null);

						// Csak akkor adjuk hozzá, ha még nem szerepel a listában (referencia szerint)
						if (!result.Any(o => ReferenceEquals(o, app)))
						{
							result.Add(app);
						}
					}
				}
				catch
				{
					// hibás vagy elérhetetlen objektum – elnyelhető
				}
			}

			return result;
		}

		/// <summary>
		/// Meghívja az <c>ole32.dll</c> <c>GetRunningObjectTable</c> API-ját, amely elérhetővé teszi a futó COM objektumokat.
		/// </summary>
		/// <param name="reserved">Mindig 0.</param>
		/// <param name="prot">A visszaadott <see cref="IRunningObjectTable"/> példány, ha sikeres.</param>
		/// <returns>0, ha sikeres (S_OK); különben hibakód.</returns>
		[DllImport("ole32.dll")]
		private static extern int GetRunningObjectTable(int reserved, out IRunningObjectTable? prot);

		/// <summary>
		/// Meghívja az <c>ole32.dll</c> <c>CreateBindCtx</c> API-ját, amely létrehoz egy bind kontextust.
		/// </summary>
		/// <param name="reserved">Mindig 0.</param>
		/// <param name="ppbc">A visszaadott <see cref="IBindCtx"/> példány, ha sikeres.</param>
		/// <returns>0, ha sikeres (S_OK); különben hibakód.</returns>
		[DllImport("ole32.dll")]
		private static extern int CreateBindCtx(int reserved, out IBindCtx? ppbc);
	}
}
