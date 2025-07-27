using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace ComAutoWrapper
{
	/// <summary>
	/// Magas szintű segédfüggvények COM objektumok tulajdonságainak biztonságos lekérdezéséhez.
	/// </summary>
	public class ComAutoHelper
	{
		/// <summary>
		/// Megpróbál lekérni egy property értéket a megadott COM objektumtól, paraméterek nélkül.
		/// A hívás nem dob kivételt, sikertelenség esetén false értékkel tér vissza.
		/// </summary>
		/// <typeparam name="T">A visszatérési érték típusa.</typeparam>
		/// <param name="comObject">A COM objektum, amelytől a property-t le szeretnénk kérni.</param>
		/// <param name="propertyName">A lekérdezendő property neve.</param>
		/// <param name="value">A lekért érték, ha sikeres a hívás; egyébként a típus alapértelmezett értéke.</param>
		/// <returns><c>true</c>, ha a lekérés sikeres volt és az érték típuskompatibilis; különben <c>false</c>.</returns>
		public static bool TryGetProperty<T>(object comObject, string propertyName, out T? value)
			=> TryGetProperty(comObject, propertyName, out value, null);

		/// <summary>
		/// Megpróbál lekérni egy property értéket a megadott COM objektumtól, opcionális paraméterekkel.
		/// A hívás nem dob kivételt, sikertelenség esetén false értékkel tér vissza.
		/// </summary>
		/// <typeparam name="T">A visszatérési érték típusa.</typeparam>
		/// <param name="comObject">A COM objektum, amelytől a property-t le szeretnénk kérni.</param>
		/// <param name="propertyName">A lekérdezendő property neve.</param>
		/// <param name="value">A lekért érték, ha sikeres a hívás; egyébként a típus alapértelmezett értéke.</param>
		/// <param name="args">Opcionális paraméterek (pl. indexelt property-khez).</param>
		/// <returns><c>true</c>, ha a lekérés sikeres volt és az érték típuskompatibilis; különben <c>false</c>.</returns>
		public static bool TryGetProperty<T>(object comObject, string propertyName, out T? value, params object[]? args)
		{
			value = default;
			try
			{
				object? result = ComReleaseHelper.Track(comObject.GetType().InvokeMember(
					propertyName,
					BindingFlags.GetProperty,
					null,
					comObject,
					args ?? Array.Empty<object>()));

				if (result != null && Marshal.IsComObject(result))
					ComReleaseHelper.Track(result);

				if (result is T casted)
				{
					value = casted;
					return true;
				}

				return false;
			}
			catch
			{
				return false;
			}
		}

		/// <summary>
		/// Megvizsgálja, hogy létezik-e a megadott property a COM objektumon.
		/// A vizsgálat nem dob kivételt, ha nem sikerül, false értékkel tér vissza.
		/// </summary>
		/// <param name="comObject">A vizsgálandó COM objektum.</param>
		/// <param name="propertyName">A keresett property neve.</param>
		/// <returns><c>true</c>, ha a property elérhető; különben <c>false</c>.</returns>
		public static bool PropertyExists(object comObject, string propertyName)
		{
			try
			{
				if (comObject is not IDispatch disp)
					return false;

				var names = new[] { propertyName };
				var dispIds = new int[1];
				Guid guid = Guid.Empty;

				int hResult = disp.GetIDsOfNames(ref guid, names, 1, 0, dispIds);
				return hResult == 0; // S_OK
			}
			catch
			{
				return false;
			}
		}

		/// <summary>
		/// Lekéri az Excel.Application COM objektumhoz tartozó Windows folyamatot (Process).
		/// </summary>
		/// <param name="excelApp">Az Excel COM objektum.</param>
		/// <returns>A hozzá tartozó Process példány.</returns>
		/// <exception cref="InvalidOperationException">
		/// Ha nem sikerül lekérni az ablak handle-t vagy a folyamatazonosítót.
		/// </exception>
		/// <example>
		/// using System.Diagnostics;
		/// ...
		/// var proc = ComAutoHelper.GetProcessByExcelHandle(excelApp);
		/// Console.WriteLine("Excel PID: " + proc.Id);
		/// </example>
		/// <returns>
		///<c>true</c>, ha a lekérés sikeres volt és az érték típuskompatibilis; különben<c>false</c>.
		///</returns>
		public static Process? GetProcessByExcelHandle(object excelApp)
		{
			int hwnd = ComInvoker.GetProperty<int>(excelApp!, "Hwnd", null);
			if (hwnd == 0)
				throw new InvalidOperationException("Could not retrieve Excel window handle.");

			GetWindowThreadProcessId(hwnd, out nint processID);
			if (processID == 0)
				throw new InvalidOperationException("Could not retrieve Excel process ID.");

			return Process.GetProcessById(processID.ToInt32());
		}

		[System.Runtime.InteropServices.DllImport("user32.dll")]
		private static extern uint GetWindowThreadProcessId(int hWnd, out nint lpdwProcessId);

	}
}
