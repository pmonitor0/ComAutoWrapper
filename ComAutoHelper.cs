using System;
using System.Collections.Generic;
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
	}
}
