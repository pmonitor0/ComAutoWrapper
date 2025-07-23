using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ComAutoWrapper
{
	/// <summary>
	/// Segédosztály COM objektumok property-jeinek és metódusainak dinamikus eléréséhez.
	/// </summary>
	public static class ComInvoker
	{
		/// <summary>
		/// Lekér egy property értéket a megadott COM objektumtól, paraméterek nélkül.
		/// </summary>
		/// <typeparam name="T">A visszatérési érték típusa.</typeparam>
		/// <param name="comObject">A COM objektum, amelytől a property-t le szeretnénk kérni.</param>
		/// <param name="propertyName">A lekérdezendő property neve.</param>
		/// <returns>A property értéke, vagy <c>default(T)</c>, ha a lekérés sikertelen vagy nem konvertálható.</returns>
		public static T? GetProperty<T>(object comObject, string propertyName)
			=> GetProperty<T>(comObject, propertyName, null);

		/// <summary>
		/// Lekér egy property értéket a megadott COM objektumtól, opcionális paraméterekkel (pl. indexelt property).
		/// </summary>
		/// <typeparam name="T">A visszatérési érték típusa.</typeparam>
		/// <param name="comObject">A COM objektum, amelytől a property-t le szeretnénk kérni.</param>
		/// <param name="propertyName">A lekérdezendő property neve.</param>
		/// <param name="parameters">Opcionális paraméterek, például indexelt property-k esetén.</param>
		/// <returns>A property értéke, vagy <c>default(T)</c>, ha a lekérés sikertelen vagy nem konvertálható.</returns>
		/// <exception cref="InvalidOperationException">Ha COM kivétel történik a property elérésekor.</exception>
		public static T? GetProperty<T>(object comObject, string propertyName, object[]? parameters)
		{
			try
			{
				object? result = ComReleaseHelper.Track(comObject.GetType().InvokeMember(
					propertyName,
					BindingFlags.GetProperty,
					null,
					comObject,
					parameters));

				if (result is T typed)
					return typed;

				if (typeof(T) == typeof(object))
					return (T?)result;

				return default;
			}
			catch (TargetInvocationException tie)
			{
				ThrowComException(propertyName, tie);
				return default;
			}
		}

		/// <summary>
		/// Beállítja egy COM objektum property-jének értékét.
		/// </summary>
		/// <param name="comObject">A COM objektum, amelynek a property-jét be szeretnénk állítani.</param>
		/// <param name="propertyName">A beállítandó property neve.</param>
		/// <param name="value">A beállítandó érték.</param>
		public static void SetProperty(object comObject, string propertyName, object value)
		{
			SetProperty(comObject, propertyName, new object[] { value });
		}

		/// <summary>
		/// Beállítja egy COM objektum property-jének értékét tetszőleges paraméterlistával.
		/// </summary>
		/// <param name="comObject">A COM objektum, amelynek a property-jét be szeretnénk állítani.</param>
		/// <param name="propertyName">A beállítandó property neve.</param>
		/// <param name="parameters">A property beállításához használt paraméterek (pl. index és érték).</param>
		/// <exception cref="InvalidOperationException">Ha COM kivétel történik a property beállításakor.</exception>
		public static void SetProperty(object comObject, string propertyName, params object[] parameters)
		{
			try
			{
				comObject.GetType().InvokeMember(
					propertyName,
					BindingFlags.SetProperty,
					null,
					comObject,
					parameters);
			}
			catch (TargetInvocationException tie)
			{
				ThrowComException(propertyName, tie);
			}
		}

		/// <summary>
		/// Meghív egy metódust a megadott COM objektumon.
		/// </summary>
		/// <param name="comObject">A COM objektum, amelyen a metódust hívni szeretnénk.</param>
		/// <param name="methodName">A meghívandó metódus neve.</param>
		/// <param name="args">A metódushoz tartozó argumentumok.</param>
		/// <returns>A visszatérési érték (ha van), vagy <c>null</c>, ha a hívás sikertelen.</returns>
		/// <exception cref="InvalidOperationException">Ha COM kivétel történik a metódus hívásakor.</exception>
		public static object? CallMethod(object comObject, string methodName, params object[] args)
		{
			try
			{
				return ComReleaseHelper.Track(comObject.GetType().InvokeMember(
					methodName,
					BindingFlags.InvokeMethod,
					null,
					comObject,
					args));
			}
			catch (TargetInvocationException tie)
			{
				ThrowComException(methodName, tie);
				return null;
			}
		}

		/// <summary>
		/// Meghív egy metódust a megadott COM objektumon, és a visszatérési értéket a megadott típusra castolja.
		/// </summary>
		/// <typeparam name="T">A várt visszatérési típus.</typeparam>
		/// <param name="comObject">A COM objektum, amelyen a metódust hívni szeretnénk.</param>
		/// <param name="methodName">A meghívandó metódus neve.</param>
		/// <param name="parameters">A metódushoz tartozó argumentumok.</param>
		/// <returns>A visszatérési érték típusként, vagy <c>default(T)</c>, ha a hívás sikertelen.</returns>
		/// <exception cref="InvalidOperationException">Ha COM kivétel történik a metódus hívásakor.</exception>
		public static T? CallMethod<T>(object comObject, string methodName, params object[] parameters)
		{
			try
			{
				object? result = ComReleaseHelper.Track(comObject.GetType().InvokeMember(
					methodName,
					BindingFlags.InvokeMethod,
					null,
					comObject,
					parameters));

				return result is T typed ? typed : default;
			}
			catch (TargetInvocationException tie)
			{
				ThrowComException(methodName, tie);
				return default;
			}
		}

		/// <summary>
		/// Lekér egy listát a COM objektum publikus elérhető metódusairól és property-jeiről.
		/// </summary>
		/// <param name="comObject">A COM objektum, amelynek a tagjait listázni szeretnénk.</param>
		/// <returns>A tagok neveinek listája típusmegjelöléssel (pl. "Property: Name").</returns>
		public static List<string> ListCallableMembers(object comObject)
		{
			var type = comObject.GetType();
			var members = type.GetMembers(BindingFlags.Public | BindingFlags.Instance);
			return members.Select(m => $"{m.MemberType}: {m.Name}").Distinct().OrderBy(s => s).ToList();
		}

		/// <summary>
		/// Kivételt dob célzott COM hiba esetén, amely kinyeri a belső <see cref="COMException"/> információt.
		/// </summary>
		/// <param name="memberName">A hívott property vagy metódus neve.</param>
		/// <param name="tie">A kivétel, amelyet az InvokeMember hívás dobott.</param>
		private static void ThrowComException(string memberName, TargetInvocationException tie)
		{
			if (tie.InnerException is COMException comEx)
			{
				string msg = $"COM hiba a(z) '{memberName}' tag elérésekor. HRESULT: 0x{comEx.HResult:X8}";
				throw new InvalidOperationException(msg, comEx);
			}
			throw tie;
		}
	}
}
