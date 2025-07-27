using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;

namespace ComAutoWrapper
{
	/// <summary>
	/// Segédosztály COM objektumok manuális nyomon követéséhez és felszabadításához.
	/// </summary>
	public static class ComReleaseHelper
	{
		private static readonly List<object> _tracked = new();

		/// <summary>
		/// Hozzáad egy COM objektumot a nyomon követett példányok listájához.
		/// Azonos példányt nem ad hozzá újra (referencia szerint vizsgál).
		/// </summary>
		/// <typeparam name="T">A COM objektum típusa.</typeparam>
		/// <param name="comObject">A COM objektum, amelyet nyomon követünk.</param>
		/// <returns>Ugyanaz a COM objektum, változtatás nélkül.</returns>
		public static T Track<T>(T comObject)
		{
			if (comObject == null || !Marshal.IsComObject(comObject))
				return comObject;

			// Csak akkor adjuk hozzá, ha nincs még benne (referencia alapon)
			if (!_tracked.Any(o => ReferenceEquals(o, comObject)))
			{
				_tracked.Add(comObject!);
			}
			
			return comObject;
		}

		/// <summary>
		/// Felszabadítja az összes nyomon követett COM objektumot a <see cref="Marshal.FinalReleaseComObject"/> segítségével.
		/// Sikertelen felszabadítás esetén a kivétel elnyelésre kerül.
		/// </summary>
		public static void ReleaseAll()
		{
#pragma warning disable CA1416
			foreach (object obj in _tracked)
			{
				try
				{
					Marshal.FinalReleaseComObject(obj!);
				}
				catch
				{
					// Elnyelhető, nem kritikus
				}
			}
			_tracked.Clear();
		}

		/// <summary>
		/// Törli a nyomon követett objektumok listáját anélkül, hogy felszabadítaná őket.
		/// </summary>
		public static void Clear() => _tracked.Clear();

		/// <summary>
		/// Visszaadja a nyomon követett COM objektumok számát.
		/// </summary>
		public static int Count => _tracked.Count;

		/// <summary>
		/// Kiírja a Console-ra az összes nyomon követett COM objektum típusát.
		/// Segítséget nyújt a hibakereséshez és fejlesztéshez.
		/// </summary>
		public static void DebugList()
		{
			foreach (var o in _tracked)
			{
				Console.WriteLine($"Tracked: {o?.GetType()}");
			}
		}

		/// <summary>
		/// Eltávolít egy adott COM objektumot a követett listából, ha benne van.
		/// </summary>
		/// <param name="comObject">A COM objektum, amelyet törölni szeretnél.</param>
		/// <returns><c>true</c>, ha sikerült eltávolítani.</returns>
		public static bool Remove(object comObject)
		{
			return _tracked.Remove(comObject);
		}

		/// <summary>
		/// Felszabadítja az összes COM objektumot, majd kiüríti a listát.
		/// </summary>
		public static void Reset()
		{
			ReleaseAll();
			_tracked.Clear();
		}
		/// <summary>
		/// Megvizsgálja, hogy a megadott COM objektum jelen van-e a nyilvántartásban.
		/// </summary>
		public static bool IsTracked(object comObject)
		{
			return _tracked.Contains(comObject);
		}
	}
}
