using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ComAutoWrapper
{
	/// <summary>
	/// A szabványos COM <c>IDispatch</c> interfész alacsony szintű leképezése.
	/// Lehetővé teszi késői kötésű tagelérést és típusinformációk elérését <c>ITypeInfo</c> segítségével.
	/// </summary>
	[ComImport]
	[Guid("00020400-0000-0000-C000-000000000046")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	internal interface IDispatch
	{
		/// <summary>
		/// Lekérdezi, hogy a COM objektum mennyi típusinformációval rendelkezik (0 vagy 1).
		/// </summary>
		/// <param name="Count">A típusinformációk száma.</param>
		/// <returns>HRESULT kód (0 = S_OK).</returns>
		[PreserveSig]
		int GetTypeInfoCount(out int Count);

		/// <summary>
		/// Lekéri az adott típusinformációt (<see cref="ITypeInfo"/>).
		/// </summary>
		/// <param name="iTInfo">A kért típusinformáció indexe (általában 0).</param>
		/// <param name="lcid">A nyelvi azonosító (pl. 1033 = en-US).</param>
		/// <param name="typeInfo">Az eredményül kapott <see cref="ITypeInfo"/> objektum.</param>
		/// <returns>HRESULT kód (0 = S_OK).</returns>
		[PreserveSig]
		int GetTypeInfo(int iTInfo, int lcid, out ITypeInfo typeInfo);

		/// <summary>
		/// Leképezi a tagneveket diszpidhívásokhoz használható azonosítókká (<c>DispId</c>).
		/// </summary>
		/// <param name="riid">Mindig <see cref="Guid.Empty"/>.</param>
		/// <param name="rgsNames">A lekérdezendő tagnevek tömbje.</param>
		/// <param name="cNames">A nevek száma.</param>
		/// <param name="lcid">A nyelvi azonosító.</param>
		/// <param name="rgDispId">A visszatérő azonosítók tömbje.</param>
		/// <returns>HRESULT kód (0 = S_OK).</returns>
		[PreserveSig]
		int GetIDsOfNames(
			ref Guid riid,
			[MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgsNames,
			int cNames,
			int lcid,
			[MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);

		/// <summary>
		/// Meghív egy metódust vagy property-t a <c>DispId</c> alapján.
		/// Ez az alacsony szintű belső hívás a késői kötésű COM elérés alapja.
		/// </summary>
		/// <param name="dispIdMember">A meghívandó tag DispId azonosítója.</param>
		/// <param name="riid">Mindig <see cref="Guid.Empty"/>.</param>
		/// <param name="lcid">A nyelvi azonosító.</param>
		/// <param name="wFlags">A hívás típusa (<c>DISPATCH_METHOD</c>, <c>DISPATCH_PROPERTYGET</c>, stb.).</param>
		/// <param name="pDispParams">A híváshoz használt paraméterek.</param>
		/// <param name="pVarResult">A visszatérési érték.</param>
		/// <param name="pExcepInfo">Kivételinformáció, ha hiba történik.</param>
		/// <param name="pArgErr">Hibás argumentum indexek.</param>
		/// <returns>HRESULT kód (0 = S_OK).</returns>
		[PreserveSig]
		int Invoke(
			int dispIdMember,
			ref Guid riid,
			uint lcid,
			ushort wFlags,
			ref DISPPARAMS pDispParams,
			out object pVarResult,
			ref EXCEPINFO pExcepInfo,
			IntPtr[] pArgErr);
	}
}
