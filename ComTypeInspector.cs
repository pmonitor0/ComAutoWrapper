using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ComAutoWrapper
{
	/// <summary>
	/// Segédosztály COM objektumok típusinformációinak és tagjainak introspekciójához.
	/// Az <see cref="IDispatch"/> és <see cref="ITypeInfo"/> interfészeken keresztül működik.
	/// </summary>
	public static class ComTypeInspector
	{
		/// <summary>
		/// Lekéri a COM objektum típusának nevét az <c>ITypeInfo.GetDocumentation</c> alapján.
		/// </summary>
		/// <param name="comObject">A COM objektum, amelynek a típusnevét le szeretnénk kérni.</param>
		/// <returns>A típus neve (általában az interfész neve), vagy <c>null</c>, ha nem elérhető.</returns>
		public static string? GetTypeName(object comObject)
		{
			if (comObject is not IDispatch dispatch)
				return null;

			dispatch.GetTypeInfo(0, 1033, out var typeInfo);

			typeInfo.GetDocumentation(-1, out var name, out _, out _, out _);

			return name.TrimStart('_');
		}

		/// <summary>
		/// Lekéri a COM objektum összes elérhető tagját három kategóriában:
		/// metódusok, olvasható property-k, és írható property-k.
		/// </summary>
		/// <param name="comObject">A vizsgálandó COM objektum.</param>
		/// <returns>
		/// Egy tuple három listával:
		/// <list type="bullet">
		/// <item><description><c>Methods</c>: elérhető metódusok nevei</description></item>
		/// <item><description><c>PropertyGets</c>: olvasható property-k nevei</description></item>
		/// <item><description><c>PropertySets</c>: írható property-k nevei</description></item>
		/// </list>
		/// </returns>
		public static (List<string> Methods, List<string> PropertyGets, List<string> PropertySets)
			ListMembers(object comObject)
		{
			var methods = new List<string>();
			var gets = new List<string>();
			var sets = new List<string>();

			if (comObject is not IDispatch dispatch)
				return (methods, gets, sets);

			dispatch.GetTypeInfo(0, 0, out var typeInfo);
			typeInfo.GetTypeAttr(out var pTypeAttr);
			var typeAttr = Marshal.PtrToStructure<TYPEATTR>(pTypeAttr)!;

			for (int i = 0; i < typeAttr.cFuncs; i++)
			{
				typeInfo.GetFuncDesc(i, out var pFuncDesc);
				var funcDesc = Marshal.PtrToStructure<FUNCDESC>(pFuncDesc)!;

				typeInfo.GetDocumentation(funcDesc.memid, out var name, out _, out _, out _);

				switch (funcDesc.invkind)
				{
					case INVOKEKIND.INVOKE_FUNC:
						methods.Add(name);
						break;
					case INVOKEKIND.INVOKE_PROPERTYGET:
						gets.Add(name);
						break;
					case INVOKEKIND.INVOKE_PROPERTYPUT:
					case INVOKEKIND.INVOKE_PROPERTYPUTREF:
						sets.Add(name);
						break;
				}

				typeInfo.ReleaseFuncDesc(pFuncDesc);
			}

			typeInfo.ReleaseTypeAttr(pTypeAttr);
			return (methods, gets, sets);
		}
	}
}
