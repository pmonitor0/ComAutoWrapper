using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ComAutoWrapper
{
	public static class ComTypeInspector
	{
		public static string? GetTypeName(object comObject)
		{
			if (comObject is not IDispatch dispatch)
				return null;

			dispatch.GetTypeInfo(0, 1033, out var typeInfo);

			typeInfo.GetDocumentation(-1, out var name, out _, out _, out _);

			return name.TrimStart('_');
		}

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
