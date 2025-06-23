using System;
using System.Runtime.InteropServices;
using System.Runtime.InteropServices.ComTypes;

namespace ComAutoWrapper
{
	[ComImport]
	[Guid("00020400-0000-0000-C000-000000000046")]
	[InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
	internal interface IDispatch
	{
		[PreserveSig] int GetTypeInfoCount(out int Count);
		[PreserveSig] int GetTypeInfo(int iTInfo, int lcid, out ITypeInfo typeInfo);
		[PreserveSig]
		int GetIDsOfNames(ref Guid riid,
			[MarshalAs(UnmanagedType.LPArray, ArraySubType = UnmanagedType.LPWStr)] string[] rgsNames,
			int cNames, int lcid,
			[MarshalAs(UnmanagedType.LPArray)] int[] rgDispId);
		[PreserveSig]
		int Invoke(int dispIdMember, ref Guid riid, uint lcid, ushort wFlags,
			ref DISPPARAMS pDispParams, out object pVarResult,
			ref EXCEPINFO pExcepInfo, IntPtr[] pArgErr);
	}
}
