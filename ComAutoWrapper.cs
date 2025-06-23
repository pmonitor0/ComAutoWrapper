using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ComAutoWrapper
{
	public static class ComInvoker
	{
		public static T? GetProperty<T>(object comObject, string propertyName)
		{
			try
			{
				object? result = comObject.GetType().InvokeMember(
					propertyName,
					BindingFlags.GetProperty,
					null,
					comObject,
					null);

				return (result is T typed) ? typed : default;
			}
			catch (TargetInvocationException tie)
			{
				ThrowComException(propertyName, tie);
				return default;
			}
		}

		public static bool SetProperty(object comObject, string propertyName, object value)
		{
			try
			{
				comObject.GetType().InvokeMember(
					propertyName,
					BindingFlags.SetProperty,
					null,
					comObject,
					new object[] { value });
				return true;
			}
			catch (TargetInvocationException tie)
			{
				ThrowComException(propertyName, tie);
				return false;
			}
		}

		public static object? CallMethod(object comObject, string methodName, params object[] args)
		{
			try
			{
				return comObject.GetType().InvokeMember(
					methodName,
					BindingFlags.InvokeMethod,
					null,
					comObject,
					args);
			}
			catch (TargetInvocationException tie)
			{
				ThrowComException(methodName, tie);
				return null;
			}
		}

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