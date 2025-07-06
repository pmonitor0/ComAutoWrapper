using System;
using System.Reflection;
using System.Runtime.InteropServices;

namespace ComAutoWrapper
{
	public static class ComInvoker
	{
		public static T? GetProperty<T>(object comObject, string propertyName)
			=> GetProperty<T>(comObject, propertyName, null);

		public static T? GetProperty<T>(object comObject, string propertyName, object[]? parameters)
		{
			try
			{
				object? result = comObject.GetType().InvokeMember(
					propertyName,
					BindingFlags.GetProperty,
					null,
					comObject,
					parameters);

				return result is T typed ? typed : default;
			}
			catch (TargetInvocationException tie)
			{
				ThrowComException(propertyName, tie);
				return default;
			}
		}

		public static void SetProperty(object comObject, string propertyName, object value)
		{
			SetProperty(comObject, propertyName, new object[] { value });
		}

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

		public static T? CallMethod<T>(object comObject, string methodName, params object[] parameters)
		{
			try
			{
				object? result = comObject.GetType().InvokeMember(
					methodName,
					BindingFlags.InvokeMethod,
					null,
					comObject,
					parameters);

				return result is T typed ? typed : default;
			}
			catch (TargetInvocationException tie)
			{
				ThrowComException(methodName, tie);
				return default;
			}
		}

		public static List<string> ListCallableMembers(object comObject)
        {
            var type = comObject.GetType();
            var members = type.GetMembers(BindingFlags.Public | BindingFlags.Instance);
            return members.Select(m => $"{m.MemberType}: {m.Name}").Distinct().OrderBy(s => s).ToList();
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