using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;

namespace ComAutoWrapper
{
	internal class ComAutoHelper
	{
		public static bool TryGetProperty<T>(object comObject, string propertyName, out T? value, params object[]? args)
		{
			value = default;
			try
			{
				object? result = comObject.GetType().InvokeMember(
					propertyName,
					BindingFlags.GetProperty,
					null,
					comObject,
					args ?? Array.Empty<object>());

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


		public static bool PropertyExists(object comObject, string propertyName)
		{
			try
			{
				var disp = comObject as IDispatch;
				if (disp == null) return false;

				var names = new[] { propertyName };
				var dispIds = new int[1];

				Guid dispid = Guid.Empty;
				int hResult = disp.GetIDsOfNames(ref dispid, names, 1, 0, dispIds);
				return hResult == 0; // S_OK
			}
			catch
			{
				return false;
			}
		}
	}
}

