using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface
{
	public static class Util
	{

		/// <summary>
		/// Accepts a type and a string value and attempts to do a proper type conversion
		/// </summary>
		/// <param name="typeHint"></param>
		/// <param name="inValue"></param>
		/// <returns></returns>
		public static string ApplyTypeHint(Type typeHint, object inValue)
		{
			string outValue = null;
			if (typeHint != null && inValue != null)
			{
				// Type hint: DateTime
				if (typeHint == typeof(DateTime) && inValue is double)
				{
					outValue = DateTime.FromOADate((double)inValue).ToString();
				}
			}

			if (outValue == null)
			{
				outValue = inValue?.ToString();
			}
			return outValue;
		}
	}
}
