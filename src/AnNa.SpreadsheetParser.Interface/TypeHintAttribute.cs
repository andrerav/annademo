using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using AnNa.SpreadsheetParser.Interface.Sheets;

namespace AnNa.SpreadsheetParser.Interface
{
	[AttributeUsage(AttributeTargets.Field)]
	public class TypeHintAttribute : System.Attribute
	{
		private Type _typeHint;

		public TypeHintAttribute(Type typeHint)
		{
			_typeHint = typeHint;
		}

		public Type TypeHint
		{
			get { return _typeHint; }
			set { _typeHint = value; }
		}
	}

	public static class StringExtensions
	{
		public static Type GetTypeHint(this string subject, ISheetSpecification sheet)
		{
			foreach (var nestedType in sheet.GetType().GetNestedTypes())
			{
				if (typeof(ISheetColumns).IsAssignableFrom(nestedType))
				{
					var fieldInfos = nestedType.GetFields(BindingFlags.Public | BindingFlags.Static | BindingFlags.FlattenHierarchy);
					var attr = fieldInfos
						.Where(x => x.IsLiteral && !x.IsInitOnly && x.GetRawConstantValue().ToString() == subject)
						.Select(x => x.GetCustomAttributes(typeof (TypeHintAttribute), false)
							.Cast<TypeHintAttribute>().FirstOrDefault()).FirstOrDefault();
					if (attr != null)
					{
						return attr.TypeHint;
					}
				}
			}

			return null;
		} 
		
	}
}
