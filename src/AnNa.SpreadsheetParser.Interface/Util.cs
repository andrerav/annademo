using AnNa.SpreadsheetParser.Interface.Sheets;
using FastMember;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface
{
	public static class Util
	{

		public static object ApplyTypeHint<T>(object inValue, out object convertedValue)
		{
			SyntaxError dummySyntaxError;
			return ApplyTypeHint(typeof(T), inValue, out convertedValue, out dummySyntaxError);
		}

		public static object ApplyTypeHint(Type typeHint, object inValue, out object convertedValue)
		{
			SyntaxError dummySyntaxError;
			return ApplyTypeHint(typeHint, inValue, out convertedValue, out dummySyntaxError);
		}

		/// <summary>
		/// Accepts a type and a string value and attempts to do a proper type conversion
		/// </summary>
		/// <param name="inValue"></param>
		/// <returns>A simple ToString() representation of the converted value, or the input value if conversion was not possible or necessary</returns>
		public static object ApplyTypeHint(Type typeHint, object inValue, out object convertedValue, out SyntaxError syntaxError)
		{
			syntaxError = null; 
			object outValue = null;
			if (inValue != null)
			{
				// Type hint: DateTime
				if (typeHint == typeof(DateTime) || typeHint == typeof(DateTime?))
				{
					if (inValue is double)
					{
						outValue = DateTime.FromOADate((double)inValue);
					}
					else
					{
						double tmp;
						if (double.TryParse(inValue.ToString(), out tmp))
						{
							outValue = DateTime.FromOADate(tmp);
						}
						else
						{
							if (inValue is DateTime)
							{
								outValue = inValue;
							}
							else
							{
								DateTime dtmp;
								if (DateTime.TryParse(inValue.ToString(), out dtmp))
								{
									outValue = dtmp;
								}
								else
								{
									// If the type isn't nullable and the value is not null then add a syntax error
									if ((typeHint != typeof(DateTime?) && (inValue == null 
										|| !string.IsNullOrWhiteSpace(inValue.ToString())))
										
										|| (typeHint == typeof(DateTime?) && (inValue != null && !string.IsNullOrWhiteSpace(inValue.ToString())))
										)
									{
										syntaxError = CreateSyntaxError(typeHint, inValue);
									}
								}
							}
						}
					}
				}

				else if (typeHint == typeof(double) || typeHint == typeof(double?))
				{
					if (inValue is double)
					{
						outValue = inValue;
					}
					else
					{
						double tmp;
						if (double.TryParse(inValue.ToString(), out tmp))
						{
							outValue = tmp;
						}
						else
						{
							if (typeHint != typeof(double?) && (inValue == null
							|| !string.IsNullOrWhiteSpace(inValue.ToString())))
							{
								syntaxError = CreateSyntaxError(typeHint, inValue);
							}

							outValue = default(double);
						}
					}
				}

				else if (typeHint == typeof(string))
				{
					outValue = inValue.ToString();
				}

				else // Type hint not explicitly understood, attempt generic type conversion
				{
					try {
						if (typeHint.IsGenericType && typeHint.GetGenericTypeDefinition() == typeof(Nullable<>))
						{
							if (inValue != null && String.IsNullOrWhiteSpace(inValue.ToString()))
							{
								outValue = null;
							}
							else
							{
								outValue = Convert.ChangeType(inValue, typeHint.GetGenericArguments().First());
							}
						}
						else
						{
							outValue = Convert.ChangeType(inValue, typeHint);
						}
					}
					catch
					{
						syntaxError = CreateSyntaxError(typeHint, inValue);
					}
				}
			}
			else
			{
				convertedValue = null;
				return null;
			}

			convertedValue = outValue;

			if (outValue == null)
			{
				return inValue;
			}
			else
			{
				return outValue;
			}
		}

		private static SyntaxError CreateSyntaxError(Type typeHint, object inValue)
		{
			SyntaxError syntaxError = new SyntaxError();
			syntaxError.TypeHint = typeHint;
			syntaxError.RawValue = inValue.ToString();
			return syntaxError;
		}




		public static void CreateDirectoryIfNotExists(string path)
		{
			var directoryName = new FileInfo(path).DirectoryName;
			if (!File.Exists(directoryName))
			{
				Directory.CreateDirectory(directoryName);
			}
		}

		public static void RemoveEmptyRows(List<Dictionary<string, string>> result)
		{
			var tmp = result.ToList();
			tmp.Reverse();
			int i = tmp.Count() - 1;
			foreach (var row in tmp)
			{
				bool rowHasData = false;
				foreach (var key in row.Keys)
				{
					if (key.ToLower() == "number" || key == "__RowIndex")
					{
						continue;
					}
					if (!string.IsNullOrWhiteSpace(row[key]))
					{
						rowHasData = true;
						break;
					}
				}

				if (!rowHasData)
				{
					result.RemoveAt(i);
				}

				i--;
			}
		}

		/// <summary>
		/// Removes empty rows, ie. rows that does not have any values in non-ignored columns
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="result"></param>
		/// <param name="columns"></param>
		public static void RemoveEmptyRows<T>(List<T> result, List<SheetColumn> columns) where T : class, ISheetRow
		{
			result.RemoveAll(row => IsEmpty(row, columns));
		}

		/// <summary>
		/// Returns true if the given row does not have any values in non-ignored columns
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="row"></param>
		/// <param name="columns"></param>
		/// <returns></returns>
		public static bool IsEmpty<T>(T row, List<SheetColumn> columns) where T : ISheetRow
		{
			var accessor = ObjectAccessor.Create(row);

			foreach (var column in columns.Where(c => !c.Ignorable))
			{
				var value = accessor[column.FieldName];
				if (value != null && !value.Equals(GetDefault(column.FieldType)))
				{
					return false;
				}
			}
			return true;
		}

		/// <summary>
		/// Retrieve a list of column definitions from a given sheet
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheet"></param>
		/// <returns></returns>
		public static List<SheetColumn> GetColumns<T>(ITypedSheetWithBulkData<T> sheet) where T : ISheetRow
		{

			return GetColumns(typeof(T));

		}

		public static List<SheetColumn> GetColumns(SheetRow row)
		{
			return GetColumns(row.GetType());
		}

		public static List<SheetColumn> GetColumns(Type rowType)
		{
			var result = new List<SheetColumn>();

			var bindingFlags = BindingFlags.Public | BindingFlags.Instance;
			MemberInfo[] members = rowType.GetFields(bindingFlags).Cast<MemberInfo>()
				.Concat(rowType.GetProperties(bindingFlags)).ToArray();

			foreach (var member in members)
			{
				var columnAttr = member.GetCustomAttributes(typeof(ColumnAttribute), true).FirstOrDefault() as ColumnAttribute;
				if (columnAttr != null)
				{
					var column = new SheetColumn();
					column.ColumnName = columnAttr.Column;
					column.Ignorable = columnAttr.Ignorable;
					column.IgnoreableValues = columnAttr.IgnorableValues;
					column.FieldName = member.Name;
					column.FieldType = member is FieldInfo ? ((FieldInfo)member).FieldType : ((PropertyInfo)member).PropertyType;
					result.Add(column);
				}
			}

			return result;
		}


		/// <summary>
		/// Retrieve a list of form fields from a given sheet
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheet"></param>
		/// <returns></returns>
		public static List<SheetField> GetFields<T>(ITypedSheetWithBulkData<T> sheet) where T : ISheetRow
		{
			var result = new List<SheetField>();

			var sheetType = sheet.GetType();

			var bindingFlags = BindingFlags.Public | BindingFlags.Instance;
			MemberInfo[] members = sheetType.GetFields(bindingFlags).Cast<MemberInfo>()
				.Concat(sheetType.GetProperties(bindingFlags)).ToArray();

			foreach (var member in members)
			{
				var fieldAttr = member.GetCustomAttributes(typeof(FieldAttribute), true).FirstOrDefault() as FieldAttribute;
				if (fieldAttr != null)
				{
					var field = new SheetField();
					field.CellAddress = fieldAttr.CellAddress;
					field.FieldName = member.Name;
					field.FieldType = member is FieldInfo ? ((FieldInfo)member).FieldType : ((PropertyInfo)member).PropertyType;
					result.Add(field);
				}
			}

			return result;
		}

		/// <summary>
		/// Assign a value to a given row instance
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="row"></param>
		/// <param name="column"></param>
		/// <param name="inValue"></param>
		public static void SetRowValue<T>(T row, string columnName, object inValue) where T : ISheetRow
		{
			if (inValue != null)
			{
				var obj = row as object;
				SetObjectFieldValue(row.GetType(), columnName, obj, inValue);
			}
		}

		/// <summary>
		/// Dynamically set a named property for a given type to a given value
		/// </summary>
		/// <param name="type"></param>
		/// <param name="fieldName"></param>
		/// <param name="instance"></param>
		/// <param name="inValue"></param>
		public static void SetObjectFieldValue(Type type, string fieldName, object instance, object inValue)
		{
			var accessor = TypeAccessor.Create(type);
			accessor[instance, fieldName] = inValue;
		}


		/// <summary>
		/// Map a specific cell value to the corresponding row and column
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="result"></param>
		/// <param name="columnLookup"></param>
		/// <param name="dataStartRowIndex"></param>
		/// <param name="rowIndex"></param>
		/// <param name="columnIndex"></param>
		/// <param name="displayRowIndex"></param>
		/// <param name="cellValue"></param>
		/// <param name="maximumNumberOfRows"></param>
		/// <param name="cellAddress"></param>
		public static void MapCell<T>(List<T> result, Dictionary<int, SheetColumn> columnLookup, int dataStartRowIndex, int rowIndex, int columnIndex, int displayRowIndex, object cellValue, int maximumNumberOfRows, string cellAddress) where T : class, ISheetRow
		{
			var listIdx = rowIndex - dataStartRowIndex;

			// Check that we are at a valid data row
			if (rowIndex >= dataStartRowIndex && columnLookup.ContainsKey(columnIndex))
			{
				// Disregard rows beyond the maximum number of rows
				if (maximumNumberOfRows > 0 && listIdx >= maximumNumberOfRows)
				{
					return;
				}
				else
				{
					var column = columnLookup[columnIndex];

					if (IsIgnorableValue(cellValue, column))
					{
						return;
                    }


					T row = null;
					if (result.ElementAtOrDefault(listIdx) == null)
					{
						row = Activator.CreateInstance<T>();
						row.RowIndex = displayRowIndex;

						result.Insert(listIdx, row);
					}
					else
					{
						row = result[listIdx];
					}

					var inValue = cellValue;
					object convertedValue;
					SyntaxError syntaxError;


					// Apply type hinting which will do the appropriate type conversion
					var outValue = Util.ApplyTypeHint(column.FieldType, inValue, out convertedValue, out syntaxError);

					if (syntaxError != null)
					{
						syntaxError.CellAddress = cellAddress;
						syntaxError.DataField = column;
						row.SyntaxErrorContainer.AddSyntaxError(syntaxError);
					}
					else
					{
						Util.SetRowValue(row, column.FieldName, convertedValue ?? outValue);
					}
				}
			}
		}

		public static void MapFields<T>(ITypedSheetWithBulkData<T> sheet, Func<SheetField, string> getValue) where T : ISheetRow
		{
			var fields = Util.GetFields(sheet);
			foreach (var field in fields)
			{
				object convertedValue;
				SyntaxError syntaxError;
				var value = getValue(field);

				if (IsIgnorableValue(value, field))
				{
					continue;
				}

				var outValue = Util.ApplyTypeHint(field.FieldType, value, out convertedValue, out syntaxError);
				if (syntaxError != null)
				{
					syntaxError.CellAddress = field.CellAddress;
					sheet.SyntaxErrorContainer.AddSyntaxError(syntaxError);
				}
				Util.SetObjectFieldValue(sheet.GetType(), field.FieldName, sheet, convertedValue ?? outValue);
			}
		}

		public static bool IsIgnorableValue(object cellValue, SheetDataField column)
		{
			return cellValue != null && column.IgnoreableValues != null && column.IgnoreableValues.Contains(cellValue.ToString());
		}

		public static object GetDefault(this Type type)
		{
			// If no Type was supplied, if the Type was a reference type, or if the Type was a System.Void, return null
			if (type == null || !type.IsValueType || type == typeof(void))
				return null;

			// If the supplied Type has generic parameters, its default value cannot be determined
			if (type.ContainsGenericParameters)
				throw new ArgumentException(
					"{" + MethodInfo.GetCurrentMethod() + "} Error:\n\nThe supplied value type <" + type +
					"> contains generic parameters, so the default value cannot be retrieved");

			// If the Type is a primitive type, or if it is another publicly-visible value type (i.e. struct/enum), return a 
			//  default instance of the value type
			if (type.IsPrimitive || !type.IsNotPublic)
			{
				try
				{
					return Activator.CreateInstance(type);
				}
				catch (Exception e)
				{
					throw new ArgumentException(
						"{" + MethodInfo.GetCurrentMethod() + "} Error:\n\nThe Activator.CreateInstance method could not " +
						"create a default instance of the supplied value type <" + type +
						"> (Inner Exception message: \"" + e.Message + "\")", e);
				}
			}

			// Fail with exception
			throw new ArgumentException("{" + MethodInfo.GetCurrentMethod() + "} Error:\n\nThe supplied value type <" + type +
				"> is not a publicly-visible type, so the default value cannot be retrieved");
		}
	}
}
