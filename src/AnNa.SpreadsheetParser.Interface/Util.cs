using AnNa.SpreadsheetParser.Interface.Attributes;
using AnNa.SpreadsheetParser.Interface.Extensions;
using AnNa.SpreadsheetParser.Interface.Sheets;
using FastMember;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;

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

			//Special case for string: Avoids empty strings preventing clean up of empty rows
			if (typeHint == typeof(string) && string.IsNullOrWhiteSpace(inValue.ToString()))
				outValue = null;

			convertedValue = outValue;

			return outValue;
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
		/// <typeparam name="R"></typeparam>
		/// <typeparam name="F"></typeparam>
		/// <param name="sheet"></param>
		/// <returns></returns>
		public static List<SheetColumn> GetColumns<R, F>(ITypedSheet<R, F> sheet) 
			where R : ISheetRow 
			where F : ISheetFields
		{

			return GetColumns(typeof(R));

		}

		public static List<SheetColumn> GetColumns(SheetRow row)
		{
			return GetColumns(row.GetType());
		}

		public static List<SheetColumn> GetColumns(Type rowType)
		{
			var result = new List<SheetColumn>();

			foreach (var member in ReflectionHelpers.GetAllNonObsoleteFieldsAndProperties(rowType))
			{
				var columnAttr = member.GetCustomAttributes(typeof(ColumnAttribute), true).FirstOrDefault() as ColumnAttribute;
				if (columnAttr != null)
				{
					var column = new SheetColumn()
					{
						ColumnName = columnAttr.Column,
						Ignorable = columnAttr.Ignorable,
						IgnoreableValues = columnAttr.IgnorableValues,
						IsOptional = columnAttr.IsOptional,
						FieldName = member.Name,
						FieldType = member is FieldInfo ? ((FieldInfo)member).FieldType : ((PropertyInfo)member).PropertyType,
						FriendlyName = columnAttr.FriendlyName
					};
					
					result.Add(column);
				}
			}

			return result;
		}


		/// <summary>
		/// Retrieve a list of form fields from a given sheet
		/// </summary>
		/// <typeparam name="R"></typeparam>
		/// <typeparam name="F"></typeparam>
		/// <param name="sheet"></param>
		/// <returns></returns>
		public static List<SheetField> GetFields<R, F>(ITypedSheet<R, F> sheet) 
			where R : ISheetRow 
			where F : ISheetFields
		{
			var result = new List<SheetField>();

			foreach (var member in ReflectionHelpers.GetAllNonObsoleteFieldsAndProperties(typeof(F)))
			{
				var fieldAttr = member.GetCustomAttributes(typeof(FieldAttribute), true).FirstOrDefault() as FieldAttribute;
				if (fieldAttr != null)
				{
					var field = new SheetField()
					{
						CellAddress = fieldAttr.CellAddress,
						FieldName = member.Name,
						FieldType = member is FieldInfo ? ((FieldInfo)member).FieldType : ((PropertyInfo)member).PropertyType,
						FriendlyName = fieldAttr.FriendlyName,
						IsOptional = fieldAttr.IsOptional,
						IgnoreableValues = fieldAttr.IgnorableValues
					};
					
					result.Add(field);
				}
			}

			return result;
		}

		/// <summary>
		/// Retrieves a collection of field lists - one list for each property in a sheet definition with a ListMappingAttribute 
		/// </summary>
		/// <typeparam name="R"></typeparam>
		/// <typeparam name="F"></typeparam>
		/// <param name="sheet"></param>
		/// <returns></returns>
		public static List<List<SheetField>> GetListMaps<R, F>(ITypedSheet<R, F> sheet)
			where R : ISheetRow
			where F : ISheetFields
		{
			var result = new List<List<SheetField>>();

			foreach (var member in ReflectionHelpers.GetAllNonObsoleteFieldsAndProperties(typeof(F)))
			{
				if (member.MemberType == MemberTypes.Field)
				{
					var fieldType = ((FieldInfo)member).FieldType;
					if (!fieldType.IsGenericType || fieldType.GetGenericTypeDefinition() != typeof(List<>))
						continue;
				}

				if(member.MemberType == MemberTypes.Property)
				{
					var fieldType = ((PropertyInfo)member).PropertyType;
					if (!fieldType.IsGenericType || fieldType.GetGenericTypeDefinition() != typeof(List<>))
						continue;
				}

				var listMappingAttr = member.GetCustomAttributes(typeof(ListMappingAttribute), true).FirstOrDefault() as ListMappingAttribute;
				if(listMappingAttr != null)
				{
					var listMappingFields = new List<SheetField>();

					foreach (var cell in listMappingAttr.CellAddresses)
					{
						var field = new SheetField
						{
							CellAddress = cell,
							FieldName = member.Name,
							FieldType = member is FieldInfo ? ((FieldInfo)member).FieldType : ((PropertyInfo)member).PropertyType,
							FriendlyName = listMappingAttr.FriendlyName,
							IsOptional = listMappingAttr.IsOptional,
							IgnoreableValues = listMappingAttr.IgnorableValues
						};

						listMappingFields.Add(field);
					}

					if (listMappingFields.Any())
						result.Add(listMappingFields);
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
		public static void SetRowValue<T>(T row, Type columnType,  string columnName, object inValue) where T : ISheetRow
		{
			var obj = row as object;
			SetObjectFieldValue(row.GetType(), columnType, columnName, obj, inValue);
		}

		/// <summary>
		/// Dynamically set a named property for a given type to a given value
		/// </summary>
		/// <param name="instanceType"></param>
		/// <param name="fieldName"></param>
		/// <param name="instance"></param>
		/// <param name="inValue"></param>
		public static void SetObjectFieldValue(Type instanceType, Type fieldType, string fieldName, object instance, object inValue)
		{
			var accessor = TypeAccessor.Create(instanceType);

			//Can't assign null to value types
			if (inValue == null && (fieldType.IsValueType && Nullable.GetUnderlyingType(fieldType) == null))
				return;

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
						Util.SetRowValue(row, column.FieldType, column.FieldName, convertedValue ?? outValue);
					}
				}
			}
		}

		public static void MapFields<R, F>(ITypedSheet<R, F> sheet, Func<SheetField, string> getValue) where R : ISheetRow where F : ISheetFields
		{
			var fields = Util.GetFields(sheet);

			if (fields.Any() && sheet.Fields == null)
				sheet.Fields = (F)Activator.CreateInstance(typeof(F));

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
					syntaxError.DataField = field;
					sheet.SyntaxErrorContainer.AddSyntaxError(syntaxError);
				}

				Util.SetObjectFieldValue(typeof(F), field.FieldType, field.FieldName, sheet.Fields, convertedValue ?? outValue);
			}
		}

		public static void MapLists<R, F>(ITypedSheet<R,F> sheet, Func<SheetField, string> getValue) where R : ISheetRow where F : ISheetFields
		{
			var lists = Util.GetListMaps(sheet);

			if (lists.Any() && sheet.Fields == null)
				sheet.Fields = (F)Activator.CreateInstance(typeof(F));

			foreach (var list in lists)
			{
				var collectionType = list.First().FieldType; //e.g. List<string>

				var fieldName = list.First().FieldName;
				var fieldCollection = Activator.CreateInstance(collectionType);
				var addMethod = collectionType.GetMethod(nameof(list.Add));

				foreach (var field in list)
				{
					object convertedValue;
					SyntaxError syntaxError;
					var value = getValue(field);

					if (IsIgnorableValue(value, field))
					{
						continue;
					}

					var outValue = Util.ApplyTypeHint(field.FieldType.GenericTypeArguments[0], value, out convertedValue, out syntaxError);
					if (syntaxError != null)
					{
						syntaxError.CellAddress = field.CellAddress;
						syntaxError.DataField = field;
						sheet.SyntaxErrorContainer.AddSyntaxError(syntaxError);
					}

					addMethod.Invoke(fieldCollection, new object[] { convertedValue ?? outValue });
				}

				Util.SetObjectFieldValue(typeof(F), fieldCollection.GetType(), fieldName, sheet.Fields, fieldCollection);
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


		public static ITypedSheet<R1,F1> MapFrom<R1, R2, F1, F2>(this ITypedSheet<R1, F1> target, ITypedSheet<R2, F2> source) 
			where R1: ISheetRow
			where R2: ISheetRow, R1
			where F1: ISheetFields
			where F2: ISheetFields, F1
		{
			target.Fields = (F1)source.Fields;
			target.Rows = source.Rows?.Cast<R1>().ToList();

			target.SyntaxErrorContainer = source.SyntaxErrorContainer;

			return target;
		}

	}

	public class TypeAccessorHelper
	{
		TypeAccessor _accessor;

		public TypeAccessorHelper(Type t)
		{
			_accessor = TypeAccessor.Create(t);
		}

		public void Set(object instance, string fieldName, object value)
		{
			_accessor[instance, fieldName] = value;
		}

		public object Get(object instance, string fieldName)
		{
			return _accessor[instance, fieldName];
		}

	}
}
