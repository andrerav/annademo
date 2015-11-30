using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface
{
	public static class Util
	{

		/// <summary>
		/// Accepts a type and a string value and attempts to do a proper type conversion
		/// </summary>
		/// <param name="inValue"></param>
		/// <returns>A simple ToString() representation of the converted value, or the input value if conversion was not possible or necessary</returns>
		public static string ApplyTypeHint(Type typeHint, object inValue, out object convertedValue)
		{
			object outValue = null;
			if (inValue != null)
			{
				// Type hint: DateTime
				if (typeHint == typeof(DateTime))
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
							}
						}
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
				return inValue.ToString();
			}
			else
			{
				return outValue.ToString();
			}
		}

		public static string ApplyTypeHint<T>(object inValue, out object convertedValue)
		{
			return ApplyTypeHint(typeof(T), inValue, out convertedValue);
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
	}
}
