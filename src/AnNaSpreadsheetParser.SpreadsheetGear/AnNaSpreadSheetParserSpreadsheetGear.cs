using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AnNa.SpreadsheetParser.Interface;
using AnNa.SpreadsheetParser.Interface.Sheets;
using SpreadsheetGear;
using ISheet = AnNa.SpreadsheetParser.Interface.Sheets.ISheet;
using System.Globalization;
using AnNa.SpreadsheetParser.Interface.Sheets.Typed;

namespace AnNa.SpreadsheetParser.SpreadsheetGear
{
	public class AnNaSpreadSheetParserSpreadsheetGear : IAnNaSpreadSheetParser10 
	{

		private IWorkbook _workbook;
		private CultureInfo _cultureInfo = new CultureInfo(CultureInfo.CurrentCulture.LCID);

		public IWorkbook Workbook
		{
			get { return _workbook; }
			protected set { _workbook = value; }
		}

		public void OpenFile(string path, string password = null)
		{
			_workbook = Factory.GetWorkbook(path, _cultureInfo);

			if (password != null)
			{
				UnprotectSpreadsheet(password, _workbook);
			}
		}

		public void OpenFile(Stream stream, string password = null)
		{
			_workbook = Factory.GetWorkbookSet(_cultureInfo).Workbooks.OpenFromStream(stream);
			if (password != null)
			{
				UnprotectSpreadsheet(password, _workbook);
			}

		}

		public void SaveToFile(string path = null, bool createDirectoryIfNotExists = false)
		{
			if (path == null)
			{
				_workbook.Save();
			}
			else
			{
				if (createDirectoryIfNotExists)
				{
					Util.CreateDirectoryIfNotExists(path);
				}
				_workbook.SaveAs(path,FileFormat.OpenXMLWorkbookMacroEnabled);
			}
		}

		public Stream SaveToStream()
		{
			return _workbook.SaveToStream(FileFormat.OpenXMLWorkbookMacroEnabled);
		}

		public List<string> SheetNames
		{
			get
			{
				ThrowExceptionIfNotInitialized();
				return Workbook.Worksheets.Cast<IWorksheet>().Select(s => s.Name).ToList();
			}
		}

		public bool IsAnNaSpreadsheet()
		{
			if (Workbook != null)
			{
				IWorksheet versionSheet = Workbook.Worksheets["Version"];
				if (versionSheet != null)
				{
					var versionCellValue = versionSheet.UsedRange["B3"].Value;
					if (versionCellValue.ToString().StartsWith("1.0"))
					{
						return true;
					}
				}
			}

			return false;
		}

		public List<Dictionary<string, string>> GetSheetBulkData(ISheetWithBulkData sheet, int offset = 2)
		{
			ThrowExceptionIfNotInitialized();
			var worksheet = GetWorksheet(sheet.SheetName);
			return worksheet != null ? RetrieveData(worksheet, sheet, offset) : new List<Dictionary<string, string>>();
		}

		public ITypedSheet<R, F> GetSheetBulkData<R, F>(ITypedSheet<R, F> sheet) where R : class, ISheetRow where F :class, ISheetFields
		{
			ThrowExceptionIfNotInitialized();
			var worksheet = GetWorksheet(sheet.SheetName);
			return worksheet != null ? RetrieveData(worksheet, sheet) : null;
		}

		public void SetSheetBulkData(ISheetWithBulkData sheet, List<Dictionary<string, string>> contents, int offset = 2)
		{
			ThrowExceptionIfNotInitialized();
			var worksheet = GetWorksheet(sheet.SheetName);
			SetData(worksheet, sheet, contents, offset);
		}

		public void SetSheetData<R, F>(ITypedSheet<R, F> sheet) where R : class, ISheetRow where F : class, ISheetFields
		{
			ThrowExceptionIfNotInitialized();
			var worksheet = GetWorksheet(sheet.SheetName);
			SetData(worksheet, sheet);
		}

		public string GetValueAt(ISheet sheet, string cellAddress)
		{
			return GetValueAt(sheet.SheetName, cellAddress);
		}

		public string GetValueAt(string sheetName, string cellAddress)
		{
			ThrowExceptionIfNotInitialized();
			var worksheet = GetWorksheet(sheetName);
			return worksheet?.Cells[cellAddress]?.Value?.ToString();
		}

		public T GetValueAt<T>(ISheet sheet, string cellAddress)
		{
			return GetValueAt<T>(sheet.SheetName, cellAddress);
		}

		public T GetValueAt<T>(string sheetName, string cellAddress)
		{
			string dummyRawString;
			return GetValueAt<T>(sheetName, cellAddress, out dummyRawString);
		}


		public T GetValueAt<T>(ISheet sheet, string cellAddress, out string rawString)
		{
			return GetValueAt<T>(sheet.SheetName, cellAddress, out rawString);
        }

		public T GetValueAt<T>(string sheetName, string cellAddress, out string rawString)
		{
			ThrowExceptionIfNotInitialized();
			var worksheet = GetWorksheet(sheetName);
			var rawValue = worksheet?.Cells[cellAddress]?.Value;
			rawString = rawValue != null ? rawValue.ToString() : string.Empty;
			object convertedValue;
			Util.ApplyTypeHint<T>(rawValue, out convertedValue);

			if (convertedValue is T)
			{
				return (T)convertedValue;
			}
			else
			{
				return default(T);
			}
		}

		public void SetValueAt<T>(ISheet sheet, string cellAddress, T value)
		{
			SetValueAt(sheet.SheetName, cellAddress, value);
		}

		public void SetValueAt<T>(string sheetName, string cellAddress, T value)
		{
			ThrowExceptionIfNotInitialized();
			var worksheet = GetWorksheet(sheetName);
			if (worksheet != null)
			{
				worksheet.Cells[cellAddress].Value = value;
			}
		}



		private ITypedSheet<R, F> RetrieveData<R, F>(IWorksheet worksheet, ITypedSheet<R, F> sheet) where R : class, ISheetRow where F : ISheetFields
		{
			var result = new List<R>();
			int startrow = -1;
			List<SheetColumn> columnNames = Util.GetColumns(sheet); //Columns from sheet definition

			var columnLookup = CreateColumnLookup2(out startrow, worksheet, columnNames); //Columns from sheet definition found in worksheet

			foreach (var missingColumn in columnNames.Select(c=> c?.ColumnName).Except(columnLookup.Select(c=> c.Value?.ColumnName)))
			{
				sheet.AddToMissingColumns(missingColumn);
			}

			var dataStartRowIndex = startrow + sheet.RowOffset;

			foreach (IRange cell in worksheet.UsedRange)
			{
				var rowIndex = cell.Row;
				var columnIndex = cell.Column;

				// NOTE: rows in SSG are zeroindexed, thus the +1
				var displayRowIndex = rowIndex + 1;
				var cellValue = cell.Value;
				var maximumNumberOfRows = sheet.MaximumNumberOfRows;
				Util.MapCell(result, columnLookup, dataStartRowIndex, rowIndex, columnIndex, displayRowIndex, cellValue, maximumNumberOfRows, cell.Address.Replace("$", string.Empty));
			}

			Util.RemoveEmptyRows(result, columnNames);

			sheet.Rows = result;
			
			Util.MapFields(sheet, (field) => GetValueAt(sheet, field.CellAddress));
			Util.MapLists(sheet, (field) => GetValueAt(sheet, field.CellAddress));

			return sheet;
		}


		/// <summary>
		/// Retrieve data in a given worksheet
		/// </summary>
		/// <param name="worksheet"></param>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		private List<Dictionary<string, string>> RetrieveData(IWorksheet worksheet, ISheetWithBulkData sheet, int offset)
		{
			var result = new List<Dictionary<string, string>>();
			int startrow = -1;
			var columnLookup = CreateColumnLookup(out startrow, worksheet, sheet.ColumnNames);

			var dataStartRow = startrow + offset;

			foreach (IRange cell in worksheet.UsedRange)
			{
				var listIdx = cell.Row - dataStartRow;

				// Check that we are at a valid data row
				if (cell.Row >= dataStartRow && columnLookup.ContainsKey(cell.Column))
				{
					// Disregard rows beyond the maximum number of rows
					if (sheet.MaximumNumberOfRows > 0 && listIdx >= sheet.MaximumNumberOfRows)
					{
						continue;
					}

					if (result.ElementAtOrDefault(listIdx) == null)
					{
						var dict = new Dictionary<string, string>();

						// Initialize dictionary with null values
						foreach (var col in sheet.ColumnNames)
						{
							dict[col] = null;
						}

						// Set row index and column index (NOTE: rows in SSG are zeroindexed, thus the +1)
						dict["__RowIndex"] = (cell.Row+1).ToString();

						result.Insert(listIdx, dict);
					}

					var columnName = columnLookup[cell.Column];
					var inValue = cell.Value;
					var typeHint = columnName.GetTypeHint(sheet);
					object convertedValue;
					var outValue = Util.ApplyTypeHint(typeHint, inValue, out convertedValue);
					result[listIdx][columnLookup[cell.Column]] = outValue?.ToString();
				}
			}
			Util.RemoveEmptyRows(result);
			return result;
		}

		private void SetData(IWorksheet worksheet, ISheetWithBulkData sheet, List<Dictionary<string, string>> contents, int offset)
		{
			// Find all the known columns and map them to spreadsheet columns
			int startrow;
			var columnNames = sheet.ColumnNames;
			var columnLookup = CreateColumnLookup(out startrow, worksheet, columnNames);
			var dataStartRow = startrow + offset;

			// Copy the formatting of the first row to every row that will contain data
			// This is a user convenience and not part of the standard.
			if (contents.Any())
			{

				var rowOffset = 0;
				foreach (var entry in contents)
				{

					var columns = columnLookup.Keys;
					var startColumn = columns.Min();
					var endColumn = columns.Max();

					worksheet.Cells[dataStartRow, startColumn, dataStartRow, endColumn]
						.Copy(worksheet.Cells[dataStartRow+rowOffset, startColumn], PasteType.All, PasteOperation.None, false,
							false);
					rowOffset++;
				}
			}


			int i = 0;
			foreach (var entry in contents)
			{
				foreach (var col in entry.Keys.Where(k => sheet.ColumnNames.Contains(k)))
				{
					var key = columnLookup.FirstOrDefault(x => x.Value == col).Key;
					var cell = worksheet.Cells[dataStartRow + i, key];
					if (cell != null)
					{
						cell.Value = entry[col];
					}
				}
				i++;
			}
		}

		private void SetData<R, F>(IWorksheet worksheet, ITypedSheet<R, F> sheet) where R :class, ISheetRow where F :class, ISheetFields
		{
			// Find all the known columns and map them to spreadsheet columns
			int startrow;
			var columns = Util.GetColumns(sheet);
			var columnLookup = CreateColumnLookup2(out startrow, worksheet, columns);
			var dataStartRow = startrow + sheet.RowOffset;

			var rowAccessHelper = new TypeAccessorHelper(typeof(R));

			// Set bulk data
			int i = 0;
			foreach (var row in sheet.Rows)
			{
				foreach (var column in columns)
				{
					var key = columnLookup.FirstOrDefault(x => x.Value == column).Key;
					var cell = worksheet.Cells[dataStartRow + i, key];

					if (cell != null)
					{
						cell.Value = rowAccessHelper.Get(row, column.FieldName);
					}
				}
				i++;
			}

			var sheetAccessHelper = new TypeAccessorHelper(typeof(F));

			// Set field data
			var fields = Util.GetFields(sheet)
				.Concat(Util.GetListMaps(sheet).SelectMany(lm => lm));

			foreach(var field in fields)
			{
				worksheet.Cells[field.CellAddress].Value = sheetAccessHelper.Get(sheet.Fields, field.FieldName);
            }
		}

		#region Utility methods
		/// <summary>
		/// Unprotected workbook and all sheets using the given password
		/// </summary>
		/// <param name="password"></param>
		/// <param name="workbook"></param>
		public static void UnprotectSpreadsheet(string password, IWorkbook workbook)
		{
			foreach (IWorksheet worksheet in workbook.Sheets)
			{
				if (worksheet.ProtectContents)
				{
					worksheet.Unprotect(password);
					worksheet.ProtectContents = false;
				}
			}

			workbook.Unprotect(password);
			workbook.ProtectStructure = false;
		}

		private Dictionary<int, SheetColumn> CreateColumnLookup2(out int startrow, IWorksheet worksheet, List<SheetColumn> columns)
		{
			startrow = -1;
			var columnLookup = new Dictionary<int, SheetColumn>();
			foreach (var column in columns)
			{
				var cell = worksheet.UsedRange.Find(column.ColumnName, null, FindLookIn.Values, LookAt.Whole, SearchOrder.ByColumns, SearchDirection.Next, matchCase: true);

				// If the column was not found, then throw exception since this spreadsheet is probably not following the standard
				if (cell == null)
				{
					// Skip this column
					continue;
				}

				// Save the starting point for the data
				if (startrow == -1)
				{
					startrow = cell.Row;
				}
				else
				{
					if (startrow != cell.Row)
					{
						throw new InvalidColumnPositionException("All columns must be placed on the same row");
					}
				}

				columnLookup[cell.Column] = column;
			}

			return columnLookup;
		}

		private Dictionary<int, string> CreateColumnLookup(out int startrow, IWorksheet worksheet, List<string> columnNames)
		{
			startrow = -1;
			var columnLookup = new Dictionary<int, string>();
			foreach (var columnName in columnNames)
			{
				var cell = worksheet.UsedRange.Find(columnName, null, FindLookIn.Values, LookAt.Whole, SearchOrder.ByColumns, SearchDirection.Next, matchCase: true);

				// If the column was not found, then throw exception since this spreadsheet is probably not following the standard
				if (cell == null)
				{
					// Skip this column
					continue;
				}

				// Save the starting point for the data
				if (startrow == -1)
				{
					startrow = cell.Row;
				}
				else
				{
					if (startrow != cell.Row)
					{
						throw new InvalidColumnPositionException("All columns must be placed on the same row");
					}
				}

				columnLookup[cell.Column] = columnName;
			}

			return columnLookup;
		}

		private IWorksheet GetWorksheet(string name)
		{
			return Workbook.Worksheets.Cast<IWorksheet>().FirstOrDefault(ws => ws.Name.ToLower() == name.ToLower());
		}

		public void ThrowExceptionIfNotInitialized()
		{
			if (Workbook == null)
			{
				throw new InvalidOperationException(
					"You must use OpenFile() to open a spreadsheet before you can retrieve any contents");
			}
		}
		#endregion // Utility methods

		public void Dispose()
		{
			_workbook?.Close();
			_workbook = null;
		}

		public bool TryGetWorkbookVersion(out Version version, out string authority)
		{
			version = null;
			authority = SheetAuthority.AnNa; //Default to AnNa

			if (Workbook != null)
			{
				var versionSheet = Workbook.Worksheets["Version"];
				if (versionSheet != null)
				{
					var versionCellValue = versionSheet.Cells["B3"].Value.ToString();
					var separatorIndex = versionCellValue.IndexOf('-');

					string versionString = versionCellValue;
					if (separatorIndex > -1)
					{
						versionString = versionCellValue.Substring(0, separatorIndex);
						authority = versionCellValue.Substring(separatorIndex + 1);
					}


					return Version.TryParse(versionString, out version);
				}
			}

			return false;
		}
	}
}
