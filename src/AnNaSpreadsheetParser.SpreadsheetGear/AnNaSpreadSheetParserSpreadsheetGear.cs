using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AnNa.SpreadsheetParser.Interface;
using AnNa.SpreadsheetParser.Interface.Sheets;
using SpreadsheetGear;
using ISheet = AnNa.SpreadsheetParser.Interface.Sheets.ISheet;

namespace AnNa.SpreadsheetParser.SpreadsheetGear
{
	public class AnNaSpreadSheetParserSpreadsheetGear : IAnNaSpreadSheetParser10 
	{

		protected const string Version = "1.0";
		private IWorkbook _workbook;

		public IWorkbook Workbook
		{
			get { return _workbook; }
			protected set { _workbook = value; }
		}

		public void OpenFile(string path, string password = null)
		{
			_workbook = Factory.GetWorkbook(path);

			if (password != null)
			{
				UnprotectSpreadsheet(password, _workbook);
			}
		}

		public void OpenFile(Stream stream, string password = null)
		{
			_workbook = Factory.GetWorkbookSet().Workbooks.OpenFromStream(stream);

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
				_workbook.SaveAs(path,FileFormat.OpenXMLWorkbook);
			}
		}

		public Stream SaveToStream()
		{
			return _workbook.SaveToStream(FileFormat.OpenXMLWorkbook);
		}

		public List<string> SheetNames
		{
			get
			{
				ValidateWorkbook();
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

		public List<Dictionary<string, string>> GetSheetBulkData(ISheetWithBulkData sheet)
		{
			ValidateWorkbook();
			var worksheet = GetWorksheet(sheet.SheetName);
			return worksheet != null ? RetrieveData(worksheet, sheet) : new List<Dictionary<string, string>>();
		}

		public void SetSheetBulkData(ISheetWithBulkData sheet, List<Dictionary<string, string>> contents)
		{
			ValidateWorkbook();
			var worksheet = GetWorksheet(sheet.SheetName);
			SetData(worksheet, sheet, contents);
		}

		public string GetValueAt(ISheet sheet, string cellAddress)
		{
			return GetValueAt(sheet.SheetName, cellAddress);
		}

		public string GetValueAt(string sheetName, string cellAddress)
		{
			ValidateWorkbook();
			var worksheet = GetWorksheet(sheetName);
			return worksheet?.Cells[cellAddress].Value.ToString();
		}

		public void SetValueAt<T>(ISheet sheet, string cellAddress, T value)
		{
			SetValueAt(sheet.SheetName, cellAddress, value);
		}

		public void SetValueAt<T>(string sheetName, string cellAddress, T value)
		{
			ValidateWorkbook();
			var worksheet = GetWorksheet(sheetName);
			if (worksheet != null)
			{
				worksheet.Cells[cellAddress].Value = value;
			}
		}

		/// <summary>
		/// Retrieve data in a given worksheet
		/// </summary>
		/// <param name="worksheet"></param>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		private List<Dictionary<string, string>> RetrieveData(IWorksheet worksheet, ISheetWithBulkData sheet)
		{
			var result = new List<Dictionary<string, string>>();
			int startrow = -1;
			var columnLookup = CreateColumnLookup(out startrow, worksheet, sheet.ColumnNames);

			var dataStartRow = startrow + 2;

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

						// Set row index and column index
						dict["__RowIndex"] = cell.Row.ToString();

						result.Insert(listIdx, dict);
					}

					var columnName = columnLookup[cell.Column];
					var inValue = cell.Value;
					var typeHint = columnName.GetTypeHint(sheet);
					var outValue = Util.ApplyTypeHint(typeHint, inValue);

					result[listIdx][columnLookup[cell.Column]] = outValue;
				}
			}

			return result;
		}

		private void SetData(IWorksheet worksheet, ISheetWithBulkData sheet, List<Dictionary<string, string>> contents)
		{
			// Find all the known columns and map them to spreadsheet columns
			int startrow;
			var columnNames = sheet.ColumnNames;
			var columnLookup = CreateColumnLookup(out startrow, worksheet, columnNames);
			var dataStartRow = startrow + 2;

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

		private void ValidateWorkbook()
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
	}
}
