using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AnNa.SpreadsheetParser.Interface;
using AnNa.SpreadsheetParser.Interface.Sheets;
using SpreadsheetGear;

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

		public void SaveToFile(string path = null)
		{
			if (path == null)
			{
				_workbook.Save();
			}
			else
			{
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

		public List<Dictionary<string, string>> GetSheetContents(ISheetSpecification sheetSpecification)
		{
			ValidateWorkbook();
			var sheet = GetWorksheet(sheetSpecification.SheetName);
			return sheet != null ? RetrieveData(sheet, sheetSpecification) : new List<Dictionary<string, string>>();
		}

		public void SetSheetContents(ISheetSpecification sheetSpecification, List<Dictionary<string, string>> contents)
		{
			ValidateWorkbook();
			var sheet = GetWorksheet(sheetSpecification.SheetName);
			SetData(sheet, sheetSpecification, contents);
		}

		private void ValidateWorkbook()
		{
			if (Workbook == null)
			{
				throw new InvalidOperationException(
					"You must use OpenFile() to open a spreadsheet before you can retrieve any contents");
			}
		}

		public string GetValueAt(ISheetSpecification sheetSpecification, string cellAddress)
		{
			return GetValueAt(sheetSpecification.SheetName, cellAddress);
		}

		public string GetValueAt(string sheetName, string cellAddress)
		{
			ValidateWorkbook();
			var sheet = GetWorksheet(sheetName);
			return sheet?.Cells[cellAddress].Value.ToString();
		}

		public void SetValueAt<T>(ISheetSpecification sheetSpecification, string cellAddress, T value)
		{
			SetValueAt(sheetSpecification.SheetName, cellAddress, value);
		}

		public void SetValueAt<T>(string sheetName, string cellAddress, T value)
		{
			ValidateWorkbook();
			var sheet = GetWorksheet(sheetName);
			if (sheet != null)
			{
				sheet.Cells[cellAddress].Value = value;
			}
		}

		private IWorksheet GetWorksheet(string name)
		{
			return Workbook.Worksheets.Cast<IWorksheet>().FirstOrDefault(sheet => sheet.Name.ToLower() == name.ToLower());
		}

		/// <summary>
		/// Retrieve data in a given sheet
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		private List<Dictionary<string, string>> RetrieveData(IWorksheet sheet, ISheetSpecification sheetSpecification)
		{
			var result = new List<Dictionary<string, string>>();
			int startrow = -1;
			var columnLookup = CreateColumnLookup(out startrow, sheet, sheetSpecification.ColumnNames);

			var dataStartRow = startrow + 2;

			foreach (IRange cell in sheet.UsedRange)
			{
				var listIdx = cell.Row - dataStartRow;

				// Check that we are at a valid data row
				if (cell.Row >= dataStartRow && columnLookup.ContainsKey(cell.Column))
				{
					// Disregard rows beyond the maximum number of rows
					if (sheetSpecification.MaximumNumberOfRows > 0 && listIdx >= sheetSpecification.MaximumNumberOfRows)
					{
						continue;
					}

					if (result.ElementAtOrDefault(listIdx) == null)
					{
						var dict = new Dictionary<string, string>();

						// Initialize dictionary with null values
						foreach (var col in sheetSpecification.ColumnNames)
						{
							dict[col] = null;
						}

						// Set row index and column index
						dict["__RowIndex"] = cell.Row.ToString();

						result.Insert(listIdx, dict);
					}

					var columnName = columnLookup[cell.Column];
					var inValue = cell.Value;
					var typeHint = columnName.GetTypeHint(sheetSpecification);
					var outValue = Util.ApplyTypeHint(typeHint, inValue);

					result[listIdx][columnLookup[cell.Column]] = outValue;
				}
			}

			return result;
		}

		private void SetData(IWorksheet sheet, ISheetSpecification sheetSpecification, List<Dictionary<string, string>> contents)
		{
			// Find all the known columns and map them to spreadsheet columns
			var columnNames = sheetSpecification.ColumnNames;

			int startrow;
			var columnLookup = CreateColumnLookup(out startrow, sheet, columnNames);

			var dataStartRow = startrow + 2;

			foreach (IRange cell in sheet.UsedRange)
			{
				var listIdx = cell.Row - dataStartRow;

				// Check that we are at a valid data row
				if (cell.Row >= dataStartRow && columnLookup.ContainsKey(cell.Column))
				{
					// Disregard rows beyond the maximum number of rows
					if (sheetSpecification.MaximumNumberOfRows > 0 && listIdx >= sheetSpecification.MaximumNumberOfRows)
					{
						continue;
					}

					var columnName = columnLookup[cell.Column];
					cell.Value = contents[listIdx][columnName];
				}
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
			foreach (IWorksheet sheet in workbook.Sheets)
			{
				if (sheet.ProtectContents)
				{
					sheet.Unprotect(password);
					sheet.ProtectContents = false;
				}
			}

			workbook.Unprotect(password);
			workbook.ProtectStructure = false;
		}

		private Dictionary<int, string> CreateColumnLookup(out int startrow, IWorksheet sheet, List<string> columnNames)
		{
			startrow = -1;
			var columnLookup = new Dictionary<int, string>();
			foreach (var columnName in columnNames)
			{
				var cell = sheet.UsedRange.Find(columnName, null, FindLookIn.Values, LookAt.Whole, SearchOrder.ByColumns, SearchDirection.Next, matchCase: true);

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
	}
	#endregion // Utility methods
}
