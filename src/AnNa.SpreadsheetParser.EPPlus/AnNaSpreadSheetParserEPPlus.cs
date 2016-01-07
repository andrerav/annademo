using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AnNa.SpreadsheetParser.Interface;
using AnNa.SpreadsheetParser.Interface.Sheets;
using OfficeOpenXml;

namespace AnNa.SpreadSheetParser.EPPlus
{
	public class AnNaSpreadSheetParserEPPlus: IAnNaSpreadSheetParser10
	{
		private ExcelPackage _excelPackage;

		protected ExcelWorkbook Workbook => _excelPackage.Workbook;

		public void OpenFile(string path, string password = null)
		{
			_excelPackage = new ExcelPackage(new FileInfo(path), password);
		}

		public void OpenFile(Stream stream, string password = null)
		{
			_excelPackage = new ExcelPackage(stream, password);
		}

		public bool IsAnNaSpreadsheet()
		{
			if (Workbook != null)
			{
				var versionSheet = Workbook.Worksheets["Version"];
				if (versionSheet != null)
				{
					var versionCellValue = versionSheet.Cells["B3"].Value;
					if (versionCellValue.ToString().StartsWith("1.0"))
					{
						return true;
					}
				}
			}

			return false;
		}

		public void SaveToFile(string path = null, bool createDirectoryIfNotExists = false)
		{
			if (path == null)
			{
				_excelPackage.Save();
			}
			else
			{
				if (createDirectoryIfNotExists)
				{
					Util.CreateDirectoryIfNotExists(path);
				}
				_excelPackage.SaveAs(new FileInfo(path));
			}
		}

		public Stream SaveToStream()
		{
			var stream = new MemoryStream();
			_excelPackage.SaveAs(stream);
			return stream;
		}

		public List<string> SheetNames
		{
			get
			{
				ValidateWorkbook();
				return Workbook.Worksheets.Select(s => s.Name).ToList();
			}
		}

		public ITypedSheetWithBulkData<T> GetSheetBulkData<T>(ITypedSheetWithBulkData<T> sheet) where T :class, ISheetRow
		{
			ValidateWorkbook();
			var worksheet = GetWorksheet(sheet.SheetName);
			return worksheet != null ? RetrieveData(worksheet, sheet) : null;
		}

		public List<Dictionary<string, string>> GetSheetBulkData(ISheetWithBulkData sheet, int offset = 2)
		{
			ValidateWorkbook();
			var worksheet = GetWorksheet(sheet.SheetName);
			return worksheet != null ? RetrieveData(worksheet, sheet, offset) : new List<Dictionary<string, string>>();
		}

		public void SetSheetBulkData(ISheetWithBulkData sheet, List<Dictionary<string, string>> contents, int offset = 2)
		{
			ValidateWorkbook();
			var worksheet = GetWorksheet(sheet.SheetName);
			SetData(worksheet, sheet, contents, offset);
		}

		public void SetSheetData<T>(ITypedSheetWithBulkData<T> sheet) where T : class, ISheetRow
		{
			ValidateWorkbook();
			var worksheet = GetWorksheet(sheet.SheetName);
			SetData(worksheet, sheet);
		}

		public string GetValueAt(ISheet specification, string cellAddress)
		{
			return GetValueAt(specification.SheetName, cellAddress);
		}

		public string GetValueAt(string sheetName, string cellAddress)
		{
			ValidateWorkbook();
			var worksheet = GetWorksheet(sheetName);
			return worksheet?.Cells[cellAddress].Text;
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
			ValidateWorkbook();
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
			ValidateWorkbook();
			var worksheet = GetWorksheet(sheetName);
			if (worksheet != null)
			{
				worksheet.Cells[cellAddress].Value = value;
			}
		}

		private ITypedSheetWithBulkData<T> RetrieveData<T>(ExcelWorksheet worksheet, ITypedSheetWithBulkData<T> sheet) where T : class, ISheetRow
		{
			var result = new List<T>();
			int startrow = -1;
			List<SheetColumn> columnNames = Util.GetColumns(sheet);

			var columnLookup = CreateColumnLookup2(out startrow, worksheet, columnNames);

			var dataStartRowIndex = startrow + sheet.RowOffset;

			var cells = worksheet.Cells.Where(c => c.End.Column <= worksheet.Dimension.End.Column && c.End.Row <= worksheet.Dimension.End.Row).ToList();

			foreach (var cell in cells)
			{
				var rowIndex = cell.Start.Row;
				var columnIndex = cell.Start.Column;

				// NOTE: rows in SSG are zeroindexed, thus the +1
				var displayRowIndex = rowIndex + 1;
				var cellValue = cell.Value;
				var maximumNumberOfRows = sheet.MaximumNumberOfRows;
				Util.MapCell(result, columnLookup, dataStartRowIndex, rowIndex, columnIndex, displayRowIndex, cellValue, maximumNumberOfRows, cell.Address);
			}

			Util.RemoveEmptyRows(result, columnNames);

			sheet.Rows = result;
			// Map fields
			Util.MapFields(sheet, (field) => GetValueAt(sheet, field.CellAddress));
			return sheet;
		}

		private List<Dictionary<string, string>> RetrieveData(ExcelWorksheet worksheet, ISheetWithBulkData sheet, int offset)
		{
			var result = new List<Dictionary<string, string>>();
			var columnNames = sheet.ColumnNames;
			int startrow;
			var columnLookup = CreateColumnLookup(out startrow, worksheet, columnNames);

			var dataStartRow = startrow + offset;
			var cells = worksheet.Cells.Where(
				c => c.End.Column <= worksheet.Dimension.End.Column && c.End.Row <= worksheet.Dimension.End.Row).ToList();

			foreach (var cell in cells)
			{
				var listIdx = cell.Start.Row - dataStartRow;

				// Check that we are at a valid data row
				if (cell.Start.Row >= dataStartRow && columnLookup.ContainsKey(cell.Start.Column))
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
						foreach (var col in columnNames)
						{
							dict[col] = null;
						}

						// Set row index and column index
						dict["__RowIndex"] = cell.Start.Row.ToString();

						result.Insert(listIdx, dict);
					}
					var columnName = columnLookup[cell.Start.Column];
					var inValue = cell.Value;
					var typeHint = columnName.GetTypeHint(sheet);
					object convertedValue;
					var outValue = Util.ApplyTypeHint(typeHint, inValue, out convertedValue);
					result[listIdx][columnName] = outValue?.ToString();
				}
			}

			Util.RemoveEmptyRows(result);

			return result;
		}


		private void SetData(ExcelWorksheet worksheet, ISheetWithBulkData sheet, List<Dictionary<string, string>> contents, int offset)
		{
			// Find all the known columns and map them to spreadsheet columns
			var columnNames = sheet.ColumnNames;

			int startrow;
			var columnLookup = CreateColumnLookup(out startrow, worksheet, columnNames);

			var dataStartRow = startrow + offset;

			// Copy the formatting of the first row to every row that will contain data
			// This is a user convenience and not part of the standard.
			//if (contents.Any())
			//{
			//	var rowOffset = 0;
			//	foreach (var entry in contents)
			//	{

			//		var columns = columnLookup.Keys;
			//		var startColumn = columns.Min();
			//		var endColumn = columns.Max();

			//		worksheet.Cells[dataStartRow, startColumn, dataStartRow, endColumn]
			//			.Copy(worksheet.Cells[dataStartRow + rowOffset, startColumn]);

			//		rowOffset++;
			//	}
			//}

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

		private void SetData<T>(ExcelWorksheet worksheet, ITypedSheetWithBulkData<T> sheet) where T : class, ISheetRow
		{
			// Find all the known columns and map them to spreadsheet columns
			int startrow;
			var columns = Util.GetColumns(sheet);
			var columnLookup = CreateColumnLookup2(out startrow, worksheet, columns);
			var dataStartRow = startrow + sheet.RowOffset;

			var helper = new TypeAccessorHelper(typeof(T));

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
						cell.Value = helper.Get(row, column.FieldName);
					}
				}
				i++;
			}

			// Set field data
			var sheetAccessHelper = new TypeAccessorHelper(sheet.GetType());
			var fields = Util.GetFields<T>(sheet);
			foreach (var field in fields)
			{
				worksheet.Cells[field.CellAddress].Value = sheetAccessHelper.Get(sheet, field.FieldName);
			}
		}

		#region Utility Methods
		private void ValidateWorkbook()
		{
			if (Workbook == null)
			{
				throw new InvalidOperationException(
					"You must use OpenFile() to open a spreadsheet before you can retrieve any contents");
			}
		}


		private static Dictionary<int, string> CreateColumnLookup(out int startrow, ExcelWorksheet worksheet, List<string> columnNames)
		{
			startrow = -1;
			var columnLookup = new Dictionary<int, string>();
			foreach (var columnName in columnNames)
			{
				var cell = worksheet.Cells.FirstOrDefault(c => c.Value != null && c.Value.ToString() == columnName);

				// If the column was not found, then throw exception since this spreadsheet is probably not following the standard
				if (cell == null)
				{
					// Skip this column
					continue;
				}

				// Save the starting point for the data
				if (startrow == -1)
				{
					startrow = cell.Start.Row;
				}
				else
				{
					if (startrow != cell.Start.Row)
					{
						throw new InvalidColumnPositionException("All columns must be placed on the same row");
					}
				}

				columnLookup[cell.Start.Column] = columnName;
			}

			return columnLookup;
		}

		private Dictionary<int, SheetColumn> CreateColumnLookup2(out int startrow, ExcelWorksheet worksheet, List<SheetColumn> columns)
		{
			startrow = -1;
			var columnLookup = new Dictionary<int, SheetColumn>();
			foreach (var column in columns)
			{
				var cell = worksheet.Cells.FirstOrDefault(c => c.Value != null && c.Value.ToString() == column.ColumnName);

				// If the column was not found, then throw exception since this spreadsheet is probably not following the standard
				if (cell == null)
				{
					// Skip this column
					continue;
				}

				// Save the starting point for the data
				if (startrow == -1)
				{
					startrow = cell.Start.Row;
				}
				else
				{
					if (startrow != cell.Start.Row)
					{
						throw new InvalidColumnPositionException("All columns must be placed on the same row");
					}
				}

				columnLookup[cell.Start.Column] = column;
			}

			return columnLookup;
		}


		private ExcelWorksheet GetWorksheet(string sheetName)
		{
			return Workbook.Worksheets.FirstOrDefault(s => s.Name.ToLower() == sheetName.ToLower());
		}

		#endregion

		public void Dispose()
		{
			_excelPackage.Dispose();
			_excelPackage = null;
		}

		public bool TryGetWorkbookVersion(out Version version)
		{
			version = null;

			if (Workbook != null)
			{
				var versionSheet = Workbook.Worksheets["Version"];
				if (versionSheet != null)
				{
					var versionCellValue = versionSheet.Cells["B3"].Value.ToString();
					var separatorIndex = versionCellValue.IndexOf('-');
					var versionString = separatorIndex > -1 ? versionCellValue.Substring(0, separatorIndex) : versionCellValue;

					return Version.TryParse(versionString, out version);
				}
			}

			return false;
		}
	}
}
