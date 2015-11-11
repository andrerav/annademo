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

		protected const string Version = "1.0";
	    private ExcelPackage _excelPackage;

	    protected ExcelWorkbook Workbook => _excelPackage.Workbook;

	    public void OpenFile(string path, string password = null)
	    {
			_excelPackage = new ExcelPackage(new FileInfo(path), password);
		}

		public void SaveToFile(string path = null)
		{
			if (path == null)
			{
				_excelPackage.Save();
			}
			else
			{
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

	    public string GetValueAt(ISheetSpecification specification, string cellAddress)
	    {
		    return GetValueAt(specification.SheetName, cellAddress);
	    }

		public string GetValueAt(string sheetName, string cellAddress)
		{
			ValidateWorkbook();
			var sheet = GetWorksheet(sheetName);
			return sheet?.Cells[cellAddress].Text;
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

	    private ExcelWorksheet GetWorksheet(string sheetName)
	    {
		    return Workbook.Worksheets.FirstOrDefault(s => s.Name.ToLower() == sheetName.ToLower());
	    }

	    private List<Dictionary<string, string>> RetrieveData(ExcelWorksheet sheet, ISheetSpecification sheetSpecification)
	    {
			var result = new List<Dictionary<string, string>>();
		    var columnNames = sheetSpecification.ColumnNames;
			int startrow;
			var columnLookup = CreateColumnLookup(out startrow, sheet, columnNames);

			var dataStartRow = startrow + 2;

			foreach (var cell in sheet.Cells)
			{
				var listIdx = cell.Start.Row - dataStartRow;

				// Check that we are at a valid data row
				if (cell.Start.Row >= dataStartRow && columnLookup.ContainsKey(cell.Start.Column))
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
					var typeHint = columnName.GetTypeHint(sheetSpecification);
					var outValue = Util.ApplyTypeHint(typeHint, inValue);

					result[listIdx][columnName] = outValue;
				}
			}
			return result;
		}


		private void SetData(ExcelWorksheet sheet, ISheetSpecification sheetSpecification, List<Dictionary<string, string>> contents)
		{
			// Find all the known columns and map them to spreadsheet columns
			var columnNames = sheetSpecification.ColumnNames;

			int startrow;
			var columnLookup = CreateColumnLookup(out startrow, sheet, columnNames);

			var dataStartRow = startrow + 2;

			foreach (var cell in sheet.Cells)
			{
				var listIdx = cell.Start.Row - dataStartRow;

				// Check that we are at a valid data row
				if (cell.Start.Row >= dataStartRow && columnLookup.ContainsKey(cell.Start.Column))
				{
					// Disregard rows beyond the maximum number of rows
					if (sheetSpecification.MaximumNumberOfRows > 0 && listIdx >= sheetSpecification.MaximumNumberOfRows)
					{
						continue;
					}

					var columnName = columnLookup[cell.Start.Column];
					cell.Value = contents[listIdx][columnName];
				}
			}
		}

	    private static Dictionary<int, string> CreateColumnLookup(out int startrow, ExcelWorksheet sheet, List<string> columnNames)
	    {
		    startrow = -1;
			var columnLookup = new Dictionary<int, string>();
			foreach (var columnName in columnNames)
		    {
			    //var cell = sheet.Find(columnName, null, FindLookIn.Values, LookAt.Whole, SearchOrder.ByColumns, SearchDirection.Next, matchCase: true);

			    var cell = sheet.Cells.FirstOrDefault(c => c.Value != null && c.Value.ToString() == columnName);

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
	}
}
