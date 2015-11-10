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
		private ExcelWorkbook _workbook;

		protected ExcelWorkbook Workbook
		{
			get { return _workbook; }
			set { _workbook = value; }
		}

		public void OpenFile(string path, string password = null)
	    {
			_workbook = new ExcelPackage(new FileInfo(path), password).Workbook;
		}

	    public List<Dictionary<string, string>> GetSheetContents(ISheetSpecification sheetSpecification)
	    {
			ValidateWorkbook();

			foreach (ExcelWorksheet sheet in Workbook.Worksheets)
			{
				if (sheet.Name.ToLower() == sheetSpecification.SheetName.ToLower())
				{
					return RetrieveData(sheet, sheetSpecification);
				}
			}

			return new List<Dictionary<string, string>>();
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
			foreach (var sheet in Workbook.Worksheets)
			{
				if (sheet.Name.ToLower() == sheetName.ToLower())
				{
					return sheet.Cells[cellAddress].Text.ToString();
				}
			}

			return null;
		}

		private List<Dictionary<string, string>> RetrieveData(ExcelWorksheet sheet, ISheetSpecification sheetSpecification)
	    {
			var result = new List<Dictionary<string, string>>();
			var columnLookup = new Dictionary<int, string>();
			int startrow = -1;

			// Find all the known AnNa columns and map them to spreadsheet columns
			foreach (var columnName in sheetSpecification.ColumnNames)
			{
				//var cell = sheet.Find(columnName, null, FindLookIn.Values, LookAt.Whole, SearchOrder.ByColumns, SearchDirection.Next, matchCase: true);

				var cell = sheet.Cells.Where(c => c.Value != null && c.Value.ToString() == columnName).FirstOrDefault();

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
						foreach (var col in sheetSpecification.ColumnNames)
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
	}
}
