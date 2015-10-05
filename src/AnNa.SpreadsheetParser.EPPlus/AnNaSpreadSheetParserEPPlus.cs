﻿using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using AnNaSpreadSheetParser;
using OfficeOpenXml;

namespace AnNa.SpreadSheetParser.EPPlus
{
    public class AnNaSpreadSheetParserEPPlus: IAnNaSpreadSheetParser10
	{

		protected const string Version = "1.0";
		private ExcelWorkbook _workbook;

		public ExcelWorkbook Workbook
		{
			get { return _workbook; }
			protected set { _workbook = value; }
		}

		public void OpenFile(string path, string password = null)
	    {
			_workbook = new ExcelPackage(new FileInfo(path), password).Workbook;
		}

	    public List<Dictionary<string, string>> GetSheetContents(ISheetSpecification sheetSpecification)
	    {
			if (Workbook == null)
			{
				throw new InvalidOperationException("You must use OpenFile() to open a spreadsheet before you can retrieve any contents");
			}
			foreach (ExcelWorksheet sheet in Workbook.Worksheets)
			{
				if (sheet.Name.ToLower() == sheetSpecification.Sheet.ToString().ToLower())
				{
					return RetrieveData(sheet, sheetSpecification);
				}
			}

			return new List<Dictionary<string, string>>();
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


					var value = cell.Value;
					result[listIdx][columnLookup[cell.Start.Column]] = value != null ? value.ToString() : null;
				}

				// Stop if we have reached the maximum number of rows
				if (sheetSpecification.MaximumNumberOfRows > 0 && result.Count == sheetSpecification.MaximumNumberOfRows)
				{
					break;
				}
			}

			return result;

		}
	}
}