﻿using System;
using System.Collections.Generic;
using System.Linq;
using SpreadsheetGear;

namespace AnNaSpreadSheetParser
{
	public enum AnNaSheets
	{
		// ReSharper disable once InconsistentNaming
		Crew_List,
		Pax_List,
		Waste_And_Residues,
		Security
	}

	public class AnNaSpreadSheetParser10
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

		/// <summary>
		/// Enumerate workbook sheets and return contents
		/// </summary>
		/// <param name="sheetName"></param>
		/// <returns></returns>
		public List<Dictionary<string, string>> GetSheetContents(ISheetSpecification sheetSpecification)
		{
			if (Workbook == null)
			{
				throw new InvalidOperationException("You must use OpenFile() to open a spreadsheet before you can retrieve any contents");
			}
			foreach (IWorksheet sheet in Workbook.Sheets)
			{
				if (sheet.Name.ToLower() == sheetSpecification.Sheet.ToString().ToLower())
				{
					return RetrieveData(sheet, sheetSpecification);
				}
			}

			return new List<Dictionary<string, string>>();
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
			var columnLookup = new Dictionary<int, string>();
			int startrow = -1;

			// Find all the known AnNa columns and map them to spreadsheet columns
			foreach (var columnName in sheetSpecification.ColumnNames)
			{
				var cell = sheet.UsedRange.Find(columnName, null,FindLookIn.Values, LookAt.Whole, SearchOrder.ByColumns, SearchDirection.Next, matchCase:true);

				// If the column was not found, then throw exception since this spreadsheet is probably not following the standard
				if (cell == null)
				{
					throw new ColumnNotFoundException(string.Format("Unable to find column {0} in sheet {1}", columnName,
						sheetSpecification.Sheet.ToString()));
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

			var dataStartRow = startrow + 2;

			foreach (IRange cell in sheet.UsedRange)
			{
				var listIdx = cell.Row - dataStartRow;

				// Check that we are at a valid data row
				if (cell.Row >= dataStartRow && columnLookup.ContainsKey(cell.Column))
				{

					if (result.ElementAtOrDefault(listIdx) == null)
					{
						result.Insert(listIdx, new Dictionary<string, string>());
					}


					var value = cell.Value;
					result[listIdx][columnLookup[cell.Column]] = value != null ? value.ToString() : null;
				}

				// Stop if we have reached the maximum number of rows
				if (sheetSpecification.MaximumNumberOfRows > 0 && result.Count == sheetSpecification.MaximumNumberOfRows)
				{
					break;
				}
			}

			return result;
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
		#endregion // Utility methods
	}

	//public interface IRange
	//{
	//	object Value { get; set; }
	//	int Column { get; set; }
	//	int Row { get; set; }
	//}

	//public interface IWorksheet
	//{
	//	IEnumerable<IRange> UsedRange { get; set; }
	//	bool ProtectContents { get; set; }
	//	string Name { get; set; }
	//	void Unprotect(string password);
	//}

	//public interface IWorkbook
	//{
	//	IEnumerable<IWorksheet> Sheets { get; }
	//	bool ProtectStructure { get; set; }
	//	void Unprotect(string password);
	//}

	public class InvalidColumnPositionException : Exception
	{
		public InvalidColumnPositionException(string message)
			: base(message)
		{
		}
	}

	public class ColumnNotFoundException : Exception
	{
		public ColumnNotFoundException(string message)
			: base(message)
		{
		}
	}
}
