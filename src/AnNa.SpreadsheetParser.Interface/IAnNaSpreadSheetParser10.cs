using System;
using System.Collections.Generic;
using System.IO;
using AnNa.SpreadsheetParser.Interface.Sheets;

namespace AnNa.SpreadsheetParser.Interface
{
	public interface IAnNaSpreadSheetParser10
	{

		/// <summary>
		/// Open a XLSX file from a specific path.
		/// </summary>
		/// <param name="path">Path to XSLX file</param>
		/// <param name="password">Password if spreadsheet is encryptet/protected</param>
		void OpenFile(string path, string password = null);

		/// <summary>
		/// Retrieve the bulk data from a specific sheet
		/// </summary>
		/// <param name="sheetSpecification"></param>
		/// <returns></returns>
		List<Dictionary<string, string>> GetSheetContents(ISheetSpecification sheetSpecification);

		/// <summary>
		/// Write bulk data to a specific sheet
		/// </summary>
		/// <param name="sheetSpecification"></param>
		/// <param name="contents"></param>
		void SetSheetContents(ISheetSpecification sheetSpecification, List<Dictionary<string, string>> contents);

		/// <summary>
		/// Overload
		/// </summary>
		/// <param name="sheetSpecification"></param>
		/// <param name="cellAddress"></param>
		/// <returns></returns>
		string GetValueAt(ISheetSpecification sheetSpecification, string cellAddress);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc.
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <returns></returns>
		string GetValueAt(string sheet, string cellAddress);

		/// <summary>
		/// Overload
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheetSpecification"></param>
		/// <param name="cellAddress"></param>
		/// <param name="value"></param>
		void SetValueAt<T>(ISheetSpecification sheetSpecification, string cellAddress, T value);

		/// <summary>
		/// Set a value in a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <param name="value"></param>
		void SetValueAt<T>(string sheet, string cellAddress, T value);

		/// <summary>
		/// Save this spreadsheet to a specified path, or if no path is given to the file from which this spreadsheet was opened
		/// </summary>
		/// <param name="path"></param>
		void SaveToFile(string path = null);

		/// <summary>
		/// Save this spreadsheet to a stream.
		/// </summary>
		/// <returns></returns>
		Stream SaveToStream();

		/// <summary>
		/// Returns a list of sheet names found in the opened spreadsheet
		/// </summary>
		List<string> SheetNames { get; }
	}

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