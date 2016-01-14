using System;
using System.Collections.Generic;
using System.IO;
using AnNa.SpreadsheetParser.Interface.Sheets;
using System.Linq;
using AnNa.SpreadsheetParser.Interface.Attributes;
using AnNa.SpreadsheetParser.Interface.Sheets.Typed;

namespace AnNa.SpreadsheetParser.Interface
{
	public interface IAnNaSpreadSheetParser10 : IDisposable
	{

		/// <summary>
		/// Open a XLSX file from a specific path.
		/// </summary>
		/// <param name="path">Path to XSLX file</param>
		/// <param name="password">Password if spreadsheet is encryptet/protected</param>
		void OpenFile(string path, string password = null);

		/// <summary>
		/// Open a XLSX file from a specific path.
		/// </summary>
		/// <param name="stream">Path to XSLX file</param>
		/// <param name="password">Password if spreadsheet is encryptet/protected</param>
		void OpenFile(Stream stream, string password = null);

		/// <summary>
		/// Returns true if this sheet has been recognized as an AnNa spreadsheet
		/// </summary>
		/// <returns></returns>
		bool IsAnNaSpreadsheet();

		/// <summary>
		/// Retrieve the bulk data from a specific sheet
		/// </summary>
		/// <param name="sheet"></param>
		/// <returns></returns>
		List<Dictionary<string, string>> GetSheetBulkData(ISheetWithBulkData sheet, int offset = 2);

		/// <summary>
		/// Retrieve the bulk data from a specific sheet in a type safe manner
		/// </summary>
		/// <typeparam name="R"></typeparam>
		/// <typeparam name="F"></typeparam>
		/// <param name="sheet"></param>
		/// <returns></returns>
		ITypedSheet<R, F> GetSheetBulkData<R, F>(ITypedSheet<R, F> sheet) where R : class, ISheetRow where F : class, ISheetFields;

		/// <summary>
		/// Write bulk data to a specific sheet
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="contents"></param>
		void SetSheetBulkData(ISheetWithBulkData sheet, List<Dictionary<string, string>> contents, int offset = 2);

		/// <summary>
		/// Write bulk data to a specific sheet in a type safe manner
		/// </summary>
		/// <typeparam name="R"></typeparam>
		/// <typeparam name="F"></typeparam>
		/// <param name="sheet"></param>
		void SetSheetData<R, F>(ITypedSheet<R, F> sheet) where R : class, ISheetRow where F : class, ISheetFields;

		/// <summary>
		/// Overload
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <returns></returns>
		string GetValueAt(ISheet sheet, string cellAddress);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc.
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <returns></returns>
		string GetValueAt(string sheet, string cellAddress);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc. 
		/// </summary>
		/// <typeparam name="T">Type hint to apply to the retrieved value</typeparam>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <returns></returns>
		T GetValueAt<T>(ISheet sheet, string cellAddress);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc. 
		/// </summary>
		/// <typeparam name="T">Type hint to apply to the retrieved value</typeparam>
		/// <param name="sheetName"></param>
		/// <param name="cellAddress"></param>
		/// <returns></returns>
		T GetValueAt<T>(string sheetName, string cellAddress);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <param name="rawString">The raw string found at the celladdress</param>
		/// <returns></returns>
		T GetValueAt<T>(ISheet sheet, string cellAddress, out string rawString);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheetName"></param>
		/// <param name="cellAddress"></param>
		/// <param name="rawString">The raw string found at the celladdress</param>
		/// <returns></returns>
		T GetValueAt<T>(string sheetName, string cellAddress, out string rawString);

		/// <summary>
		/// Overload
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <param name="value"></param>
		void SetValueAt<T>(ISheet sheet, string cellAddress, T value);

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
		void SaveToFile(string path = null, bool createDirectoryIfNotExists = false);

		/// <summary>
		/// Save this spreadsheet to a stream.
		/// </summary>
		/// <returns></returns>
		Stream SaveToStream();

		/// <summary>
		/// Returns a list of sheet names found in the opened spreadsheet
		/// </summary>
		List<string> SheetNames { get; }

		/// <summary>
		/// Returns the workbook version number and authority as out parameters. Method returns true if successful.
		/// </summary>
		/// <returns></returns>
		bool TryGetWorkbookVersion(out Version version, out string authority);
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


	public class InvalidWorkbookVersionException : Exception
	{
	}
}