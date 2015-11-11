using System;
using System.Collections.Generic;
using System.IO;
using AnNa.SpreadsheetParser.Interface.Sheets;

namespace AnNa.SpreadsheetParser.Interface
{
	public interface IAnNaSpreadSheetParser10
	{
		void OpenFile(string path, string password = null);
		List<Dictionary<string, string>> GetSheetContents(ISheetSpecification sheetSpecification);
		void SetSheetContents(ISheetSpecification sheetSpecification, List<Dictionary<string, string>> contents);
		string GetValueAt(ISheetSpecification sheetSpecification, string cellAddress);
		string GetValueAt(string sheet, string cellAddress);
		void SetValueAt<T>(ISheetSpecification sheetSpecification, string cellAddress, T value);
		void SetValueAt<T>(string sheet, string cellAddress, T value);
		void SaveToFile(string path = null);
		Stream SaveToStream();
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