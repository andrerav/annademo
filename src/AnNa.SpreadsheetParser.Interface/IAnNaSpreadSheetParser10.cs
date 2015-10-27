using System;
using System.Collections.Generic;
using AnNa.SpreadsheetParser.Interface.Sheets;

namespace AnNa.SpreadsheetParser.Interface
{
	public interface IAnNaSpreadSheetParser10
	{
		void OpenFile(string path, string password = null);
		List<Dictionary<string, string>> GetSheetContents(ISheetSpecification sheetSpecification);
		string GetValueAt(ISheetSpecification sheetSpecification, string cellAddress);
		string GetValueAt(string sheet, string cellAddress);
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