using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public interface ISheet
	{
		string SheetName { get; }
	}


	public interface ISheetWithBulkData : ISheet
	{
		List<string> ColumnNames { get; }
		int MaximumNumberOfRows { get; }
	}
}