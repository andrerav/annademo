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

	public interface ITypedSheetWithBulkData<T> : ISheet where T : ISheetRow
	{
		SyntaxErrorContainer SyntaxErrorContainer { get; }
		List<T> Rows { get; set; }
		int MaximumNumberOfRows { get; }
		int RowOffset { get; }
	}
}