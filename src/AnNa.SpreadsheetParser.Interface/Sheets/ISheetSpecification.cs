using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public interface ISheetSpecification
	{
		//AnNaSheets Sheet { get; }
		string SheetName { get; }
		List<string> ColumnNames { get; }
		int MaximumNumberOfRows { get; }
	}
}