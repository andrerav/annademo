using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public abstract class SecuritySheetSpecification: ISheetSpecification
	{
		public string SheetName { get { return "Security"; } }
		public abstract List<string> ColumnNames { get; }
		public abstract int MaximumNumberOfRows { get; }
	}
}