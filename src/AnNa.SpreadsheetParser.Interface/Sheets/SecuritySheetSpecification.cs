using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public abstract class SecuritySheetSpecification: ISheetSpecification
	{
		public AnNaSheets Sheet { get { return AnNaSheets.Security; } }
		public abstract List<string> ColumnNames { get; }
		public abstract int MaximumNumberOfRows { get; }
	}
}