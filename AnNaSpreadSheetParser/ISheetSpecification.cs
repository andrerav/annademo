using System.Collections.Generic;

namespace AnNaSpreadSheetParser
{
	public interface ISheetSpecification
	{
		AnNaSheets Sheet { get; }
		List<string> ColumnNames { get; }
		int MaximumNumberOfRows { get; }
	}
}