using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public class SecurityPortCallsSheetSpecification : ISheetSpecification
	{
		public AnNaSheets Sheet { get { return AnNaSheets.Security; } }
		public List<string> ColumnNames
		{
			get
			{
				return new List<string>
				{
					"Date_of_arrival",
					"Date_of_departure",
					"Port",
					"Port_facility",
					"Security_level",
					"Special_or_additional_security_measures_taken_by_the_ship",
				};
			}
		}

		public int MaximumNumberOfRows { get { return 10; /* Infinite */ } }
	}
}