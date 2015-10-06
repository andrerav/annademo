using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public class SecurityS2SActivitiesSheetSpecification : ISheetSpecification
	{
		public AnNaSheets Sheet { get { return AnNaSheets.Security; } }
		public List<string> ColumnNames
		{
			get
			{
				return new List<string>
				{
					"Date_from",
					"Date_to",
					"*Location",
					"Latitude",
					"Longtitude",
					"Ship_to_ship_activitty",
					"Security_measures_applied_in_lieu"
				};
			}
		}

		public int MaximumNumberOfRows { get { return 14; /* Infinite */ } }
	}
}