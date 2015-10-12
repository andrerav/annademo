using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public class SecurityPortCallsSheetSpecification : SecuritySheetSpecification
	{
		public class Columns
		{
			public const string DateOfArrival = "Date_of_arrival";
			public const string DateOfDeparture = "Date_of_departure";
			public const string Port = "Port";
			public const string PortFacility = "Port_facility";
			public const string SecurityLevel = "Security_level";
			public const string SpecialOrAdditionalSecurityMeasuresTakenByTheShip = "Special_or_additional_security_measures_taken_by_the_ship";
		}

		public override List<string> ColumnNames
		{
			get
			{
				return new List<string>
				{
					Columns.DateOfArrival,
					Columns.DateOfDeparture,
					Columns.Port,
					Columns.PortFacility,
					Columns.SecurityLevel,
					Columns.SpecialOrAdditionalSecurityMeasuresTakenByTheShip,
				};
			}
		}

		public override int MaximumNumberOfRows { get { return 10; } }
	}
}