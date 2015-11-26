using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	[Serializable]
	public class SecurityPortCallsSheet : SecuritySheet, ISheetWithBulkData
	{
		public class Columns:ISheetColumns
		{
			[TypeHint(typeof(DateTime))]
			public const string DateOfArrival = "Date_of_arrival";

			[TypeHint(typeof(DateTime))]
			public const string DateOfDeparture = "Date_of_departure";

			public const string Port = "Port";
			public const string PortFacility = "Port_facility";
			public const string SecurityLevel = "Security_level";
			public const string SpecialOrAdditionalSecurityMeasuresTakenByTheShip = "Special_or_additional_security_measures_taken_by_the_ship";
		}

		public virtual List<string> ColumnNames
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

		public int MaximumNumberOfRows { get { return 10; } }
	}
}