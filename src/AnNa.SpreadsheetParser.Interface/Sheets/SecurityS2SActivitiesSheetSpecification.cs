using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public class SecurityS2SActivitiesSheetSpecification : SecuritySheetSpecification 
	{
		public class Columns
		{
			public const string DateFrom = "Date_from";
			public const string DateTo = "Date_to";
			public const string Location = "*Location";
			public const string Latitude = "Latitude";
			public const string Longtitude = "Longtitude";
			public const string ShipToShipActivitty = "Ship_to_ship_activitty";
			public const string SecurityMeasuresAppliedInLieu = "Security_measures_applied_in_lieu";
		}

		public override List<string> ColumnNames
		{
			get
			{
				return new List<string>
				{
					Columns.DateFrom,
					Columns.DateTo,
					Columns.Location,
					Columns.Latitude,
					Columns.Longtitude,
					Columns.ShipToShipActivitty,
					Columns.SecurityMeasuresAppliedInLieu
				};
			}
		}

		public override int MaximumNumberOfRows { get { return 14; } }

	}
}