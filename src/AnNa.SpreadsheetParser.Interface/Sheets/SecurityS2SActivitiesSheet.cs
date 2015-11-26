using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	[Serializable]
	public class SecurityS2SActivitiesSheet : SecuritySheet, ISheetWithBulkData 
	{
		public class Columns: ISheetColumns
		{
			[TypeHint(typeof(DateTime))]
			public const string DateFrom = "Date_from";

			[TypeHint(typeof(DateTime))]
			public const string DateTo = "Date_to";
			public const string Location = "*Location";
			public const string Latitude = "Latitude";
			public const string Longtitude = "Longtitude";
			public const string ShipToShipActivitty = "Ship_to_ship_activitty";
			public const string SecurityMeasuresAppliedInLieu = "Security_measures_applied_in_lieu";
		}

		public virtual List<string> ColumnNames
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

		public int MaximumNumberOfRows { get { return 14; } }

	}
}