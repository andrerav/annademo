using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public class StowawaySheet : AbstractCrewPaxListSheet
	{

		public class Columns:CommonColumns
		{
			public const string Port_Of_Embarkation = "Port_Of_Embarkation";
			public const string Port_Of_Disembarkation = "Port_Of_Disembarkation";
		}
		public override string SheetName
		{
			get { return "Stowaway_List"; }
		}

		public override List<string> ColumnNames
		{
			get
			{
				return new List<string>
				{
					Columns.Number,
					Columns.Family_Name,
					Columns.Given_Name,
					Columns.Nationality,
					Columns.Date_Of_Birth,
					Columns.Place_Of_Birth,
					Columns.Nature_Of_Identity_Document,
					Columns.Number_Of_Identity_Document,
					Columns.Port_Of_Embarkation,
					Columns.Port_Of_Disembarkation,
					Columns.Visa_Residence_Permit_Number,
				};
			}
		}

		public override int MaximumNumberOfRows { get { return -1; } }

	}
}
