using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public class PassengerListSheet : AbstractCrewPaxListSheet
	{
		public class Columns : CommonColumns
		{

			public const string Port_Of_Embarkation = "Port_Of_Embarkation";
			public const string Port_Of_Disembarkation = "Port_Of_Disembarkation";
			public const string Transit = "Transit";
			public new const string Number = "Number";
		}

		public override string SheetName { get { return "Pax_List"; } }

		public override List<string> ColumnNames {
			get
			{
				return new List<string>
				{
					Columns.Family_Name,
					Columns.Given_Name,
					Columns.Nationality,
					Columns.Date_Of_Birth,
					Columns.Place_Of_Birth,
					Columns.Nature_Of_Identity_Document,
					Columns.Number_Of_Identity_Document,
					Columns.Port_Of_Embarkation,
					Columns.Port_Of_Disembarkation,
					Columns.Transit,
					Columns.Number,
					Columns.Visa_Residence_Permit_Number
				};
			}
		}
		public override int MaximumNumberOfRows { get { return -1; /* Infinite */ } }
	}
}