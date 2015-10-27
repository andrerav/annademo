using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public class CrewListSheetSpecification : AbstractCrewPaxListSheetSpecification
	{
		public class Columns : CommonColumns
		{
			public const string Duty_Of_Crew = "Duty_Of_Crew";
			public const string Gender = "Gender";
			public const string Crew_Effects = "Crew_Effects";
		}

		public override string SheetName
		{
			get { return "Crew_List"; }
		}

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
					Columns.Duty_Of_Crew,
					Columns.Number,
					Columns.Gender,
					Columns.Visa_Residence_Permit_Number,
					Columns.Crew_Effects
				};
			}
		}
	
		public override int MaximumNumberOfRows { get { return -1; /* Infinite */ } }
	}
}