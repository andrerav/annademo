using System.Collections.Generic;

namespace AnNaSpreadSheetParser
{
	public class CrewListSheetSpecification : ISheetSpecification
	{
		public AnNaSheets Sheet { get { return AnNaSheets.Crew_List; }}
		public List<string> ColumnNames {
			get
			{
				return new List<string>
				{
					"Family_Name",
					"Given_Name",
					"Nationality",
					"Date_Of_Birth",
					"Place_Of_Birth",
					"Nature_Of_Identity_Document",
					"Number_Of_Identity_Document",
					"Duty_Of_Crew",
					"Number",
					"Gender",
					"Visa_Residence_Permit_Number",
					"Crew_Effects"
				};
			}
		}
	
		public int MaximumNumberOfRows { get { return -1; /* Infinite */ } }
	}
}