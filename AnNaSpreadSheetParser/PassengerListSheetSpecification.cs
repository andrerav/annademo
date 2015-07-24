using System.Collections.Generic;

namespace AnNaSpreadSheetParser
{
	public class PassengerListSheetSpecification : ISheetSpecification
	{
		public AnNaSheets Sheet { get { return AnNaSheets.Pax_List; }}
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
					"Number_Of_Identity_Document",
					"Port_Of_Embarkation",
					"Port_Of_Disembarkation",
					"Transit",
					"Number",
					"Visa_Residence_Permit_Number"
				};
			}
		}
		public int MaximumNumberOfRows { get { return -1; /* Infinite */ } }
	}
}