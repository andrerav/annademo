using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public abstract class AbstractCrewPaxListSheetSpecification : ISheetSpecification
	{
		public class CommonColumns
		{
			public const string Number = "Number";
			public const string Family_Name = "Family_Name";
			public const string Given_Name = "Given_Name";
			public const string Nationality = "Nationality";
			public const string Date_Of_Birth = "Date_Of_Birth";
			public const string Place_Of_Birth = "Place_Of_Birth";
			public const string Nature_Of_Identity_Document = "Nature_Of_Identity_Document";
			public const string Number_Of_Identity_Document = "Number_Of_Identity_Document";
			public const string Visa_Residence_Permit_Number = "Visa_Residence_Permit_Number";
		}

		public abstract AnNaSheets Sheet { get; }
		public abstract List<string> ColumnNames { get; }
		public abstract int MaximumNumberOfRows { get; }
	}
}