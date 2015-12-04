using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	public class PersonOnBoardRow : SheetRow
	{
		[Column("Family_Name", FriendlyName = "Family name")]
		public string Family_Name;

		[Column("Given_Name", FriendlyName = "Given name(s)")]
		public string Given_Name;

		[Column("Nationality")]
		public string Nationality;

		[Column("Date_Of_Birth")]
		public DateTime? Date_Of_Birth;

		[Column("Place_Of_Birth")]
		public string Place_Of_Birth;

		[Column("Nature_Of_Identity_Document")]
		public string Nature_Of_Identity_Document;

		[Column("Number_Of_Identity_Document")]
		public string Number_Of_Identity_Document;

		[Column("Visa_Residence_Permit_Number")]
		public string Visa_Residence_Permit_Number;

		[Column("Number", Ignorable = true)]
		public string Number;
	}

	[Serializable]
	public class CrewListRow : PersonOnBoardRow
	{

		[Column("Duty_Of_Crew")]
		public string Duty_Of_Crew;

		[Column("Gender")]
		public string Gender;

		[Column("Crew_Effects")]
		public string Crew_Effects;
	}

	[Serializable]
	public class PassengerListRow : PersonOnBoardRow
	{
		[Column("Port_Of_Embarkation")]
		public string Port_Of_Embarkation;
		[Column("Port_Of_Disembarkation")]
		public string Port_Of_Disembarkation;
		[Column("Transit")]
		public string Transit;
	}


	[Serializable]
	public class CrewListSheet : AbstractSheet<CrewListRow>
	{
		public override string SheetName
		{
			get
			{
				return "Crew_List";
			}
		}
	}

	[Serializable]
	public class PassengerListSheet : AbstractSheet<PassengerListRow>
	{
		public override string SheetName
		{
			get
			{
				return "Pax_List";
			}
		}
	}
}
