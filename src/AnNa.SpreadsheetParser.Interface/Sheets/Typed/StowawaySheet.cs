using System;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	public class StowawaySheetRow : SheetRow
	{
		[Column("Number", FriendlyName = "Number")]
		public string Number;
		[Column("Family_Name", FriendlyName = "Family Name")]
		public string Family_Name;
		[Column("Given_Name", FriendlyName = "Given Name")]
		public string Given_Name;

		// Supplemental information
		[Column("Nationality", FriendlyName = "Nationality")]
		public string Nationality;
		[Column("Date_Of_Birth", FriendlyName = "Date Of Birth")]
		public DateTime? Date_Of_Birth;
		[Column("Place_Of_Birth", FriendlyName = "Place Of Birth")]
		public string Place_Of_Birth;
		[Column("Nature_Of_Identity_Document", FriendlyName = "Nature Of Identity Document")]
		public string Nature_Of_Identity_Document;
		[Column("Number_Of_Identity_Document", FriendlyName = "Number Of Identity Document")]
		public string Number_Of_Identity_Document;
		[Column("Port_Of_Embarkation", FriendlyName = "Port Of Embarkation")]
		public string Port_Of_Embarkation;
		[Column("Port_Of_Disembarkation", FriendlyName = "Port Of Disembarkation")]
		public string Port_Of_Disembarkation;
		[Column("Visa_Residence_Permit_Number", FriendlyName = "Visa/Residence Permit Number")]
		public string Visa_Residence_Permit_Number;
	}

	[Serializable]
	public class StowawaySheet : AbstractSheet<StowawaySheetRow>
	{
		public override string SheetName
		{
			get
			{
				return "Stowaway_List";
			}
		}
	}
}
