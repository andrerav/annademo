using AnNa.SpreadsheetParser.Interface.Attributes;
using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	[SheetVersion(SheetGroups.Stowaway, 1, 0, "AnNa")]
	public class StowawaySheet10 : AbstractTypedSheet<StowawaySheet10.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Stowaway_List";

		public class SheetRowDefinition : SheetRow
		{
			[Column("Number", "Number")]
			public virtual string Number { get; set; }
			[Column("Family_Name", "Family Name")]
			public virtual string Family_Name { get; set; }
			[Column("Given_Name", "Given Name")]
			public virtual string Given_Name { get; set; }

			// Supplemental information
			[Column("Nationality", "Nationality")]
			public virtual string Nationality { get; set; }
			[Column("Date_Of_Birth", "Date Of Birth")]
			public virtual DateTime? Date_Of_Birth { get; set; }
			[Column("Place_Of_Birth", "Place Of Birth")]
			public virtual string Place_Of_Birth { get; set; }
			[Column("Nature_Of_Identity_Document", "Nature Of Identity Document")]
			public virtual string Nature_Of_Identity_Document { get; set; }
			[Column("Number_Of_Identity_Document", "Number Of Identity Document")]
			public virtual string Number_Of_Identity_Document { get; set; }
			[Column("Port_Of_Embarkation", "Port Of Embarkation")]
			public virtual string Port_Of_Embarkation { get; set; }
			[Column("Port_Of_Disembarkation", "Port Of Disembarkation")]
			public virtual string Port_Of_Disembarkation { get; set; }
			[Column("Visa_Residence_Permit_Number", "Visa/Residence Permit Number")]
			public virtual string Visa_Residence_Permit_Number { get; set; }
		}
	}
}
