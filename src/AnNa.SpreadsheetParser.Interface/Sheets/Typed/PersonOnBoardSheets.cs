using AnNa.SpreadsheetParser.Interface.Attributes;
using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	public abstract class PersonOnBoardRow : SheetRow
	{
		[Column("Family_Name", "Family name")]
		public virtual string Family_Name { get; set; }

		[Column("Given_Name", "Given name(s)")]
		public virtual string Given_Name { get; set; }

		[Column("Nationality")]
		public virtual string Nationality { get; set; }

		[Column("Date_Of_Birth", "Date of Birth")]
		public virtual DateTime? Date_Of_Birth { get; set; }

		[Column("Place_Of_Birth", "Place of Birth")]
		public virtual string Place_Of_Birth { get; set; }

		[Column("Nature_Of_Identity_Document", "Nature of Idendity Document")]
		public virtual string Nature_Of_Identity_Document { get; set; }

		[Column("Number_Of_Identity_Document", "Number of Identity Document")]
		public virtual string Number_Of_Identity_Document { get; set; }

		[Column("Visa_Residence_Permit_Number", "Visa Residence Permit Number")]
		public virtual string Visa_Residence_Permit_Number { get; set; }

		[Column("Number", Ignorable = true)]
		public virtual string Number { get; set; }
	}

	[Serializable]
	[SheetVersion(SheetGroup.CrewList, 1, 0, SheetAuthority.AnNa)]
	public class CrewListSheet10 : AbstractTypedSheet<CrewListSheet10.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Crew_List";

		public class SheetRowDefinition : PersonOnBoardRow
		{
			[Column("Duty_Of_Crew", "Duty of Crew")]
			public virtual string Duty_Of_Crew { get; set; }

			[Column("Gender")]
			public virtual string Gender { get; set; }

			[Column("Crew_Effects", "Crew Effects")]
			public virtual string Crew_Effects { get; set; }
		}
	}

	[Serializable]
	[SheetVersion(SheetGroup.PaxList, 1, 0, SheetAuthority.AnNa)]
	public class PassengerListSheet10 : AbstractTypedSheet<PassengerListSheet10.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Pax_List";

		public class SheetRowDefinition : PersonOnBoardRow
		{
			[Column("Port_Of_Embarkation", "Port of Embarkation")]
			public virtual string Port_Of_Embarkation { get; set; }

			[Column("Port_Of_Disembarkation", "Port of Disembarkation")]
			public virtual string Port_Of_Disembarkation { get; set; }

			[Column("Transit")]
			public virtual string Transit { get; set; }
		}
	}
}
