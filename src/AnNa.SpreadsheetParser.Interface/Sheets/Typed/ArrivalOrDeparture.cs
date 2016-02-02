using AnNa.SpreadsheetParser.Interface.Attributes;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{


	[Serializable]
	[SheetVersion(SheetGroup.ArrivalOrDeparture, 1, 0, SheetAuthority.AnNa)]
	public class ArrivalOrDepartureSheet10 : AbstractTypedSheet<ArrivalOrDepartureSheet10.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Arrival_Or_Departure";

		[Serializable]
		public class SheetRowDefinition : AbstractSheetRow
		{
			[Column("A_D", "Arrival Or Departure")]
			public virtual string ArrivalOrDeparture { get; set; }

			[Column("Ship_Name", "Ship Name")]
			public virtual string Ship_Name { get; set; }

			[Column("IMO_no", "IMO Number")]
			public virtual int? IMO_no { get; set; }

			[Column("Port_of_call", "Port of Call")]
			public virtual string Port_of_call { get; set; }

			[Column("ETA_port_of_call", "ETA Port of Call")]
			public virtual DateTime? ETA_port_of_call { get; set; }

			[Column("ATA_port_of_call", "ATA Port of Call")]
			public virtual DateTime? ATA_port_of_call { get; set; }

			[Column("ETD_from_port_of_call", "ETD from Port of Call")]
			public virtual DateTime? ETD_from_port_of_call { get; set; }

			[Column("ATD_from_port_of_call", "ATD from Port of Call")]
			public virtual DateTime? ATD_from_port_of_call { get; set; }

			[Column("Purpose_of_call", "Purpose of Call")]
			public virtual string Purpose_of_call { get; set; }

			[Column("No_crew", "Number of Crew")]
			public virtual int? No_crew { get; set; }

			[Column("No_pax", "Number of Pax")]
			public virtual int? No_pax { get; set; }

			[Column("Name_of_Master", "Name of Master")]
			public virtual string Name_of_Master { get; set; }

			[Column("Deep_draught", "Deep Draught")]
			public virtual double? Deep_draught { get; set; }

			[Column("Air_draught", "Air Draught")]
			public virtual double? Air_draught { get; set; }

			[Column("Contact", "Contact")]
			public virtual string Contact { get; set; }

			[Column("Contact_phone", "Contact Phone")]
			public virtual string Contact_phone { get; set; }
		}


	}
}
