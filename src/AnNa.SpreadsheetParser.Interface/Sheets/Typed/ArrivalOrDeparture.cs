using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{

	[Serializable]
	public class ArrivalOrDepartureSheetRow : SheetRow
	{
		[Column("A_D")]
		public string ArrivalOrDeparture;

		[Column("Ship_Name")]
		public string Ship_Name;

		[Column("IMO_no")]
		public int IMO_no;

		[Column("Port_of_call")]
		public string Port_of_call;

		[Column("ETA_port_of_call")]
		public DateTime ETA_port_of_call;

		[Column("ATA_port_of_call")]
		public DateTime? ATA_port_of_call;

		[Column("ETD_from_port_of_call")]
		public DateTime ETD_from_port_of_call;

		[Column("ATD_from_port_of_call")]
		public DateTime? ATD_from_port_of_call;

		[Column("Purpose_of_call")]
		public string Purpose_of_call;

		[Column("No_crew")]
		public int? No_crew;

		[Column("No_pax")]
		public int? No_pax;

		[Column("Name_of_Master")]
		public string Name_of_Master;

		[Column("Deep_draught")]
		public double? Deep_draught;

		[Column("Air_draught")]
		public double? Air_draught;

		[Column("Contact")]
		public string Contact;

		[Column("Contact_phone")]
		public string Contact_phone;
	}

	[Serializable]
	public class ArrivalOrDepartureSheet : AbstractSheet<ArrivalOrDepartureSheetRow>
	{
		public override string SheetName
		{
			get
			{
				return "Arrival_Or_Departure";
			}
		}
	}
}
