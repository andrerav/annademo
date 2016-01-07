using System;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	public class CruiseSheetRow : SheetRow
	{
		[Column("*Port", FriendlyName = "Port")]
		public string Port;

		[Column("*ETA_date", FriendlyName = "ETA Date")]
		public string ETA_Date;

		[Column("*ETA_time", FriendlyName = "ETA Time")]
		public string ETA_Time;
	}

	[Serializable]
	public class CruiseSheet : AbstractSheet<CruiseSheetRow>
	{
		public override string SheetName
		{
			get
			{
				return "Cruise";
			}
		}
	}
}
