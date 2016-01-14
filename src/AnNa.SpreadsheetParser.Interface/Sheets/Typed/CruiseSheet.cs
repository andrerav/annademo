using AnNa.SpreadsheetParser.Interface.Attributes;
using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{

	[Serializable]
	[SheetVersion(SheetGroup.Cruise, 1, 0, SheetAuthority.AnNa)]
	public class CruiseSheet10 : AbstractTypedSheet<CruiseSheet10.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Cruise";

		public class SheetRowDefinition : SheetRow
		{
			[Column("*Port", "Port")]
			public virtual string Port { get; set; }

			[Column("*ETA_date", "ETA Date")]
			public virtual string ETA_Date { get; set; }

			[Column("*ETA_time", "ETA Time")]
			public virtual string ETA_Time { get; set; }
		}
	}
}
