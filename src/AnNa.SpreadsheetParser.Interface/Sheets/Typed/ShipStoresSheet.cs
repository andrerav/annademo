using AnNa.SpreadsheetParser.Interface.Attributes;
using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	[SheetVersion(SheetGroup.ShipStores, 1, 0, SheetAuthority.AnNa)]
	public class ShipStoresSheet10 : AbstractTypedSheet<ShipStoresSheet10.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Ship_Stores";

		[Serializable]
		public class SheetRowDefinition : SheetRow
		{
			[Column("Name_Of_Article", "Name Of Article")]
			public virtual string Name_Of_Article { get; set; }

			[Column("Quantity", "Quantity")]
			public virtual decimal Quantity { get; set; }

			[Column("Unit", "Unit")]
			public virtual string Unit { get; set; }

			[Column("Location_On_Board", "Location On Board")]
			public virtual string Location_On_Board { get; set; }

			[Column("Official_Use", "Official Use")]
			public virtual string Official_Use { get; set; }
		}
	}

	[Serializable]
	[SheetVersion(SheetGroup.ShipStores, 1, 1, SheetAuthority.AnNa)]
	public class ShipStoresSheet11 : AbstractTypedSheet<ShipStoresSheet11.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Ship_Stores";

		[Serializable]
		public class SheetRowDefinition : ShipStoresSheet10.SheetRowDefinition
		{
			[Column("Description", "Description")]
			public virtual string Description { get; set; }
		}
	}
}
