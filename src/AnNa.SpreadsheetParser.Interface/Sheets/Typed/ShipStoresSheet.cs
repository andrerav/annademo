using System;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	public class ShipStoresSheetRow : SheetRow
	{
		[Column("Name_Of_Article", FriendlyName = "Name Of Article")]
		public string Name_Of_Article;

		[Column("Quantity", FriendlyName = "Quantity")]
		public decimal Quantity;

		[Column("Unit", FriendlyName = "Unit")]
		public string Unit;

		[Column("Description", FriendlyName = "Description")]
		public string Description;

		[Column("Location_On_Board", FriendlyName = "Location On Board")]
		public string Location_On_Board;

		[Column("Official_Use", FriendlyName = "Official Use")]
		public string Official_Use;
	}

	[Serializable]
	public class ShipStoresSheet : AbstractSheet<ShipStoresSheetRow>
	{
		public override string SheetName
		{
			get
			{
				return "Ship_Stores";
			}
		}
	}
}
