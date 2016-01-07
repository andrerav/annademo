using AnNa.SpreadsheetParser.Interface.Attributes;
using System;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	public class WasteSheetRow : SheetRow
	{
		[Column("Waste_Type", FriendlyName = "Waste Type")]
		public string Waste_Type;

		[Column("Waste_Type_Code", FriendlyName = "Waste Type Code")]
		public string Waste_Type_Code;

		[Column("Waste_Type_Description", FriendlyName = "Waste Type Description")]
		public string Waste_Type_Description;

		[Column("Amount_To_Be_Delivered", FriendlyName = "Amount to be Delivered")]
		public double Amount_To_Be_Delivered;

		[Column("Maximum_Dedicated_Storage_Capacity", FriendlyName = "Maximum Dedicated Storage Capacity")]
		public double Maximum_Dedicated_Storage_Capacity;

		[Column("Amount_Retained_On_Board", FriendlyName = "Amount Retained on Board")]
		public double Amount_Retained_On_Board;

		[Column("Port_Of_Delivery_Of_Remaining_Waste", FriendlyName = "Port of Delivery of Remaining Waste")]
		public string Port_Of_Delivery_Of_Remaining_Waste;

		[Column("Estimated_Amount_Of_Waste_To_Be_Generated", FriendlyName = "Estimated Amount of Waste to be Generated")]
		public double Estimated_Amount_Of_Waste_To_Be_Generated;

		[Column("Date_Of_Delivery_Of_Remaining_Waste", FriendlyName = "Date of Delivery of Remaining Waste")]
		public DateTime? Date_Of_Delivery_Of_Remaining_Waste;
	}

	[Serializable]
	[SheetVersion(1, 0)]
	[Deprecated]
	public class WasteSheet : AbstractSheet<WasteSheetRow>
	{
		public override string SheetName => "Waste_And_Residues";
	}

	[Serializable]
	[SheetVersion(1, 1)]
	public class WasheSheet11:WasteSheet
	{
		[Field("B4", FriendlyName = "Waste to be Delivered")]
		public string WasteToBeDeliveredCell;

		[Field("D4", FriendlyName = "Port of Last Delivery")]
		public string PortOfLastDeliveryCell;

		[Field("F4", FriendlyName = "Date of last Delivery")]
		public DateTime? DateOfLastDeliveryCell;
	}
}
