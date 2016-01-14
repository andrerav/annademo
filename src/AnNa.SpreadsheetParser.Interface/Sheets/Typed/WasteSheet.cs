using AnNa.SpreadsheetParser.Interface.Attributes;
using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{

	[Serializable]
	[SheetVersion(SheetGroup.Waste, 1, 0, SheetAuthority.AnNa)]
	[Deprecated]
	public class WasteSheet10 : AbstractTypedSheet<WasteSheet10.SheetRowDefinition, ISheetFields>	{
		public override string SheetName => "Waste_And_Residues";

		public class SheetRowDefinition : SheetRow
		{
			[Column("Waste_Type", "Waste Type")]
			public virtual string Waste_Type { get; set; }

			[Column("Waste_Type_Code", "Waste Type Code")]
			public virtual string Waste_Type_Code { get; set; }

			[Column("Waste_Type_Description", "Waste Type Description")]
			public virtual string Waste_Type_Description { get; set; }

			[Column("Amount_To_Be_Delivered", "Amount to be Delivered")]
			public virtual double? Amount_To_Be_Delivered { get; set; }

			[Column("Maximum_Dedicated_Storage_Capacity", "Maximum Dedicated Storage Capacity")]
			public virtual double? Maximum_Dedicated_Storage_Capacity { get; set; }

			[Column("Amount_Retained_On_Board", "Amount Retained on Board")]
			public virtual double? Amount_Retained_On_Board { get; set; }

			[Column("Port_Of_Delivery_Of_Remaining_Waste", "Port of Delivery of Remaining Waste")]
			public virtual string Port_Of_Delivery_Of_Remaining_Waste { get; set; }

			[Column("Estimated_Amount_Of_Waste_To_Be_Generated", "Estimated Amount of Waste to be Generated")]
			public virtual double? Estimated_Amount_Of_Waste_To_Be_Generated { get; set; }

			[Column("Date_Of_Delivery_Of_Remaining_Waste", "Date of Delivery of Remaining Waste")]
			public virtual DateTime? Date_Of_Delivery_Of_Remaining_Waste { get; set; }
		}
	}

	[Serializable]
	[SheetVersion(SheetGroup.Waste, 1, 1, SheetAuthority.AnNa)]
	public class WasteSheet11 : AbstractTypedSheet<WasteSheet11.SheetRowDefinition, WasteSheet11.SheetFieldDefinition>
	{
		public override string SheetName => "Waste_And_Residues";
		public class SheetFieldDefinition : ISheetFields
		{
			[Field("B4", "Waste to be Delivered")]
			public virtual string WasteToBeDeliveredCell { get; set; }

			[Field("D4", "Port of Last Delivery")]
			public virtual string PortOfLastDeliveryCell { get; set; }

			[Field("F4", "Date of last Delivery")]
			public virtual DateTime DateOfLastDeliveryCell { get; set; }
		}

		public class SheetRowDefinition : WasteSheet10.SheetRowDefinition
		{
			[Removed]
			public override string Waste_Type_Code
			{
				get; set;
			}
		}
	}
}
