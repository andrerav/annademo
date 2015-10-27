using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public class WasteSheetSpecification : ISheetSpecification
	{
		public class Columns
		{
			public const string Waste_Type = "Waste_Type";
			public const string Waste_Type_Code = "Waste_Type_Code";
			public const string Waste_Type_Description = "Waste_Type_Description";
			public const string Amount_To_Be_Delivered = "Amount_To_Be_Delivered";
			public const string Maximum_Dedicated_Storage_Capacity = "Maximum_Dedicated_Storage_Capacity";
			public const string Amount_Retained_On_Board = "Amount_Retained_On_Board";
			public const string Port_Of_Delivery_Of_Remaining_Waste = "Port_Of_Delivery_Of_Remaining_Waste";
			public const string Estimated_Amount_Of_Waste_To_Be_Generated = "Estimated_Amount_Of_Waste_To_Be_Generated";
			public const string Date_Of_Delivery_Of_Remaining_Waste = "Date_Of_Delivery_Of_Remaining_Waste";
		}

		public string SheetName { get { return "Waste_And_Residues"; } }
		public List<string> ColumnNames {
			get
			{
				return new List<string>
				{
					Columns.Waste_Type,
					Columns.Waste_Type_Code,
					Columns.Waste_Type_Description,
					Columns.Amount_To_Be_Delivered,
					Columns.Maximum_Dedicated_Storage_Capacity,
					Columns.Amount_Retained_On_Board,
					Columns.Port_Of_Delivery_Of_Remaining_Waste,
					Columns.Estimated_Amount_Of_Waste_To_Be_Generated,
					Columns.Date_Of_Delivery_Of_Remaining_Waste
				};
			}
		}
		public int MaximumNumberOfRows { get { return -1; /* Infinite */ } }
	}
}