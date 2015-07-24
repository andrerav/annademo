using System.Collections.Generic;

namespace AnNaSpreadSheetParser
{
	public class WasteSheetSpecification : ISheetSpecification
	{
		public AnNaSheets Sheet { get { return AnNaSheets.Waste_And_Residues; }}
		public List<string> ColumnNames {
			get
			{
				return new List<string>
				{
					"Waste_Type",
					"Waste_Type_Code",
					"Waste_Type_Description",
					"Amount_To_Be_Delivered",
					"Maximum_Dedicated_Storage_Capacity",
					"Amount_Retained_On_Board",
					"Port_Of_Delivery_Of_Remaining_Waste",
					"Estimated_Amount_Of_Waste_To_Be_Generated",
					"Date_Of_Delivery_Of_Remaining_Waste"
				};
			}
		}
		public int MaximumNumberOfRows { get { return -1; /* Infinite */ } }
	}
}