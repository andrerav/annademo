using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	[Serializable]
	public class ShipStoresSheet : ISheetWithBulkData
	{
		public class Columns: ISheetColumns
		{
			public const string NameOfArticle = "Name_Of_Article";
			public const string Quantity = "Quantity";
			public const string Number = "Number";
			public const string Unit = "Unit";
			public const string LocationOnBoard = "Location_On_Board";
			public const string OfficialUse = "Official_Use";
		}

		public string SheetName { get { return "Ship_Stores"; } }
		
		public virtual List<string> ColumnNames 
		{ 
			get
			{ 
				return new List<string>
				{
					Columns.NameOfArticle,
					Columns.Quantity,
					Columns.Number,
					Columns.Unit,
					Columns.LocationOnBoard,
					Columns.OfficialUse
				};
			} 
		}

		public int MaximumNumberOfRows { get { return -1; /* Infinite */ } }
	}
}
