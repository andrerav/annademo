﻿using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public class ShipStoresSheetSpecification : ISheetSpecification
	{
		public class Columns
		{
			public const string NameOfArticle = "Name_Of_Article";
			public const string Quantity = "Quantity";
			public const string Number = "Number";
			public const string LocationOnBoard = "Location_On_Board";
			public const string OfficialUse = "Official_Use";
		}

		public string SheetName { get { return "Ship_Stores"; } }
		
		public List<string> ColumnNames 
		{ 
			get
			{ 
				return new List<string>
				{
					Columns.NameOfArticle,
					Columns.Quantity,
					Columns.Number,
					Columns.LocationOnBoard,
					Columns.OfficialUse
				};
			} 
		}

		public int MaximumNumberOfRows { get { return -1; /* Infinite */ } }
	}
}
