using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	[Serializable]
	public class CruiseSheet: ISheetWithBulkData
	{

		public class Columns
		{
			public const string Port = "*Port";
			[TypeHint(typeof(DateTime))]
			public const string EtaDate = "ETA_date";
			public const string EtaTime = "ETA_time";
		}
		public virtual string SheetName
		{
			get { return "Cruise"; }
		}

		public virtual List<string> ColumnNames
		{
			get
			{
				return new List<string>
				{
					Columns.Port,
					Columns.EtaDate,
					Columns.EtaTime,
				};
			}
		}

		public int MaximumNumberOfRows { get; }
	}
}
