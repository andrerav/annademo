using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	[Serializable]
	public abstract class SecuritySheet: ISheet
	{
		public string SheetName { get { return "Security"; } }
	}
}