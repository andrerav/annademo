using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public abstract class SecuritySheet: ISheet
	{
		public string SheetName { get { return "Security"; } }
	}
}