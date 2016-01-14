using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	public abstract partial class SheetGroup
	{
		public const string ArrivalOrDeparture = "ArrivalOrDeparture";
		public const string CrewList = "CrewList";
		public const string Cruise = "Cruise";
		public const string Dpg = "Dpg";
		public const string PaxList = "PaxList";
		public const string Security = "Security";
		public const string SecurityPortCalls = "SecurityPortCalls";
		public const string SecurityShipToShip = "SecurityShipToShip";
		public const string ShipStores = "ShipStores";
		public const string Stowaway = "Stowaway";
		public const string Waste = "Waste";
	}

	public abstract partial class SheetAuthority
	{
		public const string AnNa = "AnNa";
	}
}
