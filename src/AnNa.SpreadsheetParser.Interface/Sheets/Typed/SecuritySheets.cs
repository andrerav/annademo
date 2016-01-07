using System;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	public class SecuritySheet : ISheet
	{
		public string SheetName
		{
			get
			{
				return "Security";
			}
		}

		[Field("A8", FriendlyName = "Valid ISSC")]
		public string Valid_ISSC;
		[Field("B8", FriendlyName = "Administration or RSO")]
		public string Administration_Or_RSO;
		[Field("C8", FriendlyName = "ISSC Type")]
		public string ISSC_Type;
		[Field("D8", FriendlyName = "ISSC Issuer")]
		public string ISSC_Issuer;
		[Field("E8", FriendlyName = "ISSC Expiry Date")]
		public DateTime? ISSC_Expiry_Date;
		[Field("F8", FriendlyName = "Reason for no valid ISSC")]
		public string Reason_For_No_Valid_ISSC;
		[Field("G8", FriendlyName = "Security-related Matters to Report")]
		public string Security_Related_Matters_To_Report;
		[Field("A11", FriendlyName = "Security Level")]
		public string Security_Level;
		[Field("B11", FriendlyName = "SSP Onboard")]
		public string SSP_Onboard;

		// CSO
		[Field("C11", FriendlyName = "CSO Family Name")]
		public string CSO_Family_Name;
		[Field("D11", FriendlyName = "CSO Given Name")]
		public string CSO_Given_Name;
		[Field("H11", FriendlyName = "CSO Phone 24/7")]
		public string CSO_Phone_24_7;
		[Field("L11", FriendlyName = "CSO Email")]
		public string CSO_Email;
		[Field("M11", FriendlyName = "CSO Company")]
		public string CSO_Company;
	}

	[Serializable]
	public class SecuritySheetLast10PortCallsRow : SheetRow
	{
		[Column("Date_of_arrival", FriendlyName = "Date Of Arrival")]
		public DateTime? Date_Of_Arrival;
		[Column("Date_of_departure", FriendlyName = "Date Of Departure")]
		public DateTime? Date_Of_Departure;
		[Column("Port", FriendlyName = "Port")]
		public string Port;
		[Column("Port_facility", FriendlyName = "Port Facility")]
		public string Port_Facility;
		[Column("Security_level", FriendlyName = "Security Level")]
		public string Security_Level;
		[Column("Special_or_additional_security_measures_taken_by_the_ship", FriendlyName = "Special Or Additional Security Measures Taken By The Ship")]
		public string Special_Or_Additional_Security_Measures;
	}

	[Serializable]
	public class SecuritySheetS2SActivitiesRow : SheetRow
	{
		[Column("Date_from", FriendlyName = "Date From")]
		public DateTime? Date_From;
		[Column("Date_to", FriendlyName = "Date To")]
		public DateTime? Date_To;
		[Column("*Location", FriendlyName = "Location")]
		public string Location;

		// Latitude
		[Column("Latitude_direction", FriendlyName = "Latitude Direction")]
		public string Latitude_Direction;
		[Column("Latitude_Degrees", FriendlyName = "Latitude Degrees")]
		public int Latitude_Degrees;
		[Column("Latitude_Minutes", FriendlyName = "Latitude Minutes")]
		public int Latitude_Minutes;
		[Column("Latitude_Seconds", FriendlyName = "Latitude Seconds")]
		public int Latitude_Seconds;

		// Longitude
		[Column("Longtitude_Direction", FriendlyName = "Longitude Direction")]
		public string Longitude_Direction;
		[Column("Longtitude_Degrees", FriendlyName = "Longitude Degrees")]
		public int Longitude_Degrees;
		[Column("Longtitude_Minutes", FriendlyName = "Longitude Minutes")]
		public int Longitude_Minutes;
		[Column("Longtitude_Seconds", FriendlyName = "Longitude Seconds")]
		public int Longitude_Seconds;

		[Column("Ship_to_ship_activitty", FriendlyName = "Ship-to-Ship Activity")]
		public string Ship_To_Ship_Activity;
		[Column("Security_measures_applied_in_lieu", FriendlyName = "Security Measures Applied in Lieu")]
		public string Security_Measures_Applied_In_Lieu;
	}

	[Serializable]
	public class SecuritySheetLast10PortCalls : AbstractSheet<SecuritySheetLast10PortCallsRow>
	{
		public override string SheetName
		{
			get
			{
				return "Security";
			}
		}
	}

	[Serializable]
	public class SecuritySheetS2SActivities : AbstractSheet<SecuritySheetS2SActivitiesRow>
	{
		public override string SheetName
		{
			get
			{
				return "Security";
			}
		}
	}
}
