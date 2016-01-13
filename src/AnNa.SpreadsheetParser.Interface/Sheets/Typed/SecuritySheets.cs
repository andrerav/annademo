﻿using AnNa.SpreadsheetParser.Interface.Attributes;
using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	[SheetVersion(SheetGroups.Security, 1, 0, "AnNa")]
	public class SecuritySheet10 : AbstractTypedSheet<ISheetRow, SecuritySheet10.SheetFieldDefinition>
	{
		public override string SheetName => "Security";

		public class SheetFieldDefinition : ISheetFields
		{
			[Field("A8", "Valid ISSC")]
			public virtual string Valid_ISSC { get; set; }
			[Field("B8", "Administration or RSO")]
			public virtual string Administration_Or_RSO { get; set; }
			[Field("C8", "ISSC Type")]
			public virtual string ISSC_Type { get; set; }
			[Field("D8", "ISSC Issuer")]
			public virtual string ISSC_Issuer { get; set; }
			[Field("E8", "ISSC Expiry Date")]
			public virtual DateTime? ISSC_Expiry_Date { get; set; }
			[Field("F8", "Reason for no valid ISSC")]
			public virtual string Reason_For_No_Valid_ISSC { get; set; }
			[Field("G8", "Security-related Matters to Report")]
			public virtual string Security_Related_Matters_To_Report { get; set; }
			[Field("A11", "Security Level")]
			public virtual string Security_Level { get; set; }
			[Field("B11", "SSP Onboard")]
			public virtual string SSP_Onboard { get; set; }

			// CSO
			[Field("C11", "CSO Family Name")]
			public virtual string CSO_Family_Name { get; set; }
			[Field("D11", "CSO Given Name")]
			public virtual string CSO_Given_Name { get; set; }
			[Field("H11", "CSO Phone 24/7")]
			public virtual string CSO_Phone_24_7 { get; set; }
			[Field("L11", "CSO Email")]
			public virtual string CSO_Email { get; set; }
			[Field("M11", "CSO Company")]
			public virtual string CSO_Company { get; set; }
		}
		
	}

	[Serializable]
	[SheetVersion(SheetGroups.SecurityPortCalls, 1, 0, "AnNa")]
	public class SecuritySheetLast10PortCalls10 : AbstractTypedSheet<SecuritySheetLast10PortCalls10.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Security";

		public override int MaximumNumberOfRows => 10;

		public class SheetRowDefinition : SheetRow
		{
			[Column("Date_of_arrival", "Date Of Arrival")]
			public virtual DateTime? Date_Of_Arrival { get; set; }
			[Column("Date_of_departure", "Date Of Departure")]
			public virtual DateTime? Date_Of_Departure { get; set; }
			[Column("Port", "Port")]
			public virtual string Port { get; set; }
			[Column("Port_facility", "Port Facility")]
			public virtual string Port_Facility { get; set; }
			[Column("Security_level", "Security Level")]
			public virtual string Security_Level { get; set; }
			[Column("Special_or_additional_security_measures_taken_by_the_ship", "Special Or Additional Security Measures Taken By The Ship")]
			public virtual string Special_Or_Additional_Security_Measures { get; set; }
		}
	}

	[Serializable]
	[SheetVersion(SheetGroups.SecurityShipToShip, 1, 0, "AnNa")]
	public class SecuritySheetS2SActivities10 : AbstractTypedSheet<SecuritySheetS2SActivities10.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Security";

		public class SheetRowDefinition : SheetRow
		{
			[Column("Date_from", "Date From")]
			public virtual DateTime? Date_From { get; set; }
			[Column("Date_to", "Date To")]
			public virtual DateTime? Date_To { get; set; }
			[Column("*Location", "Location")]
			public virtual string Location { get; set; }

			// Latitude
			[Column("Latitude_direction", "Latitude Direction")]
			public virtual string Latitude_Direction { get; set; }
			[Column("Latitude_Degrees", "Latitude Degrees")]
			public virtual int? Latitude_Degrees { get; set; }
			[Column("Latitude_Minutes", "Latitude Minutes")]
			public virtual int? Latitude_Minutes { get; set; }
			[Column("Latitude_Seconds", "Latitude Seconds")]
			public virtual int? Latitude_Seconds { get; set; }

			// Longitude
			[Column("Longtitude_Direction", "Longitude Direction")]
			public virtual string Longitude_Direction { get; set; }
			[Column("Longtitude_Degrees", "Longitude Degrees")]
			public virtual int? Longitude_Degrees { get; set; }
			[Column("Longtitude_Minutes", "Longitude Minutes")]
			public virtual int? Longitude_Minutes { get; set; }
			[Column("Longtitude_Seconds", "Longitude Seconds")]
			public virtual int? Longitude_Seconds { get; set; }

			[Column("Ship_to_ship_activitty", "Ship-to-Ship Activity")]
			public virtual string Ship_To_Ship_Activity { get; set; }
			[Column("Security_measures_applied_in_lieu", "Security Measures Applied in Lieu")]
			public virtual string Security_Measures_Applied_In_Lieu { get; set; }
		}
	}
}