using System;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{
	[Serializable]
	public class DpgSheetRow : SheetRow
	{
		[Column("DG_classification", FriendlyName = "Dangerous Goods Classification")]
		public string DG_Classification;
		[Column("IMO_Hazard_Class", FriendlyName = "IMO Hazard Class")]
		public string Imo_Hazard_Class;
		[Column("UN_Number", FriendlyName = "UN Number")]
		public int Un_Number;
		[Column("Transport_Unit_ID", FriendlyName = "Transport Unit ID")]
		public string Transport_Unit_Id;
		[Column("Textual_Reference", FriendlyName = "Textual Reference")]
		public string Textual_Reference;
		[Column("Stowage_position", FriendlyName = "On Board Location")]
		public string Stowage_Position;
		[Column("Gross_Quantity", FriendlyName = "Gross Quantity")]
		public decimal Gross_Quantity;
		[Column("Unit", FriendlyName = "Unit")]
		public decimal Unit;

		// Conditional Information
		[Column("Net_Quantity", FriendlyName = "Net Quantity")]
		public string Net_Quantity;
		[Column("Flashpoint", FriendlyName = "Flashpoint")]
		public decimal Flashpoint;
		[Column("MARPOL_Pollution_Code", FriendlyName = "MARPOL Pollution Code")]
		public string MARPOL_Pollution_Code;
		[Column("Port_of_loading", FriendlyName = "Port Of Loading")]
		public string Port_Of_Loading;
		[Column("Port_Of_Discharge", FriendlyName = "Port Of Discharge")]
		public string Port_Of_Discharge;
		[Column("Transport_document_ID", FriendlyName = "Transport Document ID")]
		public string Transport_Document_Id;
		[Column("Number_of_Packages", FriendlyName = "Number Of Packages")]
		public int Number_Of_Packages;
		[Column("Package_type", FriendlyName = "Package Type")]
		public string Package_Type;
		[Column("Packing_group", FriendlyName = "Packing Group")]
		public string Packing_Group;
		[Column("Subsidiary_Risks", FriendlyName = "Subsidiary Risks")]
		public string Subsidiary_Risks;
		[Column("INF_ship_class", FriendlyName = "INF Ship Class")]
		public string INF_Ship_Class;
		[Column("*Marks_&_Numbers", FriendlyName = " Marks & Numbers")]
		public string Marks_And_Numbers;
		[Column("EmS", FriendlyName = "Emergency Measures")]
		public string Emergency_Measures;
		[Column("Additional_information", FriendlyName = "Additional Information")]
		public string Additional_Information;

		// Supplemental information
		[Column("Radioactivity_level", FriendlyName = "Radioactivity Level")]
		public string Radioactivity_Level;
		[Column("Criticality", FriendlyName = "Criticality")]
		public string Criticality;
	}

	[Serializable]
	public class DpgSheet : AbstractSheet<DpgSheetRow>
	{
		public override string SheetName
		{
			get
			{
				return "Dangerous_And_Polluting_Goods";
			}
		}
	}
}
