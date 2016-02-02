using AnNa.SpreadsheetParser.Interface.Attributes;
using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets.Typed
{


	[Serializable]
	[SheetVersion(SheetGroup.Dpg, 1, 0, SheetAuthority.AnNa)]
	public class DpgSheet10 : AbstractTypedSheet<DpgSheet10.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Dangerous_And_Poluting_Goods";

		[Serializable]
		public class SheetRowDefinition : AbstractSheetRow
		{
			[Column("DG_classification", "Dangerous Goods Classification")]
			public virtual string DG_Classification { get; set; }
			[Column("IMO_Hazard_Class", "IMO Hazard Class")]
			public virtual string Imo_Hazard_Class { get; set; }
			[Column("UN_Number", "UN Number")]
			public virtual int? Un_Number { get; set; }
			[Column("Transport_Unit_ID", "Transport Unit ID")]
			public virtual string Transport_Unit_Id { get; set; }
			[Column("Textual_Reference", "Textual Reference")]
			public virtual string Textual_Reference { get; set; }
			[Column("Stowage_position", "On Board Location")]
			public virtual string Stowage_Position { get; set; }
			[Column("Gross_Quantity", "Gross Quantity")]
			public virtual decimal? Gross_Quantity { get; set; }
			// Conditional Information
			[Column("Net_Quantity", "Net Quantity")]
			public virtual string Net_Quantity { get; set; }
			[Column("Flashpoint")]
			public virtual decimal? Flashpoint { get; set; }
			[Column("MARPOL_Pollution_Code", "MARPOL Pollution Code")]
			public virtual string MARPOL_Pollution_Code { get; set; }
			[Column("Port_of_loading", "Port Of Loading")]
			public virtual string Port_Of_Loading { get; set; }
			[Column("Port_Of_Discharge", "Port Of Discharge")]
			public virtual string Port_Of_Discharge { get; set; }
			[Column("Transport_document_ID", "Transport Document ID")]
			public virtual string Transport_Document_Id { get; set; }
			[Column("Number_of_Packages", "Number Of Packages")]
			public virtual int? Number_Of_Packages { get; set; }
			[Column("Package_type", "Package Type")]
			public virtual string Package_Type { get; set; }
			[Column("Packing_group", "Packing Group")]
			public virtual string Packing_Group { get; set; }
			[Column("Subsidiary_Risks", "Subsidiary Risks")]
			public virtual string Subsidiary_Risks { get; set; }
			[Column("INF_ship_class", "INF Ship Class")]
			public virtual string INF_Ship_Class { get; set; }
			[Column("*Marks_&_Numbers", " Marks & Numbers")]
			public virtual string Marks_And_Numbers { get; set; }
			[Column("EmS", "Emergency Measures")]
			public virtual string Emergency_Measures { get; set; }
			[Column("Additional_information", "Additional Information")]
			public virtual string Additional_Information { get; set; }

			// Supplemental information
			[Column("Radioactivity_level", "Radioactivity Level")]
			public virtual string Radioactivity_Level { get; set; }
			[Column("Criticality")]
			public virtual string Criticality { get; set; }

		}
	}

	[Serializable]
	[SheetVersion(SheetGroup.Dpg, 1, 1, SheetAuthority.AnNa)]
	public class DpgSheet11 : AbstractTypedSheet<DpgSheet11.SheetRowDefinition, ISheetFields>
	{
		public override string SheetName => "Dangerous_And_Polluting_Goods";

		[Serializable]
		public class SheetRowDefinition : DpgSheet10.SheetRowDefinition
		{
			[Column("Unit")]
			public virtual string Unit { get; set; }
		}
	}
}
