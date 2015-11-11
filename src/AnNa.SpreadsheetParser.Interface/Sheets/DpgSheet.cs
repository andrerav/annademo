using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public class DpgSheet : ISheetWithBulkData
	{
		public class Columns: ISheetColumns
		{
			public const string DGClassification = "DG_classification";
			public const string ImoHazardClass = "IMO_Hazard_Class";
			public const string UnNumber = "UN_Number";
			public const string TransportUnitId = "Transport_Unit_ID";
			public const string TextualReference = "Textual_Reference";
			public const string StowagePosition = "Stowage_position";
			public const string GrossQuantity = "Gross_Quantity";

			public const string NetQuantity = "Net_Quantity";
			public const string Flashpoint = "Flashpoint";
			public const string MARPOLPollutionCode = "MARPOL_Pollution_Code";
			public const string PortOfLoading = "Port_of_loading";
			public const string PortOfDischarge = "Port_Of_Discharge";
			public const string TransportDocumentId = "Transport_document_ID";
			public const string NumberOfPackages = "Number_of_Packages";
			public const string PackageType = "Package_type";
			public const string PackingGroup = "Packing_group";
			public const string SubsidiaryRisks = "Subsidiary_Risks"; 
			public const string INFShipClass = "INF_ship_class";
			public const string MarksAndNumbers = "*Marks_&_Numbers";
			public const string EmergencyMeasures = "EmS";
			public const string AdditionalInformation = "Additional_information";

			public const string RadioactivityLevel = "Radioactivity_level";
			public const string Criticality = "Criticality";
		}

		public string SheetName { get { return "Dangerous_And_Poluting_Goods"; } }

		public List<string> ColumnNames
		{
			get
			{
				return new List<string>
				{
					// Dangerous and Polluting Cargo
					Columns.DGClassification,
					Columns.ImoHazardClass,
					Columns.UnNumber,
					Columns.TransportUnitId,
					Columns.TextualReference,
					Columns.StowagePosition,
					Columns.GrossQuantity,

					// Conditional Information
					Columns.NetQuantity,
					Columns.Flashpoint,
					Columns.MARPOLPollutionCode,
					Columns.PortOfLoading,
					Columns.PortOfDischarge,
					Columns.TransportDocumentId,
					Columns.NumberOfPackages,
					Columns.PackageType,
					Columns.PackingGroup,
					Columns.SubsidiaryRisks,
					Columns.INFShipClass,
					Columns.MarksAndNumbers,
					Columns.EmergencyMeasures,
					Columns.AdditionalInformation,

					// Supplemental Information
					Columns.RadioactivityLevel,
					Columns.Criticality
				};
			}
		}

		public int MaximumNumberOfRows { get { return -1; } /* infinite */ }
	}
}
