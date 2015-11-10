using System;
using System.Linq;
using AnNa.SpreadsheetParser.Interface;
using AnNa.SpreadsheetParser.Interface.Sheets;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AnNaSpreadSheetParserTest
{
	public class AnNaSpreadsheetParserTestBase
	{
		protected IAnNaSpreadSheetParser10 parser;

		protected virtual void GetParser<T>() where T: class, IAnNaSpreadSheetParser10, new()
		{
			parser = new T();
			parser.OpenFile("./../../AnNaTestSheet.xlsx");
		}

		[TestMethod]
		public void ReadCrewList()
		{
			Assert.IsTrue(parser.GetSheetContents(new CrewListSheetSpecification()).Any());
		}

		[TestMethod]
		public void ReadPaxList()
		{
			Assert.IsTrue(parser.GetSheetContents(new PassengerListSheetSpecification()).Any());
		}
		[TestMethod]
		public void ReadWasteList()
		{
			Assert.IsTrue(parser.GetSheetContents(new WasteSheetSpecification()).Any());
		}

		[TestMethod]
		public void ReadLast10CallsList()
		{
			Assert.IsTrue(parser.GetSheetContents(new SecurityPortCallsSheetSpecification()).Any());
		}
		[TestMethod]
		public void ReadS2SList()
		{
			Assert.IsTrue(parser.GetSheetContents(new SecurityS2SActivitiesSheetSpecification()).Any());
		}

		[TestMethod]
		public void ReadShipStoresList()
		{
			Assert.IsTrue(parser.GetSheetContents(new ShipStoresSheetSpecification()).Any());
		}

		[TestMethod]
		public void ReadDPGList()
		{
			Assert.IsTrue(parser.GetSheetContents(new DPGSheetSpecification()).Any());
		}

		[TestMethod]
		public void ReadDPGColumns()
		{
			var sheetContents = parser.GetSheetContents(new DPGSheetSpecification());
			var headerRow = 5;

			// Dangerous and Polluting Cargo
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.DGClassification));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.ImoHazardClass));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.UnNumber));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.TransportUnitId));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.TextualReference));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.StowagePosition));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.GrossQuantity));

			// Conditional Information
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.NetQuantity));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.Flashpoint));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.MARPOLPollutionCode));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.PortOfLoading));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.PortOfDischarge));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.TransportDocumentId));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.NumberOfPackages));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.PackageType));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.PackingGroup));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.SubsidiaryRisks));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.INFShipClass));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.MarksAndNumbers));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.EmergencyMeasures));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.AdditionalInformation));

			// Supplemental Information
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.RadioactivityLevel));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DPGSheetSpecification.Columns.Criticality));
		}

		[TestMethod]
		public void GetVersionByGetValueAt()
		{
			var value = parser.GetValueAt(new CrewListSheetSpecification(), "A1");
			Assert.IsTrue(!string.IsNullOrWhiteSpace(value));
			Assert.IsTrue(value == "Version: 1.0");
		}

		[TestMethod]
		public void GetVersionByGetValueAtString()
		{
			var value = parser.GetValueAt(new CrewListSheetSpecification().SheetName, "A1");
			Assert.IsTrue(!string.IsNullOrWhiteSpace(value));
			Assert.IsTrue(value == "Version: 1.0");
		}

		[TestMethod]
		public void ParseODDate1()
		{
			var items = parser.GetSheetContents(new SecurityPortCallsSheetSpecification());
			foreach (var item in items)
			{
				var dateOfArrivalStr = item[SecurityPortCallsSheetSpecification.Columns.DateOfArrival];
				var dateOfDepStr = item[SecurityPortCallsSheetSpecification.Columns.DateOfDeparture];

				DateTime dateOfArrival = DateTime.Parse(dateOfArrivalStr);
				DateTime dateOfDep = DateTime.Parse(dateOfDepStr);

				Assert.IsTrue(dateOfArrival < dateOfDep);
			}
		}

		[TestMethod]
		public void ParseODDate2()
		{
			var items = parser.GetSheetContents(new CrewListSheetSpecification());
			foreach (var item in items)
			{
				var dateOfBirthStr = item[AbstractCrewPaxListSheetSpecification.CommonColumns.Date_Of_Birth];

				if (dateOfBirthStr != null)
				{
					DateTime dateOfBirth = DateTime.Parse(dateOfBirthStr);
					Assert.IsTrue(dateOfBirth > DateTime.MinValue);
				}
			}
		}
	}
}