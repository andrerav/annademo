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
			Assert.IsTrue(parser.GetSheetBulkData(new CrewListSheet()).Any());
		}

		[TestMethod]
		public void ReadPaxList()
		{
			Assert.IsTrue(parser.GetSheetBulkData(new PassengerListSheet()).Any());
		}
		[TestMethod]
		public void ReadWasteList()
		{
			Assert.IsTrue(parser.GetSheetBulkData(new WasteSheet()).Any());
		}

		[TestMethod]
		public void ReadLast10CallsList()
		{
			Assert.IsTrue(parser.GetSheetBulkData(new SecurityPortCallsSheet()).Any());
		}
		[TestMethod]
		public void ReadS2SList()
		{
			Assert.IsTrue(parser.GetSheetBulkData(new SecurityS2SActivitiesSheet()).Any());
		}

		[TestMethod]
		public void ReadShipStoresList()
		{
			Assert.IsTrue(parser.GetSheetBulkData(new ShipStoresSheet()).Any());
		}

		[TestMethod]
		public void ReadDPGList()
		{
			Assert.IsTrue(parser.GetSheetBulkData(new DpgSheet()).Any());
		}

		[TestMethod]
		public void ReadDPGColumns()
		{
			var sheetContents = parser.GetSheetBulkData(new DpgSheet());
			var headerRow = 5;

			// Dangerous and Polluting Cargo
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.DGClassification));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.ImoHazardClass));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.UnNumber));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.TransportUnitId));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.TextualReference));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.StowagePosition));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.GrossQuantity));

			// Conditional Information
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.NetQuantity));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.Flashpoint));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.MARPOLPollutionCode));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.PortOfLoading));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.PortOfDischarge));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.TransportDocumentId));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.NumberOfPackages));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.PackageType));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.PackingGroup));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.SubsidiaryRisks));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.INFShipClass));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.MarksAndNumbers));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.EmergencyMeasures));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.AdditionalInformation));

			// Supplemental Information
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.RadioactivityLevel));
			Assert.IsTrue(sheetContents[headerRow].ContainsKey(DpgSheet.Columns.Criticality));
		}

		[TestMethod]
		public void GetVersionByGetValueAt()
		{
			var value = parser.GetValueAt(new CrewListSheet(), "A1");
			Assert.IsTrue(!string.IsNullOrWhiteSpace(value));
			Assert.IsTrue(value == "Version: 1.0");
		}

		[TestMethod]
		public void GetVersionByGetValueAtString()
		{
			var value = parser.GetValueAt(new CrewListSheet().SheetName, "A1");
			Assert.IsTrue(!string.IsNullOrWhiteSpace(value));
			Assert.IsTrue(value == "Version: 1.0");
		}

		[TestMethod]
		public void ParseODDate1()
		{
			var items = parser.GetSheetBulkData(new SecurityPortCallsSheet());
			foreach (var item in items)
			{
				var dateOfArrivalStr = item[SecurityPortCallsSheet.Columns.DateOfArrival];
				var dateOfDepStr = item[SecurityPortCallsSheet.Columns.DateOfDeparture];

				DateTime dateOfArrival = DateTime.Parse(dateOfArrivalStr);
				DateTime dateOfDep = DateTime.Parse(dateOfDepStr);

				Assert.IsTrue(dateOfArrival < dateOfDep);
			}
		}

		[TestMethod]
		public void ParseODDate2()
		{
			var items = parser.GetSheetBulkData(new CrewListSheet());
			foreach (var item in items)
			{
				var dateOfBirthStr = item[AbstractCrewPaxListSheet.CommonColumns.Date_Of_Birth];

				if (dateOfBirthStr != null)
				{
					DateTime dateOfBirth = DateTime.Parse(dateOfBirthStr);
					Assert.IsTrue(dateOfBirth > DateTime.MinValue);
				}
			}
		}

		[TestMethod]
		public void SetValuesTest1()
		{
			var newValue = "Test1";
			var address = "A1";
			parser.SetValueAt(new CrewListSheet(), address, newValue);
			Assert.IsTrue(parser.GetValueAt(new CrewListSheet(), address) == newValue, "Values not equal");
		}

		[TestMethod]
		public void GetAndSetLast10CallsList()
		{
			var securityPortCallsSheetSpecification = new SecurityPortCallsSheet();
			var contents = parser.GetSheetBulkData(securityPortCallsSheetSpecification);
			var newValue = "Testing";
			contents[0][SecurityPortCallsSheet.Columns.SpecialOrAdditionalSecurityMeasuresTakenByTheShip] = newValue;
			Assert.IsTrue(contents.Any());
			parser.SetSheetBulkData(securityPortCallsSheetSpecification, contents);

			contents = parser.GetSheetBulkData(securityPortCallsSheetSpecification);
			Assert.IsTrue(contents[0][SecurityPortCallsSheet.Columns.SpecialOrAdditionalSecurityMeasuresTakenByTheShip] == newValue);
		}

		[TestMethod]
		public void SaveToStreamTest1()
		{
			var securityPortCallsSheetSpecification = new SecurityPortCallsSheet();
			var contents = parser.GetSheetBulkData(securityPortCallsSheetSpecification);
			Assert.IsTrue(contents.Any());
			parser.SetSheetBulkData(securityPortCallsSheetSpecification, contents);
			var stream = parser.SaveToStream();
			Assert.IsTrue(stream.Length > 0);
		}

		[TestMethod]
		public void CountSheetNames()
		{
			var sheets = parser.SheetNames;
			Assert.IsTrue(sheets.Count > 0);
		}

	}
}