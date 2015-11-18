using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
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
			if (parser == null)
			{
				parser = new T();
				parser.OpenFile("./../../AnNaTestSheet.xlsx");
			}
		}

		[TestCleanup]
		public void Cleanup()
		{
			parser.Dispose();
		}


		[TestMethod]
		public void ReadCruiseList()
		{
			Assert.IsTrue(parser.GetSheetBulkData(new CruiseSheet()).Any());
		}

		[TestMethod]
		public void ReadStowawayList()
		{
			Assert.IsTrue(parser.GetSheetBulkData(new StowawaySheet()).Any());
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

		[TestMethod]
		public void AddEntryToCrewList()
		{
			var crewListSheet = new CrewListSheet();
			var crewList = parser.GetSheetBulkData(crewListSheet);

			Assert.IsTrue(crewList.Count == 2);

			var entry = new Dictionary<string, string>();
			entry[CrewListSheet.Columns.Family_Name] = "Andersen";
			entry[CrewListSheet.Columns.Given_Name] = "Per";
			entry[CrewListSheet.Columns.Date_Of_Birth] = "12.12.1970 13:00";

			crewList.Add(entry);

			parser.SetSheetBulkData(crewListSheet, crewList);
			crewList = parser.GetSheetBulkData(crewListSheet);
			Assert.IsTrue(crewList.Count == 3);

			Assert.IsTrue(crewList.Last()[CrewListSheet.Columns.Family_Name] == "Andersen");
		}



		[TestMethod]
		public void SaveToNewFileTest1()
		{
			var path = "./../../SaveToNewFileTest1/" + Guid.NewGuid().ToString() + ".xlsx";
			parser.SaveToFile(path, true);
			Assert.IsTrue(File.Exists(path));
			File.Delete(path);
		}

		[TestMethod]
		public void Row1ShouldBeCopiedTest()
		{
			var crewListSheet = new CrewListSheet();
			var crewList = parser.GetSheetBulkData(crewListSheet);

			Assert.IsTrue(crewList.Count == 2);

			var newEntry = new Dictionary<string, string>();
			newEntry[CrewListSheet.Columns.Family_Name] = "Andersen";
			newEntry[CrewListSheet.Columns.Given_Name] = "Per";
			newEntry[CrewListSheet.Columns.Date_Of_Birth] = "12.12.1970 13:00";

			crewList.Add(newEntry);
			crewList.Add(newEntry);
			crewList.Add(newEntry);
			crewList.Add(newEntry);

			parser.SetSheetBulkData(crewListSheet, crewList);

			var newCrewList = parser.GetSheetBulkData(crewListSheet);
			Assert.IsTrue(crewList.Count == newCrewList.Count);

			Assert.IsTrue(crewList.Last()[CrewListSheet.Columns.Family_Name] == "Andersen");
			var path = "./../../SaveToNewFileTest1/" + Guid.NewGuid().ToString() + ".xlsx";
			parser.SaveToFile(path, true);

			parser.OpenFile(path);
			crewList = parser.GetSheetBulkData(crewListSheet);

			Assert.IsTrue(crewList[5][CrewListSheet.Columns.Number_Of_Identity_Document] 
							== crewList[0][CrewListSheet.Columns.Number_Of_Identity_Document]);
			File.Delete(path);


		}
	}
}