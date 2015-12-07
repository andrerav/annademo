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

		protected virtual void GetParser<T>() where T : class, IAnNaSpreadSheetParser10, new()
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
			//parser.Dispose();
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
		public void ReadArrivalOrDepartureList()
		{
			var rows = parser.GetSheetBulkData(new AnNa.SpreadsheetParser.Interface.Sheets.Typed.ArrivalOrDepartureSheet()).Rows;
			Assert.IsTrue(rows.Any());
		}


		[TestMethod]
		public void ApplicationOfDateTimeTypeHint()
		{
			var data = parser.GetSheetBulkData(new PassengerListSheet());
			foreach (var row in data)
			{
				var dob1 = row[PassengerListSheet.Columns.Date_Of_Birth];
				if (!string.IsNullOrWhiteSpace(dob1))
				{
					object dobConverted;
					var typeHintedDob = Util.ApplyTypeHint<DateTime>(dob1, out dobConverted);
					Assert.IsTrue(dobConverted is DateTime);
					Assert.IsTrue(typeHintedDob.ToString() == dobConverted.ToString());
					Assert.IsTrue((DateTime)dobConverted > DateTime.MinValue);
				}
			}
		}

		[TestMethod]
		public void ReadDPGColumns()
		{
			var sheetContents = parser.GetSheetBulkData(new DpgSheet());
			var row = sheetContents.First();

			// Dangerous and Polluting Cargo
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.DGClassification));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.ImoHazardClass));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.UnNumber));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.TransportUnitId));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.TextualReference));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.StowagePosition));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.GrossQuantity));

			// Conditional Information
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.NetQuantity));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.Flashpoint));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.MARPOLPollutionCode));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.PortOfLoading));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.PortOfDischarge));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.TransportDocumentId));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.NumberOfPackages));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.PackageType));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.PackingGroup));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.SubsidiaryRisks));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.INFShipClass));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.MarksAndNumbers));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.EmergencyMeasures));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.AdditionalInformation));

			// Supplemental Information
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.RadioactivityLevel));
			Assert.IsTrue(row.ContainsKey(DpgSheet.Columns.Criticality));
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
			Assert.IsTrue(contents[0][SecurityPortCallsSheet.Columns.SpecialOrAdditionalSecurityMeasuresTakenByTheShip] ==
						  newValue);
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
			var previousCount = crewList.Count;
            Assert.IsTrue(previousCount > 0);

			var entry = new Dictionary<string, string>();
			entry[CrewListSheet.Columns.Family_Name] = "Andersen";
			entry[CrewListSheet.Columns.Given_Name] = "Per";
			entry[CrewListSheet.Columns.Date_Of_Birth] = "12.12.1970 13:00";

			crewList.Add(entry);

			parser.SetSheetBulkData(crewListSheet, crewList);
			crewList = parser.GetSheetBulkData(crewListSheet);
			Assert.IsTrue(crewList.Count == previousCount + 1);

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
			//var crewListSheet = new CrewListSheet();
			//var crewList = parser.GetSheetBulkData(crewListSheet);

			//Assert.IsTrue(crewList.Count == 2);

			//var newEntry = new Dictionary<string, string>();
			//newEntry[CrewListSheet.Columns.Family_Name] = "Andersen";
			//newEntry[CrewListSheet.Columns.Given_Name] = "Per";
			//newEntry[CrewListSheet.Columns.Date_Of_Birth] = "12.12.1970 13:00";

			//crewList.Add(newEntry);
			//crewList.Add(newEntry);
			//crewList.Add(newEntry);
			//crewList.Add(newEntry);

			//parser.SetSheetBulkData(crewListSheet, crewList);

			//var newCrewList = parser.GetSheetBulkData(crewListSheet);
			//Assert.IsTrue(crewList.Count == newCrewList.Count);

			//Assert.IsTrue(crewList.Last()[CrewListSheet.Columns.Family_Name] == "Andersen");
			//var path = "./../../SaveToNewFileTest1/" + Guid.NewGuid().ToString() + ".xlsx";
			//parser.SaveToFile(path, true);

			//parser.OpenFile(path);
			//crewList = parser.GetSheetBulkData(crewListSheet);

			//Assert.IsTrue(crewList[5][CrewListSheet.Columns.Number_Of_Identity_Document]
			//              == crewList[0][CrewListSheet.Columns.Number_Of_Identity_Document]);
			//File.Delete(path);


		}

		[TestMethod]
		public void OpenFromStreamTest()
		{
			var stream = parser.SaveToStream();
			parser.OpenFile(stream);
			Assert.IsTrue(parser.GetSheetBulkData(new CrewListSheet()).Any());
			Assert.IsTrue(parser.GetSheetBulkData(new PassengerListSheet()).Any());
			Assert.IsTrue(parser.GetSheetBulkData(new WasteSheet()).Any());

		}

		[TestMethod]
		public void SerializeTest()
		{
			var crewlistSheet = new CrewListSheet();
			System.Xml.Serialization.XmlSerializer x = new System.Xml.Serialization.XmlSerializer(crewlistSheet.GetType());
			StringWriter sw = new StringWriter();
			x.Serialize(sw, crewlistSheet);
			Assert.IsTrue(sw.ToString().Length > 0);
		}

		[TestMethod]
		public void SheetNamesTest()
		{
			Assert.IsTrue(parser.SheetNames.Count > 0);
		}

		[TestMethod]
		public void IgnoreableRowValueTest()
		{

			var column = new SheetColumn();
			column.ColumnName = "Test1";
			column.FieldName = "Test1";
			column.FieldType = typeof(string);
			const string v1 = "ABC123";
			const string v2 = "DEF456";
			column.IgnoreableValues = new string[] { v1, v2 };

			Assert.IsTrue(Util.IsIgnorableValue(v1, column));
			Assert.IsTrue(Util.IsIgnorableValue(v2, column));
			Assert.IsFalse(Util.IsIgnorableValue(Guid.NewGuid().ToString(), column));
		}
	}
}