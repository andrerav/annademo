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
			parser.OpenFile("./../../TestSheet.xlsx");
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
	}
}