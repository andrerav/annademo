using System;
using System.Linq;
using AnNaSpreadSheetParser;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AnNaSpreadSheetParserTest
{
	[TestClass]
	public class UnitTest1
	{
		private static AnNaSpreadSheetParser10 GetParser()
		{
			var parser = new AnNaSpreadSheetParser10();
			parser.OpenFile("./../../../lol.xlsx");
			return parser;
		}

		[TestMethod]
		public void ReadCrewList()
		{
			var parser = GetParser();
			Assert.IsTrue(parser.GetSheetContents(new CrewListSheetSpecification()).Any());
		}

		[TestMethod]
		public void ReadPaxList()
		{
			var parser = GetParser();
			Assert.IsTrue(parser.GetSheetContents(new PassengerListSheetSpecification()).Any());
		}
		[TestMethod]
		public void ReadWasteList()
		{
			var parser = GetParser();
			Assert.IsTrue(parser.GetSheetContents(new WasteSheetSpecification()).Any());
		}

		[TestMethod]
		public void ReadLast10CallsList()
		{
			var parser = GetParser();
			Assert.IsTrue(parser.GetSheetContents(new SecurityPortCallsSheetSpecification()).Any());
		}
		[TestMethod]
		public void ReadS2SList()
		{
			var parser = GetParser();
			Assert.IsTrue(parser.GetSheetContents(new SecurityS2SActivitiesSheetSpecification()).Any());
		}
	}
}
