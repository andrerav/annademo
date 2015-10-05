using System;
using AnNa.SpreadSheetParser.EPPlus;
using AnNaSpreadSheetParser;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace AnNaSpreadSheetParserTest
{
	[TestClass]
	public class AnNaSpreadSheetParserEPPlusTests: AnNaSpreadsheetParserTestBase
	{
		[TestInitialize]
		public void TestInitialize()
		{
			base.GetParser<AnNaSpreadSheetParserEPPlus>();
		}
	}
}
