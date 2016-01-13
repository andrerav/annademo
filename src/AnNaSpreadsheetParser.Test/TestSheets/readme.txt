Test classes are written for each version of the Excel template, e.g AnNaSpreadsheetParserTest10 contains test for version 1.0. The version number is set by overriding an abstract Version property in the test class.

AnNaSpreadSheetParseTestBase.InitializeParser searches through this folder for .xlsx-files, trying to find a test sheet that matches the Version property in the test class.

Versioned test sheets must be named using the following format:
anything-before-the-hash-tag-doesnt-matter#major-minor.xlsx
