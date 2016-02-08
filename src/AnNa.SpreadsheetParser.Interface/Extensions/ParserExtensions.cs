using System;
using System.Collections.Generic;
using System.Linq;
using AnNa.SpreadsheetParser.Interface.Sheets.Typed;
using System.Reflection;

namespace AnNa.SpreadsheetParser.Interface.Extensions
{

	public static class ParserExtensions
	{
		public static Dictionary<ReflectionHelpers.SheetDefinitionMetaData, object> ParseWorkbook(this IAnNaSpreadSheetParser10 parser, Version workbookVersion, string authority)
		{
			var result = new Dictionary<ReflectionHelpers.SheetDefinitionMetaData, object>();

			parser.ThrowExceptionIfNotInitialized();

			var sheetDefinitionGroups = ReflectionHelpers.GetSheetDefinitionsOrderedByAuthority(authority);

			var sheetNames = parser.GetSheetNames();

			var method = typeof(IAnNaSpreadSheetParser10).GetMethods().First(m => m.Name == nameof(IAnNaSpreadSheetParser10.GetSheetBulkData) && m.GetParameters().Count() == 1);
			foreach (var group in sheetDefinitionGroups)
			{
				foreach (var sheetMetaData in group)
				{
					if (sheetMetaData.Version > workbookVersion ||
						!(sheetMetaData.Authority == authority || sheetMetaData.Authority == SheetAuthority.AnNa)) //Fallback to AnNa definitions
						continue;

					object contents = null;

					//Check if workbook contains sheet
					if (sheetNames.All(s => s.ToLowerInvariant() != sheetMetaData.SheetName.ToLowerInvariant()))
						continue;

					var instance = Activator.CreateInstance(sheetMetaData.Type);                               // ex. WasteSheet11

					var genericMethod = method.MakeGenericMethod(sheetMetaData.TypeParameters.ToArray());      // GetSheetBulkData<WasteSheet11.SheetRowDefinition, WasteSheet11.SheetFieldDefinition>

					contents = genericMethod.Invoke(parser, new object[] { instance });    // parser.GetSheetBulkData([WasteSheet11 instance])


					result.Add(sheetMetaData, contents);

					break; //Break inner lopp after parsing with first compatible sheet definition
				}
			}

			return result;
		}
	}

}
