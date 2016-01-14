using System;
using System.Collections.Generic;
using System.Linq;
using AnNa.SpreadsheetParser.Interface.Sheets.Typed;

namespace AnNa.SpreadsheetParser.Interface.Extensions
{

	public static class ParserExtensions
	{
		public static Dictionary<Type, object> ParseWorkbook(this IAnNaSpreadSheetParser10 parser, string path, out Version workbookVersion, out string authority)
		{
			var result = new Dictionary<Type, object>();
			parser.OpenFile(path);

			if (!parser.TryGetWorkbookVersion(out workbookVersion, out authority))
				throw new InvalidWorkbookVersionException();

			var sheetDefinitionGroups = ReflectionHelpers.GetSheetDefinitionsOrderedByAuthority(authority);

			var method = typeof(IAnNaSpreadSheetParser10).GetMethods().First(m => m.Name == nameof(IAnNaSpreadSheetParser10.GetSheetBulkData) && m.GetParameters().Count() == 1);
			foreach (var group in sheetDefinitionGroups)
			{
				foreach (var sheetDefinition in group)
				{
					if (sheetDefinition.Version > workbookVersion ||
						!(sheetDefinition.Authority == authority || sheetDefinition.Authority == SheetAuthority.AnNa)) //Kun fallback til AnNa, ingen andre
						continue;

					object contents = null;

					var instance = Activator.CreateInstance(sheetDefinition.Type);                               // ex. WasteSheet11
					var genericMethod = method.MakeGenericMethod(sheetDefinition.TypeParameters.ToArray());      // GetSheetBulkData<WasteSheet11.SheetRowDefinition, WasteSheet11.SheetFieldDefinition>

					contents = genericMethod.Invoke(parser, new object[] { instance });    // parser.GetSheetBulkData([WasteSheet11 instance])

					result.Add(sheetDefinition.Type, contents);

					break; //Hopp ut av indre-løkke etter parsing med første relevante sheet definisjon
				}
			}

			return result;
		}
	}

}
