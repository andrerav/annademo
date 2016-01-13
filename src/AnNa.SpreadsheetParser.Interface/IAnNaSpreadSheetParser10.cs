using System;
using System.Collections.Generic;
using System.IO;
using AnNa.SpreadsheetParser.Interface.Sheets;
using System.Linq;
using AnNa.SpreadsheetParser.Interface.Attributes;

namespace AnNa.SpreadsheetParser.Interface
{
	public interface IAnNaSpreadSheetParser10 : IDisposable
	{

		/// <summary>
		/// Open a XLSX file from a specific path.
		/// </summary>
		/// <param name="path">Path to XSLX file</param>
		/// <param name="password">Password if spreadsheet is encryptet/protected</param>
		void OpenFile(string path, string password = null);

		/// <summary>
		/// Open a XLSX file from a specific path.
		/// </summary>
		/// <param name="stream">Path to XSLX file</param>
		/// <param name="password">Password if spreadsheet is encryptet/protected</param>
		void OpenFile(Stream stream, string password = null);

		/// <summary>
		/// Returns true if this sheet has been recognized as an AnNa spreadsheet
		/// </summary>
		/// <returns></returns>
		bool IsAnNaSpreadsheet();

		/// <summary>
		/// Retrieve the bulk data from a specific sheet
		/// </summary>
		/// <param name="sheet"></param>
		/// <returns></returns>
		List<Dictionary<string, string>> GetSheetBulkData(ISheetWithBulkData sheet, int offset = 2);

		/// <summary>
		/// Retrieve the bulk data from a specific sheet in a type safe manner
		/// </summary>
		/// <typeparam name="R"></typeparam>
		/// <typeparam name="F"></typeparam>
		/// <param name="sheet"></param>
		/// <returns></returns>
		ITypedSheet<R, F> GetSheetBulkData<R, F>(ITypedSheet<R, F> sheet) where R : class, ISheetRow where F : class, ISheetFields;

		/// <summary>
		/// Write bulk data to a specific sheet
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="contents"></param>
		void SetSheetBulkData(ISheetWithBulkData sheet, List<Dictionary<string, string>> contents, int offset = 2);

		/// <summary>
		/// Write bulk data to a specific sheet in a type safe manner
		/// </summary>
		/// <typeparam name="R"></typeparam>
		/// <typeparam name="F"></typeparam>
		/// <param name="sheet"></param>
		void SetSheetData<R, F>(ITypedSheet<R, F> sheet) where R : class, ISheetRow where F : class, ISheetFields;

		/// <summary>
		/// Overload
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <returns></returns>
		string GetValueAt(ISheet sheet, string cellAddress);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc.
		/// </summary>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <returns></returns>
		string GetValueAt(string sheet, string cellAddress);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc. 
		/// </summary>
		/// <typeparam name="T">Type hint to apply to the retrieved value</typeparam>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <returns></returns>
		T GetValueAt<T>(ISheet sheet, string cellAddress);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc. 
		/// </summary>
		/// <typeparam name="T">Type hint to apply to the retrieved value</typeparam>
		/// <param name="sheetName"></param>
		/// <param name="cellAddress"></param>
		/// <returns></returns>
		T GetValueAt<T>(string sheetName, string cellAddress);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <param name="rawString">The raw string found at the celladdress</param>
		/// <returns></returns>
		T GetValueAt<T>(ISheet sheet, string cellAddress, out string rawString);

		/// <summary>
		/// Get a value from a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheetName"></param>
		/// <param name="cellAddress"></param>
		/// <param name="rawString">The raw string found at the celladdress</param>
		/// <returns></returns>
		T GetValueAt<T>(string sheetName, string cellAddress, out string rawString);

		/// <summary>
		/// Overload
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <param name="value"></param>
		void SetValueAt<T>(ISheet sheet, string cellAddress, T value);

		/// <summary>
		/// Set a value in a specific cell in a specific sheet. Cells are addressed in the form "A1", "B3" etc.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		/// <param name="sheet"></param>
		/// <param name="cellAddress"></param>
		/// <param name="value"></param>
		void SetValueAt<T>(string sheet, string cellAddress, T value);

		/// <summary>
		/// Save this spreadsheet to a specified path, or if no path is given to the file from which this spreadsheet was opened
		/// </summary>
		/// <param name="path"></param>
		void SaveToFile(string path = null, bool createDirectoryIfNotExists = false);

		/// <summary>
		/// Save this spreadsheet to a stream.
		/// </summary>
		/// <returns></returns>
		Stream SaveToStream();

		/// <summary>
		/// Returns a list of sheet names found in the opened spreadsheet
		/// </summary>
		List<string> SheetNames { get; }

		/// <summary>
		/// Returns the workbook version number and authority as out parameters. Method returns true if successful.
		/// </summary>
		/// <returns></returns>
		bool TryGetWorkbookVersion(out Version version, out string authority);
	}

	public static class ParserExtensions
	{
		public static Dictionary<Type, object> ParseWorkbook(this IAnNaSpreadSheetParser10 parser, string path, out Version workbookVersion, out string authority)
		{
			var result = new Dictionary<Type, object>();
			parser.OpenFile(path);

			if (!parser.TryGetWorkbookVersion(out workbookVersion, out authority))
				throw new InvalidWorkbookVersionException();

			var _lambdaSafeAuthority = authority; //out or ref parameters not allowed in lambdas :\

			var type = typeof(AbstractTypedSheet<,>);

			var sheetDefinitionGroups = AppDomain.CurrentDomain.GetAssemblies()
				.SelectMany(s => s.GetTypes())
				.Where(p => !p.IsAbstract
							&& p.BaseType != null
							&& p.BaseType.IsGenericType
							&& p.BaseType.GetGenericTypeDefinition() == type
							&& p.CustomAttributes.Any(ca => ca.AttributeType == typeof(SheetVersionAttribute)))
				.Select(t =>
				{
					var versionParams = t.CustomAttributes.Single(ca => ca.AttributeType == typeof(SheetVersionAttribute)).ConstructorArguments.Select(ca => ca.Value);
					var genericTypes = t.GetInterfaces()
							.Where(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(ITypedSheet<,>))
							.Single()
							.GetGenericArguments();
					return new
					{
						Type = t,
						TypeParameters = genericTypes,
						GroupingKey = versionParams.ElementAt(0),
						Version = new Version((int)versionParams.ElementAt(1), (int)versionParams.ElementAt(2)),
						Authority = versionParams.ElementAt(3).ToString()
					};
				})
				.OrderByDescending(t => t.Authority == _lambdaSafeAuthority) //Prøv først å finne en kompatibel versjon blant authority-spesifikke definisjoner - fall tilbake på AnNa
				.ThenByDescending(t => t.Version)
				.GroupBy(g => g.GroupingKey)
				.ToList();

			var method = typeof(IAnNaSpreadSheetParser10).GetMethods().First(m => m.Name == nameof(IAnNaSpreadSheetParser10.GetSheetBulkData) && m.GetParameters().Count() == 1);
			foreach (var group in sheetDefinitionGroups)
			{
				foreach (var sheetDefinition in group)
				{
					if (sheetDefinition.Version > workbookVersion || 
						!(sheetDefinition.Authority == authority || sheetDefinition.Authority == "AnNa")) //Kun fallback til AnNa, ingen andre
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


	public class InvalidColumnPositionException : Exception
	{
		public InvalidColumnPositionException(string message)
			: base(message)
		{
		}
	}

	public class ColumnNotFoundException : Exception
	{
		public ColumnNotFoundException(string message)
			: base(message)
		{
		}
	}


	public class InvalidWorkbookVersionException : Exception
	{
	}
}