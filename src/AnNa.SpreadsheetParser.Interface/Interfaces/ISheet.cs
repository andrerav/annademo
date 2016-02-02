using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public interface ISheet
	{
		string SheetName { get; }
	}

	public interface ISheetWithBulkData : ISheet
	{
		List<string> ColumnNames { get; }
		int MaximumNumberOfRows { get; }
	}

	public interface ITypedSheet<R, F> 
		: ISheet
		where R : ISheetRow 
		where F : ISheetFields
	{
		ErrorContainer ErrorContainer { get; set; }
		
		/// <summary>
		/// A collection of names of missing columns (uploaded sheet compared to sheet definition)
		/// </summary>
		List<string> MissingColumns { get; }

		int MaximumNumberOfRows { get; }
		int RowOffset { get; }
		F Fields { get; set; }
		List<R> Rows { get; set; }

		/// <summary>
		/// Adds a column name to the collection of missing columns for this sheet
		/// </summary>
		/// <param name="columnName"></param>
		void AddToMissingColumns(string columnName);

		bool ForceWrite { get; set; } //Treats all fields and columns as SkipOnWrite = false
		bool ForceRead { get; set; } //Treats all fields and columns as SkipOnRead = false
	}
}