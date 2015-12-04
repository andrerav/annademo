using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	[Serializable]
	public abstract class SheetRow
	{
		public int RowIndex { get; set; }
		public SyntaxErrorContainer SyntaxErrorContainer = new SyntaxErrorContainer();
	}

	[Serializable]
	public abstract class AbstractSheet<T> : ITypedSheetWithBulkData<T> where T : SheetRow
	{
		public SyntaxErrorContainer SyntaxErrorContainer => new SyntaxErrorContainer();
		public int MaximumNumberOfRows { get { return -1; } }
		public int RowOffset { get { return 2; } }
		public List<T> Rows { get; set; }
		public abstract string SheetName { get; }
	}

	[Serializable]
	public class SyntaxErrorContainer
	{
		/// <summary>
		/// A list of syntax errors
		/// </summary>
		public List<SyntaxError> SyntaxErrors { get; private set; }

		/// <summary>
		/// Add a syntax error to this container. Note that the DataField property must be set, otherwise an exception is thrown.
		/// </summary>
		/// <param name="error"></param>
		public void AddSyntaxError(SyntaxError error)
		{
			if (error.DataField == null)
			{
				throw new ArgumentNullException(nameof(error.DataField));
			}

			if (SyntaxErrors == null)
			{
				SyntaxErrors = new List<SyntaxError>();
			}
			SyntaxErrors.Add(error);
		}
	}

	[Serializable]
	public class SyntaxError
	{

		/// <summary>
		/// The cell in which this syntax error was found
		/// </summary>
		public string CellAddress;

		/// <summary>
		/// The original hinted type which could not be parsed
		/// </summary>
		public Type TypeHint;

		/// <summary>
		/// The raw value from the spreadsheet
		/// </summary>
		public string RawValue;

		/// <summary>
		/// The data field which contains erronous data
		/// </summary>
		public SheetDataField DataField;
	}

	[Serializable]
	public abstract class SheetDataField
	{
		/// <summary>
		/// Field type, ie. string, datetime, double, int, etc
		/// </summary>
		public Type FieldType;

		/// <summary>
		/// The name of the field in code
		/// </summary>
		public string FieldName;

		/// <summary>
		/// A friendly name which is recognizable to an end user (typically retrieved from the GUI)
		/// </summary>
		public string FriendlyName;
	}

	[Serializable]
	public class SheetColumn : SheetDataField
	{
		/// <summary>
		/// Column name used in the spreadsheet
		/// </summary>
		public string ColumnName;
	}

	[Serializable]
	public class SheetField : SheetDataField
	{
		/// <summary>
		/// Cell address, ie. "A1", "B3", etc
		/// </summary>
		public string CellAddress;
	}

}