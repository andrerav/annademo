using System;
using System.Collections.Generic;

namespace AnNa.SpreadsheetParser.Interface.Sheets
{
	public interface ISheetRow
	{
		int RowIndex { get; set; }
		ErrorContainer ErrorContainer { get; }
	}

	public interface ISheetFields { }

	[Serializable]
	public class SheetRow : ISheetRow
	{
		public int RowIndex { get; set; }

		private ErrorContainer _errorContainer = new ErrorContainer();
		public ErrorContainer ErrorContainer
		{
			get { return _errorContainer; }
			internal set { _errorContainer = value; }
		}
	}

	[Serializable]
	public abstract class AbstractTypedSheet<R, F> : ITypedSheet<R, F> 
		where R : class, ISheetRow
		where F : class, ISheetFields
	{
		private List<string> _missingColumns = new List<string>();
		private ErrorContainer _errorContainer = new ErrorContainer();
		public ErrorContainer ErrorContainer
		{
			get { return _errorContainer; }
			set { _errorContainer = value; }
		}

		public List<string> MissingColumns => _missingColumns;

		public abstract string SheetName { get; }
		public virtual int MaximumNumberOfRows { get { return -1; } }
		public virtual int RowOffset { get { return 2; } }

		public F Fields { get; set; }
		public List<R> Rows { get; set; }

		public void AddToMissingColumns(string columnName)
		{
			if (_missingColumns.Contains(columnName))
				return;

			_missingColumns.Add(columnName);
		}
	}


	[Serializable]
	public class ErrorContainer
	{
		/// <summary>
		/// A list of parse errors
		/// </summary>
		public List<IParseError> Errors { get; private set; }

		public ErrorContainer() { Errors = new List<IParseError>(); }

		/// <summary>
		/// Add a syntax error to this container. Note that the DataField property must be set, otherwise an exception is thrown.
		/// </summary>
		/// <param name="error"></param>
		public void AddError(IParseError error)
		{
			if (error.DataField == null)
			{
				throw new ArgumentNullException(nameof(error.DataField));
			}

			if (Errors == null)
			{
				Errors = new List<IParseError>();
			}

			Errors.Add(error);
		}
	}

	public interface IParseError
	{
		/// <summary>
		/// The cell in which this syntax error was found
		/// </summary>
		string CellAddress { get; set; }

		/// <summary>
		/// The data field which contains erronous data
		/// </summary>
		SheetDataField DataField { get; set; }

	}

	[Serializable]
	public abstract class AbstractError : IParseError
	{
		public string CellAddress { get; set; }
		public SheetDataField DataField { get; set; }
	}

	[Serializable]
	public class RequiredFieldError : AbstractError
	{
	}

	[Serializable]
	public class SyntaxError : AbstractError
	{
		/// <summary>
		/// The original hinted type which could not be parsed
		/// </summary>
		public Type TypeHint { get; set; }

		/// <summary>
		/// The raw value from the spreadsheet
		/// </summary>
		public string RawValue { get; set; }
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

		/// <summary>
		/// Explicitly states that a column/field is optional (if true, no syntax error created for null values for valuetypes)
		/// </summary>
		public bool IsOptional { get; set; }

		/// <summary>
		/// True if this column is an index column or similar and can safely be ignored when testing the row for emptiness
		/// </summary>
		public bool Ignorable;

		/// <summary>
		/// Values that can be safely ignored when parsing the spreadsheet
		/// </summary>
		public string[] IgnoreableValues { get; set; }
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