using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Attributes
{
	public abstract class SheetAttribute : Attribute
	{
		private string _friendlyName;
		public bool IsOptional { get; set; }
		public string FriendlyName => _friendlyName;

		public SheetAttribute(string friendlyName)
		{
			_friendlyName = friendlyName;
		}
	}

	[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
	public class ColumnAttribute : SheetAttribute
	{
		private string _column;
		public string Column => _column;

		public bool Ignorable { get; set; }
		public string[] IgnorableValues { get; set; }

		public ColumnAttribute(string column) : this(column, column) { }
		
		public ColumnAttribute(string column, string friendlyName) : base(friendlyName)
		{
			_column = column;
		}

	}

	[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
	public class FieldAttribute : SheetAttribute
	{
		private string _cellAddress;
		public string CellAddress => _cellAddress;

		public FieldAttribute (string cellAddress) : this(cellAddress, cellAddress) { }
		public FieldAttribute(string cellAddress, string friendlyName) : base(friendlyName)
		{
			_cellAddress = cellAddress;
		}
	}


}
