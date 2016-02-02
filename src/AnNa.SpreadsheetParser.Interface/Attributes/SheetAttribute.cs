using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Attributes
{
	public abstract class SheetAttribute : Attribute
	{
		private string _friendlyName;
		
		public string FriendlyName => _friendlyName;

		public bool IsOptional { get; set; }

		public bool SkipOnRead { get; set; }
		public bool SkipOnWrite { get; set; }

		public string[] ValuesSkippedOnRead { get; set; }

		public SheetAttribute(string friendlyName)
		{
			_friendlyName = friendlyName;
		}
	}
	/// <summary>
	/// The parser treats fields and properties in sheet definitions marked with the ColumnAttribute as a strongly typed column definition in that sheet
	/// </summary>
	[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
	public class ColumnAttribute : SheetAttribute
	{
		private string _column;
		public string Column => _column;
		
		public ColumnAttribute(string column) : this(column, column) { }
		
		public ColumnAttribute(string column, string friendlyName) : base(friendlyName)
		{
			_column = column;
		}
	}

	/// <summary>
	/// The parser treats fields and properties in sheet definitions marked with the FieldAttribute as the strongly typed definition of a single field/cell in that sheet
	/// </summary>
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

	/// <summary>
	/// Use the ListMappingAttribute to map a set of cells in a sheet to a collection
	/// </summary>
	[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
	public class ListMappingAttribute : SheetAttribute
	{
		private List<string> _cellAddresses;
		public List<string> CellAddresses => _cellAddresses;

		public ListMappingAttribute(string friendlyName, params string[] cellAddresses) : base(friendlyName)
		{
			_cellAddresses = cellAddresses.ToList();
		}
	}


}
