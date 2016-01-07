using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface
{
	[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property)]
	public class FieldAttribute : System.Attribute
	{
		private string _address;

		public FieldAttribute(string cellAddress)
		{
			CellAddress = cellAddress;
		}

		public string CellAddress
		{
			get { return _address; }
			set { _address = value; }
		}

		public string FriendlyName;
	}
}
