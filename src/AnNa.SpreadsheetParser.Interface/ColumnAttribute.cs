using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface
{
	[AttributeUsage(AttributeTargets.Field|AttributeTargets.Property, AllowMultiple = false, Inherited = true)]
	public class ColumnAttribute : System.Attribute
	{
		private string _column;
		private string _friendlyName;
		private bool _ignorable;
		private string[] _ignorableValues;

		public ColumnAttribute(string column)
		{
			Column = column;
		}

		public string Column
		{
			get { return _column; }
			set { _column = value; }
		}

		public bool Ignorable
		{
			get
			{
				return _ignorable;
			}

			set
			{
				_ignorable = value;
			}
		}

		public string[] IgnorableValues
		{
			get
			{
				return _ignorableValues;
			}

			set
			{
				_ignorableValues = value;
			}
		}

		public string FriendlyName
		{
			get
			{
				return _friendlyName;
			}

			set
			{
				_friendlyName = value;
			}
		}
	}
}
