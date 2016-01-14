using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Attributes
{
	[AttributeUsage(AttributeTargets.Class)]
	public class SheetVersionAttribute : Attribute
	{
		private Version _version;
		private string _groupingKey;
		private string _authority;
		public Version Version => _version;
		public string GroupingKey => _groupingKey;
		public string Authority => _authority;

		/// <summary>
		/// Specifies the version of a sheet definition (version: <paramref name="major"/>.<paramref name="minor"/>). 
		/// <paramref name="groupingKey"/> is used to create groups of sheet definitions (e.g., Waste, Dpg, Security etc.). 
		/// <paramref name="authority"/> specifies the sheet authority (e.g., AnNa, SSN etc.) - leave blank for AnNa!
		/// </summary>
		/// <param name="groupingKey"></param>
		/// <param name="major"></param>
		/// <param name="minor"></param>
		/// <param name="authority"></param>
		public SheetVersionAttribute(string groupingKey, int major, int minor, string authority)
		{
			_version = new Version(major, minor);
			_groupingKey = groupingKey;
			_authority = authority;
		}
	}

}
