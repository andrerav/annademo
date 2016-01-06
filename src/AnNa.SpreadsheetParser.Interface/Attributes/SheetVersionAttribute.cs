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
		public Version Version => _version;

		public SheetVersionAttribute(int major, int minor)
		{
			_version = new Version(major, minor);
		}
	}
}
