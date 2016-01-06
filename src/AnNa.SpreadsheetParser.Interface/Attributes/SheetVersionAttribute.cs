using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Attributes
{
	[AttributeUsage(AttributeTargets.Class)]
	public class SheetVersionAttribute : Attribute
	{
		public Version Version { get; set; }

		public SheetVersionAttribute(Version version)
		{
			Version = version;
		}
	}
}
