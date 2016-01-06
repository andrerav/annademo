using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Attributes
{
	/// <summary>
	/// Used to deprecate spread sheet definitions 
	/// </summary>
	[AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
	public class DeprecatedAttribute : Attribute
	{
	}
}
