using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Attributes
{
	/// <summary>
	/// Used to signal that a field or property has been removed in a new version of a sheet definition
	/// </summary>
	[AttributeUsage(AttributeTargets.Field | AttributeTargets.Property, Inherited = false)]
	public class RemovedAttribute : Attribute
	{
	}
}
