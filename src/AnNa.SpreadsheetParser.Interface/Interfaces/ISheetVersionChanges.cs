using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace AnNa.SpreadsheetParser.Interface.Interfaces
{

	public enum VersionChangeType
	{
		Added,			//new features
		Changed,		//change in existing functionality
		Deprecated,		//once-stable features removed in upcoming release
		Removed,		//deprecated features removed in this release
		Fixed,			//bug fixes
		Security		//vulnerability fixes
	}


	public class SheetVersionChange
	{
		public string Message { get; set; }
		public VersionChangeType ChangeType { get; set; }
	}

	public interface ISheetVersionChanges
	{
		IEnumerable<SheetVersionChange> SheetVersionChanges { get; set; }
	}
}
