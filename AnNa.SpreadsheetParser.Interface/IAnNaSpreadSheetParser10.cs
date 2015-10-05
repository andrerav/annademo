using System;
using System.Collections.Generic;

namespace AnNaSpreadSheetParser
{
	public interface IAnNaSpreadSheetParser10
	{
		void OpenFile(string path, string password = null);
		List<Dictionary<string, string>> GetSheetContents(ISheetSpecification sheetSpecification);
	}

	public class InvalidColumnPositionException : Exception
	{
		public InvalidColumnPositionException(string message)
			: base(message)
		{
		}
	}

	public class ColumnNotFoundException : Exception
	{
		public ColumnNotFoundException(string message)
			: base(message)
		{
		}
	}
}