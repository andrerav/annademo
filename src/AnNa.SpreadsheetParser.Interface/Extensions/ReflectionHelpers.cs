﻿using AnNa.SpreadsheetParser.Interface.Attributes;
using AnNa.SpreadsheetParser.Interface.Sheets;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AnNa.SpreadsheetParser.Interface.Extensions
{

	public static class ReflectionHelpers
	{
		public class SheetDefinitionMetaData
		{
			public Type Type { get; set; }
			public Type[] TypeParameters { get; set; }
			public string GroupingKey { get; set; }
			public Version Version { get; set; }
			public string Authority { get; set; }
		}


		/// <summary>
		/// Returns a list of sheet definitions grouped by sheet definition grouping key, ordered by match on authority, then descending by defintion version
		/// </summary>
		/// <param name="authority"></param>
		/// <returns></returns>
		public static List<IGrouping<string, SheetDefinitionMetaData>> GetSheetDefinitionsOrderedByAuthority(string authority)
		{

			return AppDomain.CurrentDomain.GetAssemblies()
				.SelectMany(s => s.GetTypes())
				.Where(p => !p.IsAbstract
							&& p.BaseType != null
							&& p.BaseType.IsGenericType
							&& p.BaseType.GetGenericTypeDefinition() == typeof(AbstractTypedSheet<,>)
							&& p.CustomAttributes.Any(ca => ca.AttributeType == typeof(SheetVersionAttribute)))
				.Select(t =>
				{
					var versionParams = t.CustomAttributes.Single(ca => ca.AttributeType == typeof(SheetVersionAttribute)).ConstructorArguments.Select(ca => ca.Value);
					var genericTypes = t.GetInterfaces()
							.Where(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(ITypedSheet<,>))
							.Single()
							.GetGenericArguments();
					return new SheetDefinitionMetaData
					{
						Type = t,
						TypeParameters = genericTypes,
						GroupingKey = versionParams.ElementAt(0).ToString(),
						Version = new Version((int)versionParams.ElementAt(1), (int)versionParams.ElementAt(2)),
						Authority = versionParams.ElementAt(3).ToString()
					};
				})
				.OrderByDescending(t => t.Authority == authority) //Favor authority-specific sheet definitions
				.ThenByDescending(t => t.Version)
				.GroupBy(g => g.GroupingKey)
				.ToList();
		}

	}
}
