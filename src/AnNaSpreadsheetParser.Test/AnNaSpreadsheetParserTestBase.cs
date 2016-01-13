using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using AnNa.SpreadsheetParser.Interface;
using AnNa.SpreadsheetParser.Interface.Sheets;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using System.Reflection;
using AnNa.SpreadsheetParser.Interface.Attributes;
using AnNa.SpreadsheetParser.Interface.Sheets.Typed;
using System.Text.RegularExpressions;

namespace AnNaSpreadSheetParserTest
{
	public abstract class AnNaSpreadsheetParserTestBase
	{
		protected abstract Version Version { get; }

		protected IAnNaSpreadSheetParser10 parser;

		/// <summary>
		/// Creates an instance of the parser of type T. 
		/// Initializes the parser with a sheet from the TestSheets directory matching the Version property.
		/// </summary>
		/// <typeparam name="T"></typeparam>
		protected virtual void InitializeParser<T>() where T : class, IAnNaSpreadSheetParser10, new()
		{
			if (parser == null)
			{
				parser = new T();

				var workbookVersions = Directory.GetFiles("./../../TestSheets", "*.xlsx", SearchOption.TopDirectoryOnly)
					.Select(s=>
					{
						//Take the string between # and . in the file path then split it on -
						var versionParams = Regex.Match(s, @"#([^.]*)\.").Groups[1].Value.Split('-').Select(c => int.Parse(c)).ToList(); ;

						return new {
							Path = s,
							Version = new Version(versionParams[0], versionParams[1])
						};
					})
					.OrderByDescending(sv=> sv.Version);

				if (workbookVersions.All(sv => sv.Version != Version))
					throw new Exception($"No test sheet for version {Version.ToString()} found");

				foreach (var workbook in workbookVersions)
				{
					if (workbook.Version > Version)
						continue;

					parser.OpenFile(workbook.Path);
				}
			}
		}

		[TestCleanup]
		public void Cleanup()
		{
		}


		private void AssertSheetHasBulkData(string sheetGroup)
		{
			var type = typeof(AbstractTypedSheet<,>);
			var sheets = AppDomain.CurrentDomain.GetAssemblies()
				.SelectMany(s => s.GetTypes())
				.Where(p =>
				{
					if (p.IsAbstract || p.BaseType == null
							|| !p.BaseType.IsGenericType
							|| p.BaseType.GetGenericTypeDefinition() != type)
					{
						return false;
					}

					var attr = p.GetCustomAttribute(typeof(SheetVersionAttribute)) as SheetVersionAttribute;
					return attr != null && attr.GroupingKey == sheetGroup;
				}
				).Select(p => new
				{
					Type = p,
					TypeParameters = p.GetInterfaces()
							.Where(i => i.IsGenericType && i.GetGenericTypeDefinition() == typeof(ITypedSheet<,>))
							.Single()
							.GetGenericArguments(),
					Version = (p.GetCustomAttribute(typeof(SheetVersionAttribute)) as SheetVersionAttribute).Version
				}).OrderByDescending(p=> p.Version);


			var method = typeof(IAnNaSpreadSheetParser10).GetMethods().First(m => m.Name == nameof(IAnNaSpreadSheetParser10.GetSheetBulkData) && m.GetParameters().Count() == 1);
			foreach (var item in sheets)
			{
				if (item.Version > Version)
					continue;

				var instance = Activator.CreateInstance(item.Type);      
				var genericMethod = method.MakeGenericMethod(item.TypeParameters.ToArray());

				genericMethod.Invoke(parser, new object[] { instance });

				var rowProperty = item.Type.GetProperty(nameof(WasteSheet10.Rows), BindingFlags.Public | BindingFlags.Instance);

				Assert.IsTrue((rowProperty.GetValue(instance) as IList).Count > 0);
				return;
			}

			Assert.Fail();
		}


		[TestMethod]
		public virtual void ReadCruiseList()
		{
			AssertSheetHasBulkData(SheetGroups.Cruise);
		}

		[TestMethod]
		public virtual void ReadStowawayList()
		{
			AssertSheetHasBulkData(SheetGroups.Stowaway);
		}

		[TestMethod]
		public virtual void ReadCrewList()
		{
			AssertSheetHasBulkData(SheetGroups.CrewList);
		}

		[TestMethod]
		public virtual void ReadPaxList()
		{
			AssertSheetHasBulkData(SheetGroups.PaxList);
		}

		[TestMethod]
		public virtual void ReadWasteList()
		{
			AssertSheetHasBulkData(SheetGroups.Waste);
		}

		[TestMethod]
		public virtual void ReadLast10CallsList()
		{
			AssertSheetHasBulkData(SheetGroups.SecurityPortCalls);
		}

		[TestMethod]
		public virtual void ReadS2SList()
		{
			AssertSheetHasBulkData(SheetGroups.SecurityShipToShip);
		}

		[TestMethod]
		public virtual void ReadShipStoresList()
		{
			AssertSheetHasBulkData(SheetGroups.ShipStores);
		}

		[TestMethod]
		public virtual void ReadDPGList()
		{
			AssertSheetHasBulkData(SheetGroups.Dpg);
		}


	}
}