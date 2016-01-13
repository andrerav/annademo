using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using AnNa.SpreadsheetParser.Interface;
using AnNa.SpreadsheetParser.Interface.Sheets;
using AnNa.SpreadsheetParser.SpreadsheetGear;
using AnNa.SpreadSheetParser.EPPlus;
using Newtonsoft.Json;
using System.Reflection;
using AnNa.SpreadsheetParser.Interface.Sheets.Typed;

namespace AnNaSpreadSheetDemo
{
	public partial class MainForm : Form
	{
		public MainForm()
		{
			InitializeComponent();
		}

		private void openToolStripMenuItem_Click(object sender, EventArgs e)
		{
			OpenFile(new AnNaSpreadSheetParserSpreadsheetGear());
		}

		private void openUsingEPPlusToolStripMenuItem_Click(object sender, EventArgs e)
		{
			OpenFile(new AnNaSpreadSheetParserEPPlus());

		}

		private void OpenFile(IAnNaSpreadSheetParser10 parser)
		{
			var result = openFileDialog.ShowDialog();
			if (!File.Exists(openFileDialog.FileName))
			{
				MessageBox.Show(@"File not found", @"File not found", MessageBoxButtons.OK, MessageBoxIcon.Error);
			}
			else
			{
				object everything1 = GetContents(parser);
				var settings = new JsonSerializerSettings();
				settings.Formatting = Formatting.Indented;
				txtSerializedContents.Text = JsonConvert.SerializeObject(everything1);
			}
		}

		private Dictionary<string, object> GetContents(IAnNaSpreadSheetParser10 parser)
		{
			parser.OpenFile(openFileDialog.FileName);

			var type = typeof(ISheet);
			var types = AppDomain.CurrentDomain.GetAssemblies()
				.SelectMany(s => s.GetTypes())
				.Where(p => !p.IsAbstract
							&& type.IsAssignableFrom(p)
							&& p != typeof(ISheet)).ToList();

			var everything = new Dictionary<string, object>();
			foreach (var t in types)
			{

				bool isTyped = t.GetInterfaces().Any(x =>
				  x.IsGenericType &&
				  x.GetGenericTypeDefinition() == typeof(ITypedSheet<,>));

				var instance = Activator.CreateInstance(t);
				object contents = null;
				if (isTyped)
				{
					MethodInfo method = typeof(IAnNaSpreadSheetParser10).GetMethods().First(m => m.Name == nameof(parser.GetSheetBulkData) && m.GetParameters().Count() == 1);
					MethodInfo generic = method.MakeGenericMethod(
													t.GetInterfaces()
													.Where(i => i.IsGenericType 
														&& i.GetGenericTypeDefinition() == typeof(ITypedSheet<,>))
													.Single().GetGenericArguments()
													.ToArray());

					contents = generic.Invoke(parser, new object[] { instance });
				}
				else
				{
					contents = parser.GetSheetBulkData(instance as ISheetWithBulkData);
				}

				if (contents != null)
				{
					everything[t.Name] = contents;
				}
			}

			return everything;
		}

		public static IEnumerable<Type> GetAllTypesImplementingOpenGenericType(Type openGenericType, Assembly assembly)
		{
			return from x in Assembly.GetAssembly(typeof(Program)).GetTypes()
				   from z in x.GetInterfaces()
				   let y = x.BaseType
				   where
				   (y != null && y.IsGenericType &&
				   openGenericType.IsAssignableFrom(y.GetGenericTypeDefinition())) ||
				   (z.IsGenericType &&
				   openGenericType.IsAssignableFrom(z.GetGenericTypeDefinition()))
				   select x;
		}
	}
}
