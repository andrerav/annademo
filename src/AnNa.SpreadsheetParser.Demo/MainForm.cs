﻿using System;
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
				parser.OpenFile(openFileDialog.FileName);

				var type = typeof (ISheetSpecification);
				var types = AppDomain.CurrentDomain.GetAssemblies()
					.SelectMany(s => s.GetTypes())
					.Where(p => !p.IsAbstract
								&& type.IsAssignableFrom(p)
					            && p != typeof (ISheetSpecification)).ToList();

				var everything = new Dictionary<string, List<Dictionary<string, string>>>();
				foreach (var t in types)
				{
					var instance = Activator.CreateInstance(t);
					var contents = parser.GetSheetContents(instance as ISheetSpecification);
					everything[t.Name] = contents;
				}

				var settings = new JsonSerializerSettings();
				settings.Formatting = Formatting.Indented;

				txtSerializedContents.Text = JsonConvert.SerializeObject(everything);
			}
		}

	}
}
