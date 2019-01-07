using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ReporterLoaders
{
    public partial class MainForm : Form
    {
        public MainForm()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            Document doc = new Document(@"C:\Users\wcss\Desktop\ReporterLoader\Docs\XX2.doc");
            Table t = (Table)doc.GetChild(NodeType.Table, 0, true);
            foreach (Row r in t.Rows)
            {
                foreach (Cell c in r.Cells)
                {
                    System.Console.WriteLine(c.GetText());
                }
            }

            
        }
    }
}