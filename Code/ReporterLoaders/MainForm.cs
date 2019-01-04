using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using Xceed.Words.NET;

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
            using (DocX document = DocX.Create(@"C:\Users\wcss\Desktop\ReporterLoader\Docs\XX1.docx"))
            {
                Table t = document.Tables[0];
                foreach (Row r in t.Rows)
                {
                    foreach (Cell c in r.Cells)
                    {
                        System.Console.WriteLine(c.ToString());
                    }
                }
            }
        }
    }
}
