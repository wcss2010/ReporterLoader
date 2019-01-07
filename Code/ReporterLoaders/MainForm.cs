using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Aspose.Words;
using Aspose.Words.Tables;
using ReporterLoaders.DB.DocEntitys;

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
            PersonInfo bi = PersonInfo.GetPersonInfoObj(@"C:\Users\wcss\Desktop\ReporterLoader\Docs\xx3.doc");

            foreach (KeyValuePair<string, string> kvp in bi.BaseInfoDict)
            {
                System.Console.WriteLine(kvp.Key + "," + kvp.Value);
            }

            foreach (SchoolInfo sli in bi.SchoolInfoList)
            {
                System.Console.WriteLine(sli.StartDate + "," + sli.EndDate + "," + sli.SchoolName + "," + sli.Subject + "," + sli.License);
            }
        }
    }
}