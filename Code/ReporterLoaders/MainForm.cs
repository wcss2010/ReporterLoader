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
            PersonInfo bi = PersonInfo.GetPersonInfoObj(@"C:\Users\wcss\Desktop\ReporterLoader\Docs\xx2.doc");

            foreach (KeyValuePair<string, string> kvp in bi.BaseInfoDict)
            {
                System.Console.WriteLine(kvp.Key + "," + kvp.Value);
            }

            System.Console.WriteLine("-------------------------");

            foreach (SchoolInfo sli in bi.SchoolInfoList)
            {
                System.Console.WriteLine(sli.StartDate + "," + sli.EndDate + "," + sli.SchoolName + "," + sli.Subject + "," + sli.License);
            }

            System.Console.WriteLine("-------------------------");

            foreach (ResumeInfo sli in bi.ResumeInfoList)
            {
                System.Console.WriteLine(sli.StartDate + "," + sli.EndDate + "," + sli.WorkUnitAndJob);
            }

            System.Console.WriteLine("-------------------------");

            foreach (ProjectInfo sli in bi.ProjectInfoList)
            {
                System.Console.WriteLine(sli.Date + "," + sli.Name + "," + sli.Source + "," + sli.Job);
            }

            System.Console.WriteLine("-------------------------");

            foreach (PartTimeInfo sli in bi.PartTimeInfoList)
            {
                System.Console.WriteLine(sli.StartDate + "," + sli.EndDate + "," + sli.PartTimeContent + "," + sli.Job);
            }

            System.Console.WriteLine("-------------------------");

            foreach (HonorInfo sli in bi.HonorInfoList)
            {
                System.Console.WriteLine(sli.Date + "," + sli.Name + "," + sli.Level + "," + sli.Order);
            }

            System.Console.WriteLine("-------------------------");

            foreach (ProductionInfo sli in bi.ProductionInfoList)
            {
                System.Console.WriteLine(sli.Date + "," + sli.Name + "," + sli.PrinterAndLicenseNo + "," + sli.Order);
            }

            pictureBox1.Image = bi.HeadImage;
        }
    }
}