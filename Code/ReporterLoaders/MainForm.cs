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
            Document doc = new Document(@"C:\Users\wcss\Desktop\ReporterLoader\Docs\XX2.doc");
            Table t = (Table)doc.GetChild(NodeType.Table, 0, true);
            PersonInfo bi = new PersonInfo();
            string sKey,sValue;
            sKey = string.Empty;
            sValue = string.Empty;
            foreach (Row r in t.Rows)
            {
                foreach (Cell c in r.Cells)
                {
                    string s = c.GetText().Trim().Replace("\a", string.Empty).Replace("\0", string.Empty).Replace("\n", string.Empty).Replace("\r", string.Empty);
                    if (string.IsNullOrEmpty(sKey))
                    {
                        if (string.IsNullOrEmpty(s))
                        {
                            continue;
                        }

                        sKey = s;
                    }
                    else if (string.IsNullOrEmpty(sValue))
                    {
                        sValue = s;

                        bi.BaseInfoDict.Add(sKey, sValue);
                        sKey = string.Empty;
                        sValue = string.Empty;
                    }
                }
            }

            foreach (KeyValuePair<string, string> kvp in bi.BaseInfoDict)
            {
                System.Console.WriteLine(kvp.Key + "," + kvp.Value);
            }
        }
    }
}