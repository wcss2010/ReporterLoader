using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace ReporterLoaders.Forms
{
    public partial class ProcessForm : Form
    {
        public string HintText
        {
            get { return lblText.Text; }
            set { lblText.Text = value; }
        }

        public int ProgressValue
        {
            get { return pbrUploads.Value; }
            set { pbrUploads.Value = value; }
        }
        
        public ProcessForm()
        {
            InitializeComponent();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}