using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using Aspose.Words;
using Aspose.Words.Tables;
using ReporterLoaders.DB.DocEntitys;

namespace ReporterLoaders
{
    public partial class MainForm : Form
    {
        protected List<PersonInfo> PersonInfoList = new List<PersonInfo>();

        public MainForm()
        {
            InitializeComponent();
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (ofdWords.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    PersonInfo bi = PersonInfo.GetPersonInfoObj(ofdWords.FileName);

                    if (bi.BaseInfoDict.Count == 0 || bi.SchoolInfoList.Count == 0)
                    {
                        return;
                    }

                    PersonInfoList.Add(bi);
                    UpdatePersonInfoList();
                }
                catch (Exception ex)
                {
                    MessageBox.Show("载入失败！Ex:" + ex.ToString());
                }
            }
        }

        protected void UpdatePersonInfoList()
        {
            dgvDetail.Rows.Clear();
            foreach (PersonInfo pi in PersonInfoList)
            {
                Image curImg = new Bitmap(110, 128);
                Graphics g = Graphics.FromImage(curImg);
                try
                {
                    g.DrawImage(pi.HeadImage, new Rectangle(0, 0, curImg.Width, curImg.Height));
                }
                finally
                {
                    g.Dispose();
                }

                List<object> cells = new List<object>();
                cells.Add(curImg);
                cells.Add(pi.BaseInfoDict["姓    名"]);
                cells.Add(pi.BaseInfoDict["性    别"]);
                cells.Add(pi.BaseInfoDict["民    族"]);
                cells.Add(pi.BaseInfoDict["政治面貌"]);
                cells.Add(pi.BaseInfoDict["本科专业"]);
                cells.Add(pi.BaseInfoDict["最高学位"]);
                cells.Add(pi.BaseInfoDict["所在省市"]);
                cells.Add(pi.BaseInfoDict["工作单位"]);
                cells.Add(pi.BaseInfoDict["所属部门"]);
                cells.Add(pi.BaseInfoDict["行政职务"]);
                cells.Add(pi.BaseInfoDict["涉密程度"]);
                cells.Add(pi.BaseInfoDict["单位电话"]);

                int rowIndex = dgvDetail.Rows.Add(cells.ToArray());
                dgvDetail.Rows[rowIndex].Tag = pi;
            }
        }

        private void dgvDetail_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.ColumnIndex == dgvDetail.Columns.Count - 1)
            {
                if (MessageBox.Show("真的要删除吗？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
                {
                    PersonInfoList.Remove((PersonInfo)dgvDetail.Rows[e.RowIndex].Tag);
                    UpdatePersonInfoList();
                }
            }
        }

        private void btnExportTo_Click(object sender, EventArgs e)
        {
            try
            {
                Workbook book = new Workbook(); //创建工作簿
                
                string desktopDir = Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory);
                string filepath = Path.Combine(desktopDir, "output.xlsx");
                foreach (PersonInfo pi in PersonInfoList)
                {
                    Worksheet sheet = book.Worksheets[book.Worksheets.Add(SheetType.Worksheet)];
                    sheet.Name = pi.BaseInfoDict["姓    名"];

                    //基本信息
                    sheet.Cells[0, 0].PutValue("基本信息：");
                    int rowIndex = 1;
                    foreach (KeyValuePair<string, string> kvp in pi.BaseInfoDict)
                    {
                        sheet.Cells[rowIndex, 0].PutValue(kvp.Key);
                        sheet.Cells[rowIndex, 1].PutValue(kvp.Value);
                        rowIndex++;
                    }

                    //受教育情况

                    //主要工作简历

                    //主要科研成绩

                    //兼职情况（技术或学术）

                    //科技获奖和荣誉情况（省部级以上）

                    //主要著作和专利情况

                    sheet.AutoFitColumns(); //自适应宽
                }
                                
                book.Save(filepath); //保存
                GC.Collect();

                MessageBox.Show(filepath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成excel出错：" + ex.Message);
            }
        }
    }
}