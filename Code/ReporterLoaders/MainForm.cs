using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.Text;
using System.Windows.Forms;
using Aspose.Cells;
using Aspose.Words;
using Aspose.Words.Tables;
using Noear.Weed;
using ReporterLoaders.DB.DocEntitys;
using ReporterLoaders.DB.MySql;
using ReporterLoaders.Forms;

namespace ReporterLoaders
{
    public partial class MainForm : Form
    {
        public static string DataDir = Path.Combine(Application.StartupPath, "Data");

        public static string FileDir = Path.Combine("C:\\", "HuiZongFiles");

        public static MainForm Instance { get; set; }

        protected List<PersonInfo> PersonInfoList = new List<PersonInfo>();

        private BackgroundWorker _worker = new BackgroundWorker();

        protected ProcessForm ProcessFormObj { get; set; }

        public MainForm()
        {
            InitializeComponent();

            Instance = this;
        }

        private void btnOpen_Click(object sender, EventArgs e)
        {
            if (ofdWords.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    foreach (string ss in ofdWords.FileNames)
                    {
                        PersonInfo bi = PersonInfo.GetPersonInfoObj(ss);

                        if (bi.BaseInfoDict.Count == 0 || bi.SchoolInfoList.Count == 0)
                        {
                            continue;
                        }

                        PersonInfoList.Add(bi);
                    }

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
                    sheet.Cells[0, 0].PutValue("基本信息");
                    int rowIndex = 1;
                    foreach (KeyValuePair<string, string> kvp in pi.BaseInfoDict)
                    {
                        sheet.Cells[rowIndex, 0].PutValue(kvp.Key);
                        sheet.Cells[rowIndex, 1].PutValue(kvp.Value);
                        rowIndex++;
                    }

                    //受教育情况
                    rowIndex++;
                    sheet.Cells[rowIndex, 0].PutValue("受教育情况");
                    rowIndex++;
                    foreach (SchoolInfo shi in pi.SchoolInfoList)
                    {
                        sheet.Cells[rowIndex, 0].PutValue(shi.StartDate);
                        sheet.Cells[rowIndex, 1].PutValue(shi.EndDate);
                        sheet.Cells[rowIndex, 2].PutValue(shi.SchoolName);
                        sheet.Cells[rowIndex, 3].PutValue(shi.Subject);
                        sheet.Cells[rowIndex, 4].PutValue(shi.License);

                        rowIndex++;
                    }

                    //主要工作简历
                    rowIndex++;
                    sheet.Cells[rowIndex, 0].PutValue("主要工作简历");
                    rowIndex++;
                    foreach (ResumeInfo rei in pi.ResumeInfoList)
                    {
                        sheet.Cells[rowIndex, 0].PutValue(rei.StartDate);
                        sheet.Cells[rowIndex, 1].PutValue(rei.EndDate);
                        sheet.Cells[rowIndex, 2].PutValue(rei.WorkUnitAndJob);

                        rowIndex++;
                    }

                    //主要科研成绩
                    rowIndex++;
                    sheet.Cells[rowIndex, 0].PutValue("主要科研成绩");
                    rowIndex++;
                    foreach (ProjectInfo pti in pi.ProjectInfoList)
                    {
                        sheet.Cells[rowIndex, 0].PutValue(pti.Date);
                        sheet.Cells[rowIndex, 1].PutValue(pti.Name);
                        sheet.Cells[rowIndex, 2].PutValue(pti.Source);
                        sheet.Cells[rowIndex, 3].PutValue(pti.Job);

                        rowIndex++;
                    }

                    //兼职情况（技术或学术）
                    rowIndex++;
                    sheet.Cells[rowIndex, 0].PutValue("兼职情况（技术或学术）");
                    rowIndex++;
                    foreach (PartTimeInfo prti in pi.PartTimeInfoList)
                    {
                        sheet.Cells[rowIndex, 0].PutValue(prti.StartDate);
                        sheet.Cells[rowIndex, 1].PutValue(prti.EndDate);
                        sheet.Cells[rowIndex, 2].PutValue(prti.PartTimeContent);
                        sheet.Cells[rowIndex, 3].PutValue(prti.Job);

                        rowIndex++;
                    }

                    //科技获奖和荣誉情况（省部级以上）
                    rowIndex++;
                    sheet.Cells[rowIndex, 0].PutValue("科技获奖和荣誉情况（省部级以上）");
                    rowIndex++;
                    foreach (HonorInfo hio in pi.HonorInfoList)
                    {
                        sheet.Cells[rowIndex, 0].PutValue(hio.Date);
                        sheet.Cells[rowIndex, 1].PutValue(hio.Name);
                        sheet.Cells[rowIndex, 2].PutValue(hio.Level);
                        sheet.Cells[rowIndex, 3].PutValue(hio.Order);

                        rowIndex++;
                    }

                    //主要著作和专利情况
                    rowIndex++;
                    sheet.Cells[rowIndex, 0].PutValue("主要著作和专利情况");
                    rowIndex++;
                    foreach (ProductionInfo pui in pi.ProductionInfoList)
                    {
                        sheet.Cells[rowIndex, 0].PutValue(pui.Date);
                        sheet.Cells[rowIndex, 1].PutValue(pui.Name);
                        sheet.Cells[rowIndex, 2].PutValue(pui.PrinterAndLicenseNo);
                        sheet.Cells[rowIndex, 3].PutValue(pui.Order);

                        rowIndex++;
                    }

                    sheet.AutoFitColumns(); //自适应宽
                }

                book.Worksheets.RemoveAt(0);
                book.Save(filepath); //保存
                GC.Collect();

                MessageBox.Show(filepath);
                Process.Start(filepath);
            }
            catch (Exception ex)
            {
                MessageBox.Show("生成excel出错：" + ex.Message);
            }
        }

        private void btnUploadData_Click(object sender, EventArgs e)
        {
            if (_worker.IsBusy)
            {
                return;
            }

            if (PersonInfoList.Count == 0)
            {
                return;
            }

            if (MessageBox.Show("真的要上传吗？", "提示", MessageBoxButtons.YesNo) == DialogResult.Yes)
            {
                ProcessFormObj = new ProcessForm();
                ProcessFormObj.Show();
                ProcessFormObj.BringToFront();

                btnUploadData.Enabled = false;
                _worker.RunWorkerAsync();
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void btnConfig_Click(object sender, EventArgs e)
        {
            ConfigForm cf = new ConfigForm();
            if (cf.ShowDialog() == DialogResult.OK)
            {
                Process.Start(Application.ExecutablePath);
                Application.Exit();
            }
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            string localDirs = ConfigHelper.GetAppConfig("LoclDirs");
            if (string.IsNullOrEmpty(localDirs))
            {
                FileDir = Path.Combine("C:\\", "HuiZongFiles");
            }
            else
            {
                FileDir = localDirs;
            }

            try
            {
                Directory.CreateDirectory(DataDir);
            }
            catch (Exception ex) { }

            try
            {
                Directory.CreateDirectory(FileDir);
            }
            catch (Exception ex) { }

            _worker.WorkerSupportsCancellation = true;
            _worker.DoWork += _worker_DoWork;
        }

        private void _worker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                if (IsHandleCreated)
                {
                    Invoke(new MethodInvoker(delegate()
                    {
                        if (ProcessFormObj != null)
                        {
                            ProcessFormObj.MaxProgressValue = PersonInfoList.Count;
                        }
                    }));
                }

                int progress = 0;
                foreach (PersonInfo pi in PersonInfoList)
                {
                    progress++;
                    if (IsHandleCreated)
                    {
                        Invoke(new MethodInvoker(delegate ()
                        {
                            if (ProcessFormObj != null)
                            {
                                ProcessFormObj.ProgressValue = progress;
                            }
                        }));
                    }

                    //窗体关闭则取消
                    if (ProcessFormObj == null || ProcessFormObj.IsDisposed)
                    {
                        break;
                    }

                    DataItem updateDataObj = null;
                    DbContext mysqlDBContext = ConnManager.Context;
                    string personId = Guid.NewGuid().ToString();
                    string unitId = Guid.NewGuid().ToString();

                    string workUnit = pi.BaseInfoDict["工作单位"];
                    if (string.IsNullOrEmpty(workUnit))
                    {
                        continue;
                    }
                    string personName = pi.BaseInfoDict["姓    名"];
                    if (string.IsNullOrEmpty(personName))
                    {
                        continue;
                    }
                    string mobilePhone = pi.BaseInfoDict["手    机"];
                    if (string.IsNullOrEmpty(mobilePhone))
                    {
                        continue;
                    }

                    string picFileName = personName + "_" + mobilePhone + ".png";

                    ////单位与人员信息
                    string curUnitId = mysqlDBContext.table("d_unit").where("name='" + workUnit + "'").select("id").getValue<string>(string.Empty);
                    if (string.IsNullOrEmpty(curUnitId))
                    {
                        //单位信息
                        updateDataObj = new DataItem();
                        updateDataObj.set("id", unitId);
                        updateDataObj.set("name", workUnit);
                        updateDataObj.set("sealname", workUnit);
                        updateDataObj.set("nickname", workUnit);
                        updateDataObj.set("address", pi.BaseInfoDict["通信地址"]);
                        updateDataObj.set("linkman", string.Empty);
                        updateDataObj.set("linknum", string.Empty);
                        updateDataObj.set("secgrade", string.Empty);
                        mysqlDBContext.table("d_unit").insert(updateDataObj);
                    }
                    else
                    {
                        unitId = curUnitId;

                        updateDataObj = new DataItem();
                        updateDataObj.set("id", unitId);
                        updateDataObj.set("name", workUnit);
                        updateDataObj.set("sealname", workUnit);
                        updateDataObj.set("nickname", workUnit);
                        updateDataObj.set("address", pi.BaseInfoDict["通信地址"]);
                        updateDataObj.set("linkman", string.Empty);
                        updateDataObj.set("linknum", string.Empty);
                        updateDataObj.set("secgrade", string.Empty);
                        mysqlDBContext.table("d_unit").where("id='" + unitId + "'").update(updateDataObj);
                    }
                    string curPersonId = mysqlDBContext.table("d_person").where("name='" + personName + "' and mobilephone='" + mobilePhone + "'").select("id").getValue<string>(string.Empty);
                    if (string.IsNullOrEmpty(curPersonId))
                    {
                        DateTime birthdayDate = DateTime.Now;
                        try
                        {
                            string[] birth = pi.BaseInfoDict["出生年月"].Split(new string[] { "." }, StringSplitOptions.None);
                            birthdayDate = new DateTime(int.Parse(birth[0]), int.Parse(birth[1]), 1);
                        }
                        catch (Exception ex) { }

                        updateDataObj = new DataItem();
                        updateDataObj.set("id", personId);
                        updateDataObj.set("name", pi.BaseInfoDict["姓    名"]);
                        updateDataObj.set("post", pi.BaseInfoDict["行政职务"].Trim() + "/" + pi.BaseInfoDict["技术职称"].Trim());
                        updateDataObj.set("specialty", pi.BaseInfoDict["现从事专业"]);
                        updateDataObj.set("gender", pi.BaseInfoDict["性    别"]);
                        updateDataObj.set("birthday", birthdayDate);
                        updateDataObj.set("phone", pi.BaseInfoDict["单位电话"]);
                        updateDataObj.set("mobilephone", pi.BaseInfoDict["手    机"]);
                        updateDataObj.set("address", pi.BaseInfoDict["通信地址"]);
                        updateDataObj.set("unitid", unitId);
                        mysqlDBContext.table("d_person").insert(updateDataObj);
                    }
                    else
                    {
                        personId = curPersonId;

                        DateTime birthdayDate = DateTime.Now;
                        try
                        {
                            string[] birth = pi.BaseInfoDict["出生年月"].Split(new string[] { "." }, StringSplitOptions.None);
                            birthdayDate = new DateTime(int.Parse(birth[0]), int.Parse(birth[1]), 1);
                        }
                        catch (Exception ex) { }

                        updateDataObj = new DataItem();
                        updateDataObj.set("id", personId);
                        updateDataObj.set("name", pi.BaseInfoDict["姓    名"]);
                        updateDataObj.set("post", pi.BaseInfoDict["行政职务"].Trim() + "/" + pi.BaseInfoDict["技术职称"].Trim());
                        updateDataObj.set("specialty", pi.BaseInfoDict["现从事专业"]);
                        updateDataObj.set("gender", pi.BaseInfoDict["性    别"]);
                        updateDataObj.set("birthday", birthdayDate);
                        updateDataObj.set("phone", pi.BaseInfoDict["单位电话"]);
                        updateDataObj.set("mobilephone", pi.BaseInfoDict["手    机"]);
                        updateDataObj.set("address", pi.BaseInfoDict["通信地址"]);
                        updateDataObj.set("unitid", unitId);
                        mysqlDBContext.table("d_person").where("id='" + personId + "'").update(updateDataObj);
                    }

                    //清理数据
                    mysqlDBContext.table("d_personex").where("id = '" + personId + "'").delete();
                    mysqlDBContext.table("d_person_jianli").where("personid = '" + personId + "'").delete();
                    mysqlDBContext.table("d_person_jianzhiqingkuang").where("personid = '" + personId + "'").delete();
                    mysqlDBContext.table("d_person_keyanchengji").where("personid = '" + personId + "'").delete();
                    mysqlDBContext.table("d_person_rongyuqingkuang").where("personid = '" + personId + "'").delete();
                    mysqlDBContext.table("d_person_schooldetail").where("personid = '" + personId + "'").delete();
                    mysqlDBContext.table("d_person_zhuanliqingkuang").where("personid = '" + personId + "'").delete();

                    updateDataObj = new DataItem();
                    updateDataObj.set("id", personId);
                    updateDataObj.set("benkezhuanye", pi.BaseInfoDict["本科专业"]);
                    updateDataObj.set("chuanzhen", pi.BaseInfoDict["传    真"]);
                    updateDataObj.set("email", pi.BaseInfoDict["电子邮箱"]);
                    updateDataObj.set("minzu", pi.BaseInfoDict["民    族"]);
                    updateDataObj.set("shemichengdu", pi.BaseInfoDict["涉密程度"]);
                    updateDataObj.set("shifoudingjuguowai", pi.BaseInfoDict["是否具有外国国籍或港澳台居民身份，是否拥有境外永久居留权（绿卡），如有，请申明"]);
                    updateDataObj.set("suoshubumen", pi.BaseInfoDict["所属部门"]);
                    updateDataObj.set("suozaichengshi", pi.BaseInfoDict["所在省市"]);
                    updateDataObj.set("youbian", pi.BaseInfoDict["邮政编码"]);
                    updateDataObj.set("zhengzhimianmao", pi.BaseInfoDict["政治面貌"]);
                    updateDataObj.set("zuigaoxuewei", pi.BaseInfoDict["最高学位"]);
                    updateDataObj.set("zhaopian", picFileName);

                    mysqlDBContext.table("d_personex").insert(updateDataObj);

                    ////受教育情况
                    foreach (SchoolInfo sii in pi.SchoolInfoList)
                    {
                        updateDataObj = new DataItem();
                        updateDataObj.set("id", Guid.NewGuid().ToString());
                        updateDataObj.set("personid", personId);
                        updateDataObj.set("starttime", sii.StartDate);
                        updateDataObj.set("endtime", sii.EndDate);
                        updateDataObj.set("school", sii.SchoolName);
                        updateDataObj.set("kemu", sii.Subject);
                        updateDataObj.set("xueli", sii.License);

                        mysqlDBContext.table("d_person_schooldetail").insert(updateDataObj);
                    }

                    ////主要工作简历
                    foreach (ResumeInfo rii in pi.ResumeInfoList)
                    {
                        updateDataObj = new DataItem();
                        updateDataObj.set("id", Guid.NewGuid().ToString());
                        updateDataObj.set("personid", personId);
                        updateDataObj.set("starttime", rii.StartDate);
                        updateDataObj.set("endtime", rii.EndDate);
                        updateDataObj.set("unitandjob", rii.WorkUnitAndJob);

                        mysqlDBContext.table("d_person_jianli").insert(updateDataObj);
                    }

                    ////主要科研成绩
                    foreach (ProjectInfo pii in pi.ProjectInfoList)
                    {
                        updateDataObj = new DataItem();
                        updateDataObj.set("id", Guid.NewGuid().ToString());
                        updateDataObj.set("personid", personId);
                        updateDataObj.set("time", pii.Date);
                        updateDataObj.set("name", pii.Name);
                        updateDataObj.set("source", pii.Source);
                        updateDataObj.set("job", pii.Job);

                        mysqlDBContext.table("d_person_keyanchengji").insert(updateDataObj);
                    }

                    ////兼职情况（技术或学术）
                    foreach (PartTimeInfo pti in pi.PartTimeInfoList)
                    {
                        updateDataObj = new DataItem();
                        updateDataObj.set("id", Guid.NewGuid().ToString());
                        updateDataObj.set("personid", personId);
                        updateDataObj.set("starttime", pti.StartDate);
                        updateDataObj.set("endtime", pti.EndDate);
                        updateDataObj.set("detail", pti.PartTimeContent);
                        updateDataObj.set("job", pti.Job);

                        mysqlDBContext.table("d_person_jianzhiqingkuang").insert(updateDataObj);
                    }

                    ////科技获奖和荣誉情况（省部级以上）
                    foreach (HonorInfo hii in pi.HonorInfoList)
                    {
                        updateDataObj = new DataItem();
                        updateDataObj.set("id", Guid.NewGuid().ToString());
                        updateDataObj.set("personid", personId);
                        updateDataObj.set("time", hii.Date);
                        updateDataObj.set("name", hii.Name);
                        updateDataObj.set("level", hii.Level);
                        updateDataObj.set("`order`", hii.Order);

                        mysqlDBContext.table("d_person_rongyuqingkuang").insert(updateDataObj);
                    }

                    ////主要著作和专利情况
                    foreach (ProductionInfo pni in pi.ProductionInfoList)
                    {
                        updateDataObj = new DataItem();
                        updateDataObj.set("id", Guid.NewGuid().ToString());
                        updateDataObj.set("personid", personId);
                        updateDataObj.set("time", pni.Date);
                        updateDataObj.set("name", pni.Name);
                        updateDataObj.set("printerandprivateno", pni.PrinterAndLicenseNo);
                        updateDataObj.set("`order`", pni.Order);

                        mysqlDBContext.table("d_person_zhuanliqingkuang").insert(updateDataObj);
                    }

                    //保存头像
                    if (pi.HeadImage != null)
                    {
                        MemoryStream mstream = new MemoryStream();
                        pi.HeadImage.Save(mstream, ImageFormat.Png);
                        byte[] ImgByte = mstream.ToArray();
                        mstream.Close();

                        File.WriteAllBytes(Path.Combine(FileDir, picFileName), ImgByte);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("上传出错！Ex:" + ex.ToString());
            }
            finally
            {
                if (IsHandleCreated)
                {
                    Invoke(new MethodInvoker(delegate ()
                    {
                        btnUploadData.Enabled = true;

                        if (ProcessFormObj != null)
                        {
                            ProcessFormObj.Close();
                        }
                    }));
                }
            }
        }
    }
}