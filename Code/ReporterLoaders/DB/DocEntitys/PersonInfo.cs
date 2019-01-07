using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Tables;

namespace ReporterLoaders.DB.DocEntitys
{
    public class PersonInfo
    {
        /// <summary>
        /// 基础数据字典
        /// </summary>
        public Dictionary<string, string> BaseInfoDict = new Dictionary<string, string>();

        /// <summary>
        /// 教育情况列表
        /// </summary>
        public List<SchoolInfo> SchoolInfoList = new List<SchoolInfo>();

        /// <summary>
        /// 主要工作简历
        /// </summary>
        public List<ResumeInfo> ResumeInfoList = new List<ResumeInfo>();

        /// <summary>
        /// 主要科研成绩
        /// </summary>
        public List<ProjectInfo> ProjectInfoList = new List<ProjectInfo>();

        /// <summary>
        /// 兼职情况（技术或学术）
        /// </summary>
        public List<PartTimeInfo> PartTimeInfoList = new List<PartTimeInfo>();

        /// <summary>
        /// 科技获奖和荣誉情况（省部级以上）
        /// </summary>
        public List<HonorInfo> HonorInfoList = new List<HonorInfo>();

        public static PersonInfo GetPersonInfoObj(string wordFiles)
        {
            PersonInfo pi = new PersonInfo();

            if (File.Exists(wordFiles))
            {
                Document doc = new Document(wordFiles);

                #region 提取基础信息
                Table t = (Table)doc.GetChild(NodeType.Table, 0, true);
                string sKey, sValue;
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

                            //添加数据行
                            pi.BaseInfoDict.Add(sKey, sValue);
                            sKey = string.Empty;
                            sValue = string.Empty;
                        }
                    }
                }
                #endregion

                #region 受教育情况
                Table t2 = (Table)doc.GetChild(NodeType.Table, 1, true);
                foreach (Row r in t2.Rows)
                {
                    SchoolInfo sii = new SchoolInfo();
                    sii.StartDate = r.Cells[0].GetText();
                    sii.EndDate = r.Cells[1].GetText();
                    sii.SchoolName = r.Cells[2].GetText();
                    sii.Subject = r.Cells[3].GetText();
                    sii.License = r.Cells[4].GetText();

                    pi.SchoolInfoList.Add(sii);
                }
                #endregion

                #region 主要工作简历
                Table t3 = (Table)doc.GetChild(NodeType.Table, 2, true);
                foreach (Row r in t2.Rows)
                {
                    ResumeInfo wui = new ResumeInfo();
                    wui.StartDate = r.Cells[0].GetText();
                    wui.EndDate = r.Cells[1].GetText();
                    wui.WorkUnitAndJob = r.Cells[2].GetText();

                    pi.ResumeInfoList.Add(wui);
                }
                #endregion

                #region 主要科研成绩
                Table t4 = (Table)doc.GetChild(NodeType.Table, 3, true);
                foreach (Row r in t2.Rows)
                {
                    ProjectInfo ti = new ProjectInfo();
                    ti.Date = r.Cells[0].GetText();
                    ti.Name = r.Cells[1].GetText();
                    ti.Source = r.Cells[2].GetText();
                    ti.Job = r.Cells[3].GetText();

                    pi.ProjectInfoList.Add(ti);
                }
                #endregion

                #region 兼职情况
                Table t5 = (Table)doc.GetChild(NodeType.Table, 4, true);
                foreach (Row r in t2.Rows)
                {
                    PartTimeInfo pti = new PartTimeInfo();
                    pti.StartDate = r.Cells[0].GetText();
                    pti.EndDate = r.Cells[1].GetText();
                    pti.PartTimeContent = r.Cells[2].GetText();
                    pti.Job = r.Cells[3].GetText();

                    pi.PartTimeInfoList.Add(pti);
                }
                #endregion

                #region 科技获奖和荣誉情况
                Table t6 = (Table)doc.GetChild(NodeType.Table, 5, true);
                foreach (Row r in t2.Rows)
                {
                    HonorInfo hii = new HonorInfo();
                    hii.Date = r.Cells[0].GetText();
                    hii.Name = r.Cells[1].GetText();
                    hii.Level = r.Cells[2].GetText();
                    hii.Order = r.Cells[3].GetText();

                    pi.HonorInfoList.Add(hii);
                }
                #endregion

                #region 主要著作和专利情况
                Table t7 = (Table)doc.GetChild(NodeType.Table, 6, true);
                foreach (Row r in t2.Rows)
                {
                    foreach (Cell c in r.Cells)
                    {

                    }
                }
                #endregion
            }

            return pi;
        }
    }

    /// <summary>
    /// 受教育情况
    /// </summary>
    public class SchoolInfo
    {
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string SchoolName { get; set; }
        public string Subject { get; set; }
        public string License { get; set; }
    }

    /// <summary>
    /// 主要工作简历
    /// </summary>
    public class ResumeInfo
    {
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string WorkUnitAndJob { get; set; }
    }

    /// <summary>
    /// 主要科研成绩
    /// </summary>
    public class ProjectInfo
    {
        public string Date { get; set; }
        public string Name { get; set; }
        public string Source { get; set; }
        public string Job { get; set; }
    }

    /// <summary>
    /// 兼职情况（技术或学术）
    /// </summary>
    public class PartTimeInfo
    {
        public string StartDate { get; set; }
        public string EndDate { get; set; }
        public string PartTimeContent { get; set; }
        public string Job { get; set; }
    }

    /// <summary>
    /// 科技获奖和荣誉情况（省部级以上）
    /// </summary>
    public class HonorInfo
    {
        public string Date { get; set; }
        public string Name { get; set; }
        public string Level { get; set; }
        public string Order { get; set; }
    }
}