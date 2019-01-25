using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using Aspose.Words;
using Aspose.Words.Drawing;
using Aspose.Words.Tables;

namespace ReporterLoaders.DB.DocEntitys
{
    public class PersonInfo
    {
        /// <summary>
        /// 基础数据字典
        /// </summary>
        public SortedList<string, string> BaseInfoDict = new SortedList<string, string>();

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

        /// <summary>
        /// 主要著作和专利情况
        /// </summary>
        public List<ProductionInfo> ProductionInfoList = new List<ProductionInfo>();

        /// <summary>
        /// 头像
        /// </summary>
        public Image HeadImage { get; set; }

        /// <summary>
        /// 获取简历信息
        /// </summary>
        /// <param name="wordFiles"></param>
        /// <returns></returns>
        public static PersonInfo GetPersonInfoObj(string wordFiles)
        {
            PersonInfo pi = new PersonInfo();

            if (File.Exists(wordFiles))
            {
                Document doc = new Document(wordFiles);

                #region 头像
                MemoryStream ms = new MemoryStream();
                try
                {
                    Shape imgShape = (Shape)doc.GetChild(NodeType.Shape, 0, true);
                    if (imgShape != null)
                    {
                        imgShape.ImageData.Save(ms);
                        ms.Position = 0;
                        pi.HeadImage = Image.FromStream(ms);
                    }
                }
                finally
                {
                    ms.Dispose();
                }
                #endregion

                #region 提取基础信息
                Table t = (Table)doc.GetChild(NodeType.Table, 0, true);
                string sKey, sValue;
                sKey = string.Empty;
                sValue = string.Empty;
                foreach (Row r in t.Rows)
                {
                    foreach (Cell c in r.Cells)
                    {
                        string s = ReplaceAllEmptyChars(c.GetText());
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
                Table tt2 = (Table)doc.GetChild(NodeType.Table, 1, true);
                foreach (Row r in tt2.Rows)
                {
                    if (string.IsNullOrEmpty(r.Cells[0].GetText()) || r.Cells[0].GetText().Contains("时间"))
                    {
                        continue;
                    }

                    SchoolInfo sii = new SchoolInfo();
                    sii.StartDate = ReplaceAllEmptyChars(r.Cells[0].GetText());
                    sii.EndDate = ReplaceAllEmptyChars(r.Cells[1].GetText());
                    sii.SchoolName = ReplaceAllEmptyChars(r.Cells[2].GetText());
                    sii.Subject = ReplaceAllEmptyChars(r.Cells[3].GetText());
                    sii.License = ReplaceAllEmptyChars(r.Cells[4].GetText());

                    pi.SchoolInfoList.Add(sii);
                }
                #endregion

                #region 主要工作简历
                Table tt3 = (Table)doc.GetChild(NodeType.Table, 2, true);
                foreach (Row r in tt3.Rows)
                {
                    if (string.IsNullOrEmpty(r.Cells[0].GetText()) || r.Cells[0].GetText().Contains("时间"))
                    {
                        continue;
                    }

                    ResumeInfo wui = new ResumeInfo();
                    wui.StartDate = ReplaceAllEmptyChars(r.Cells[0].GetText());
                    wui.EndDate = ReplaceAllEmptyChars(r.Cells[1].GetText());
                    wui.WorkUnitAndJob = ReplaceAllEmptyChars(r.Cells[2].GetText());

                    pi.ResumeInfoList.Add(wui);
                }
                #endregion

                #region 主要科研成绩
                Table tt4 = (Table)doc.GetChild(NodeType.Table, 3, true);
                foreach (Row r in tt4.Rows)
                {
                    if (string.IsNullOrEmpty(r.Cells[0].GetText()) || r.Cells[0].GetText().Contains("时间"))
                    {
                        continue;
                    }

                    ProjectInfo ti = new ProjectInfo();
                    ti.Date = ReplaceAllEmptyChars(r.Cells[0].GetText());
                    ti.Name = ReplaceAllEmptyChars(r.Cells[1].GetText());
                    ti.Source = ReplaceAllEmptyChars(r.Cells[2].GetText());
                    ti.Job = ReplaceAllEmptyChars(r.Cells[3].GetText());

                    pi.ProjectInfoList.Add(ti);
                }
                #endregion

                #region 兼职情况
                Table tt5 = (Table)doc.GetChild(NodeType.Table, 4, true);
                foreach (Row r in tt5.Rows)
                {
                    if (string.IsNullOrEmpty(r.Cells[0].GetText()) || r.Cells[0].GetText().Contains("时间"))
                    {
                        continue;
                    }

                    PartTimeInfo pti = new PartTimeInfo();
                    pti.StartDate = ReplaceAllEmptyChars(r.Cells[0].GetText());
                    pti.EndDate = ReplaceAllEmptyChars(r.Cells[1].GetText());
                    pti.PartTimeContent = ReplaceAllEmptyChars(r.Cells[2].GetText());
                    pti.Job = ReplaceAllEmptyChars(r.Cells[3].GetText());

                    pi.PartTimeInfoList.Add(pti);
                }
                #endregion

                #region 科技获奖和荣誉情况
                Table tt6 = (Table)doc.GetChild(NodeType.Table, 5, true);
                foreach (Row r in tt6.Rows)
                {
                    if (string.IsNullOrEmpty(r.Cells[0].GetText()) || r.Cells[0].GetText().Contains("时间"))
                    {
                        continue;
                    }

                    HonorInfo hii = new HonorInfo();
                    hii.Date = ReplaceAllEmptyChars(r.Cells[0].GetText());
                    hii.Name = ReplaceAllEmptyChars(r.Cells[1].GetText());
                    hii.Level = ReplaceAllEmptyChars(r.Cells[2].GetText());
                    hii.Order = ReplaceAllEmptyChars(r.Cells[3].GetText());

                    pi.HonorInfoList.Add(hii);
                }
                #endregion

                #region 主要著作和专利情况
                Table tt7 = (Table)doc.GetChild(NodeType.Table, 6, true);
                foreach (Row r in tt7.Rows)
                {
                    if (string.IsNullOrEmpty(r.Cells[0].GetText()) || r.Cells[0].GetText().Contains("时间"))
                    {
                        continue;
                    }

                    ProductionInfo nii = new ProductionInfo();
                    nii.Date = ReplaceAllEmptyChars(r.Cells[0].GetText());
                    nii.Name = ReplaceAllEmptyChars(r.Cells[1].GetText());
                    nii.PrinterAndLicenseNo = ReplaceAllEmptyChars(r.Cells[2].GetText());
                    nii.Order = ReplaceAllEmptyChars(r.Cells[3].GetText());

                    pi.ProductionInfoList.Add(nii);
                }
                #endregion
            }

            return pi;
        }

        /// <summary>
        /// 替换空字符
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        protected static string ReplaceAllEmptyChars(string str)
        {
            if (string.IsNullOrEmpty(str))
            {
                return string.Empty;
            }
            else
            {
                return str.Trim().Replace("\a", string.Empty).Replace("\0", string.Empty).Replace("\n", string.Empty).Replace("\r", string.Empty);
            }
        }
    }

    /// <summary>
    /// 受教育情况
    /// </summary>
    public class SchoolInfo
    {
        /// <summary>
        /// 起始时间
        /// </summary>
        public string StartDate { get; set; }
        /// <summary>
        /// 终止时间
        /// </summary>
        public string EndDate { get; set; }
        /// <summary>
        /// 教育或培训单位
        /// </summary>
        public string SchoolName { get; set; }
        /// <summary>
        /// 学习专业
        /// </summary>
        public string Subject { get; set; }
        /// <summary>
        /// 学历
        /// </summary>
        public string License { get; set; }
    }

    /// <summary>
    /// 主要工作简历
    /// </summary>
    public class ResumeInfo
    {
        /// <summary>
        /// 起始时间
        /// </summary>
        public string StartDate { get; set; }
        /// <summary>
        /// 终止时间
        /// </summary>
        public string EndDate { get; set; }
        /// <summary>
        /// 在何单位、任何职务
        /// </summary>
        public string WorkUnitAndJob { get; set; }
    }

    /// <summary>
    /// 主要科研成绩
    /// </summary>
    public class ProjectInfo
    {
        /// <summary>
        /// 起止时间
        /// </summary>
        public string Date { get; set; }
        /// <summary>
        /// 项目名称
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 项目来源
        /// </summary>
        public string Source { get; set; }
        /// <summary>
        /// 主要贡献
        /// </summary>
        public string Job { get; set; }
    }

    /// <summary>
    /// 兼职情况（技术或学术）
    /// </summary>
    public class PartTimeInfo
    {
        /// <summary>
        /// 起始时间
        /// </summary>
        public string StartDate { get; set; }
        /// <summary>
        /// 终止时间
        /// </summary>
        public string EndDate { get; set; }
        /// <summary>
        /// 兼职情况
        /// </summary>
        public string PartTimeContent { get; set; }
        /// <summary>
        /// 职务
        /// </summary>
        public string Job { get; set; }
    }

    /// <summary>
    /// 科技获奖和荣誉情况（省部级以上）
    /// </summary>
    public class HonorInfo
    {
        /// <summary>
        /// 时间
        /// </summary>
        public string Date { get; set; }
        /// <summary>
        /// 项目（课题）名称或荣誉名称
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 种类及等级
        /// </summary>
        public string Level { get; set; }
        /// <summary>
        /// 排名
        /// </summary>
        public string Order { get; set; }
    }

    /// <summary>
    /// 主要著作和专利情况
    /// </summary>
    public class ProductionInfo
    {
        /// <summary>
        /// 时间
        /// </summary>
        public string Date { get; set; }
        /// <summary>
        /// 著作/专利名称
        /// </summary>
        public string Name { get; set; }
        /// <summary>
        /// 著作出版社/专利号
        /// </summary>
        public string PrinterAndLicenseNo { get; set; }
        /// <summary>
        /// 排名
        /// </summary>
        public string Order { get; set; }
    }
}