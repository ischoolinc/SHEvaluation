using Aspose.Words;
using FISCA.Presentation.Controls;
//using K12.Data;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SH_ClassYearScoreReport
{
    public partial class MainForm : BaseForm
    {
        BackgroundWorker bgWorkerReport;
        string _SchoolYear = "";
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();

        List<string> _ClassIDList = new List<string>();

        // 樣板設定檔
        private List<Configure> _ConfigureList = new List<Configure>();

        public Configure _Configure { get; private set; }
        private string _DefalutSchoolYear = "";
        private string _DefaultSemester = "";

        List<string> errorClassList = new List<string>();

        public MainForm()
        {
            InitializeComponent();

            bgWorkerReport = new BackgroundWorker();
            bgWorkerReport.DoWork += BgWorkerReport_DoWork;
            bgWorkerReport.ProgressChanged += BgWorkerReport_ProgressChanged;
            bgWorkerReport.WorkerReportsProgress = true;
            bgWorkerReport.RunWorkerCompleted += BgWorkerReport_RunWorkerCompleted;
        }

        public void SetClassIDList(List<string> classIDList)
        {
            _ClassIDList = classIDList;
        }

        private void BgWorkerReport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // 產生
            try
            {
                object[] objArray = (object[])e.Result;

                btnSaveConfig.Enabled = true;
                btnPrint.Enabled = true;

                if (errorClassList.Count > 0)
                {
                    string err = "下列班級因成績項目超過樣板支援上限，\n或者班級學生數超過樣板支援上限，\n超出部分科目成績無法印出，建議調整樣板內容。";
                    err += "\n" + string.Join("\n", errorClassList);
                    MessageBox.Show(err);
                }

                #region 儲存檔案

                string fileDateTime = DateTime.Now.ToString("yyyyMMddHHmmss");
                string reportName = "" + _SchoolYear + "學年度" + "班級學年成績單" + fileDateTime;
                string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                path = Path.Combine(path, reportName + ".docx");

                if (File.Exists(path))
                {
                    int i = 1;
                    while (true)
                    {
                        string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                        if (!File.Exists(newPath))
                        {
                            path = newPath;
                            break;
                        }
                    }
                }

                Document document = new Document();
                try
                {
                    document = (Document)objArray[0];
                    document.Save(path, SaveFormat.Docx);
                    System.Diagnostics.Process.Start(path);
                }
                catch
                {
                    System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = reportName + ".docx";
                    sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);

                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
                #endregion

                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級學年成績單 產生完成");

            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("產生過程發生錯誤，" + ex.Message);
            }
        }

        private void BgWorkerReport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 100)
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級學年成績單 產生完成");
            else
                FISCA.Presentation.MotherForm.SetStatusBarMessage("班級學年成績單 產生中...", e.ProgressPercentage);
        }

        private void BgWorkerReport_DoWork(object sender, DoWorkEventArgs e)
        {
            // 取得相關資料
            bgWorkerReport.ReportProgress(1);
            //errorMessage = "";
            errorClassList.Clear();
            AccessHelper accessHelper = new AccessHelper();
            List<ClassRecord> overflowRecords = new List<ClassRecord>();
            Exception exc = null;

            Configure conf = _Configure;
            //建立測試的選取學生
            List<ClassRecord> selectedClasses = accessHelper.ClassHelper.GetSelectedClass();
            List<StudentRecord> selectedStudents = new List<StudentRecord>();


            foreach (ClassRecord classRec in selectedClasses)
            {
                foreach (StudentRecord stuRec in classRec.Students)
                {
                    // 非一般生跳過
                    if (stuRec.Status != "一般")
                        continue;

                    if (!selectedStudents.Contains(stuRec))
                        selectedStudents.Add(stuRec);
                }
            }
            //建立合併欄位總表
            DataTable table = new DataTable();
            #region 所有的合併欄位
            table.Columns.Add("學校名稱");
            table.Columns.Add("學校地址");
            table.Columns.Add("學校電話");
            table.Columns.Add("科別名稱");
            table.Columns.Add("班級科別名稱");
            table.Columns.Add("班級");
            table.Columns.Add("班導師");
            table.Columns.Add("學年度");


            for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
            {
                table.Columns.Add("科目名稱" + subjectIndex);
                table.Columns.Add("學分數" + subjectIndex);
            }
            for (int i = 1; i <= conf.StudentLimit; i++)
            {
                table.Columns.Add("座號" + i);
                table.Columns.Add("學號" + i);
                table.Columns.Add("姓名" + i);
                table.Columns.Add("學年學業成績" + i);
                table.Columns.Add("學年專業科目成績" + i);
                table.Columns.Add("學年實習科目成績" + i);
                table.Columns.Add("學年應得學分" + i);
                table.Columns.Add("學年實得學分" + i);
                table.Columns.Add("應得學分累計" + i);
                table.Columns.Add("實得學分累計" + i);

                //table.Columns.Add("學生類別排名1名稱" + i);
                //table.Columns.Add("學生類別排名2名稱" + i);

                foreach (string key in new string[] { "學業", "專業科目", "實習科目" })
                {
                    foreach (string key2 in new string[] { "班", "年", "科" })
                    {
                        table.Columns.Add("學年" + key + "成績" + key2 + "排名" + i);
                        table.Columns.Add("學年" + key + "成績" + key2 + "排名母數" + i);
                    }
                }
            }

            List<string> itemNameList = new List<string>();
            List<string> itemTypeList = new List<string>();
            List<string> rankTypeList = new List<string>();

            Dictionary<string, string> rankTypeMapDict = new Dictionary<string, string>();
            rankTypeMapDict.Add("班排名", "班排名");
            rankTypeMapDict.Add("科排名", "科排名");
            rankTypeMapDict.Add("年排名", "年排名");
            //rankTypeMapDict.Add("類別1排名", "類別1");
            //rankTypeMapDict.Add("類別2排名", "類別2");

            rankTypeList.Add("班排名");
            rankTypeList.Add("科排名");
            rankTypeList.Add("年排名");
            //rankTypeList.Add("類別1排名");
            //rankTypeList.Add("類別2排名");

            // 排名、五標、組距
            //List<string> r2List = new List<string>();
            //r2List.Add("rank");
            //r2List.Add("matrix_count");
            //r2List.Add("pr");
            //r2List.Add("percentile");
            //r2List.Add("avg_top_25");
            //r2List.Add("avg_top_50");
            //r2List.Add("avg");
            //r2List.Add("avg_bottom_50");
            //r2List.Add("avg_bottom_25");
            //r2List.Add("pr_88");
            //r2List.Add("pr_75");
            //r2List.Add("pr_50");
            //r2List.Add("pr_25");
            //r2List.Add("pr_12");
            //r2List.Add("std_dev_pop");
            //r2List.Add("level_gte100");
            //r2List.Add("level_90");
            //r2List.Add("level_80");
            //r2List.Add("level_70");
            //r2List.Add("level_60");
            //r2List.Add("level_50");
            //r2List.Add("level_40");
            //r2List.Add("level_30");
            //r2List.Add("level_20");
            //r2List.Add("level_10");
            //r2List.Add("level_lt10");

            //List<string> r2ListOld = new List<string>();
            //r2ListOld.Add("pr");
            //r2ListOld.Add("percentile");
            //r2ListOld.Add("avg_top_25");
            //r2ListOld.Add("avg_top_50");
            //r2ListOld.Add("avg");
            //r2ListOld.Add("avg_bottom_50");
            //r2ListOld.Add("avg_bottom_25");
            //r2ListOld.Add("pr_88");
            //r2ListOld.Add("pr_75");
            //r2ListOld.Add("pr_50");
            //r2ListOld.Add("pr_25");
            //r2ListOld.Add("pr_12");
            //r2ListOld.Add("std_dev_pop");
            //r2ListOld.Add("level_gte100");
            //r2ListOld.Add("level_90");
            //r2ListOld.Add("level_80");
            //r2ListOld.Add("level_70");
            //r2ListOld.Add("level_60");
            //r2ListOld.Add("level_50");
            //r2ListOld.Add("level_40");
            //r2ListOld.Add("level_30");
            //r2ListOld.Add("level_20");
            //r2ListOld.Add("level_10");
            //r2ListOld.Add("level_lt10");

            //// 班級五標與組距
            //List<string> r2ListClass = new List<string>();
            //r2ListClass.Add("matrix_count");
            //r2ListClass.Add("avg_top_25");
            //r2ListClass.Add("avg_top_50");
            //r2ListClass.Add("avg");
            //r2ListClass.Add("avg_bottom_50");
            //r2ListClass.Add("avg_bottom_25");
            //r2ListClass.Add("pr_88");
            //r2ListClass.Add("pr_75");
            //r2ListClass.Add("pr_50");
            //r2ListClass.Add("pr_25");
            //r2ListClass.Add("pr_12");
            //r2ListClass.Add("std_dev_pop");
            //r2ListClass.Add("level_gte100");
            //r2ListClass.Add("level_90");
            //r2ListClass.Add("level_80");
            //r2ListClass.Add("level_70");
            //r2ListClass.Add("level_60");
            //r2ListClass.Add("level_50");
            //r2ListClass.Add("level_40");
            //r2ListClass.Add("level_30");
            //r2ListClass.Add("level_20");
            //r2ListClass.Add("level_10");
            //r2ListClass.Add("level_lt10");

            for (int Num = 1; Num <= conf.StudentLimit; Num++)
            {
                for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
                {

                    table.Columns.Add("科目成績" + Num + "-" + subjectIndex);
                    table.Columns.Add("科目學年學分" + Num + "-" + subjectIndex);

                    table.Columns.Add("班排名" + Num + "-" + subjectIndex);
                    table.Columns.Add("班排名母數" + Num + "-" + subjectIndex);
                    table.Columns.Add("科排名" + Num + "-" + subjectIndex);
                    table.Columns.Add("科排名母數" + Num + "-" + subjectIndex);
                    //table.Columns.Add("類別1排名" + Num + "-" + subjectIndex);
                    //table.Columns.Add("類別1排名母數" + Num + "-" + subjectIndex);
                    //table.Columns.Add("類別2排名" + Num + "-" + subjectIndex);
                    //table.Columns.Add("類別2排名母數" + Num + "-" + subjectIndex);
                    table.Columns.Add("年排名" + Num + "-" + subjectIndex);
                    table.Columns.Add("年排名母數" + Num + "-" + subjectIndex);

                    //table.Columns.Add("科目成績(原始)" + Num + "-" + subjectIndex);
                    //table.Columns.Add("班排名(原始)" + Num + "-" + subjectIndex);
                    //table.Columns.Add("班排名(原始)母數" + Num + "-" + subjectIndex);
                    //table.Columns.Add("科排名(原始)" + Num + "-" + subjectIndex);
                    //table.Columns.Add("科排名(原始)母數" + Num + "-" + subjectIndex);
                    //table.Columns.Add("年排名(原始)" + Num + "-" + subjectIndex);
                    //table.Columns.Add("年排名(原始)母數" + Num + "-" + subjectIndex);
                    //table.Columns.Add("類別1排名(原始)" + Num + "-" + subjectIndex);
                    //table.Columns.Add("類別1排名(原始)母數" + Num + "-" + subjectIndex);
                    //table.Columns.Add("類別2排名(原始)" + Num + "-" + subjectIndex);
                    //table.Columns.Add("類別2排名(原始)母數" + Num + "-" + subjectIndex);

                    //foreach (string s1 in rankTypeMapDict.Values)
                    //{
                    //    foreach (string s2 in r2ListOld)
                    //    {
                    //        string ssKey = "科目" + s1 + "_" + s2 + "_" + Num + "-" + subjectIndex;
                    //        string ssKey1 = "科目(原始)" + s1 + "_" + s2 + "_" + Num + "-" + subjectIndex;
                    //        table.Columns.Add(ssKey);
                    //        table.Columns.Add(ssKey1);
                    //    }
                    //}

                }

                // 學生學業
                //table.Columns.Add("加權平均" + Num);
                //table.Columns.Add("加權平均班排名" + Num);
                //table.Columns.Add("加權平均班排名母數" + Num);
                //table.Columns.Add("加權平均科排名" + Num);
                //table.Columns.Add("加權平均科排名母數" + Num);
                //table.Columns.Add("加權平均年排名" + Num);
                //table.Columns.Add("加權平均年排名母數" + Num);
                //table.Columns.Add("類別排名1" + Num);
                //table.Columns.Add("類別1加權平均" + Num);
                //table.Columns.Add("類別1加權平均排名" + Num);
                //table.Columns.Add("類別1加權平均排名母數" + Num);

                //table.Columns.Add("類別排名2" + Num);
                //table.Columns.Add("類別2加權平均" + Num);
                //table.Columns.Add("類別2加權平均排名" + Num);
                //table.Columns.Add("類別2加權平均排名母數" + Num);

                // 學生學業(原始)
                //table.Columns.Add("加權平均(原始)" + Num);
                //table.Columns.Add("加權平均班排名(原始)" + Num);
                //table.Columns.Add("加權平均班排名(原始)母數" + Num);
                //table.Columns.Add("加權平均科排名(原始)" + Num);
                //table.Columns.Add("加權平均科排名(原始)母數" + Num);
                //table.Columns.Add("加權平均年排名(原始)" + Num);
                //table.Columns.Add("加權平均年排名(原始)母數" + Num);
                //table.Columns.Add("類別1加權平均排名(原始)" + Num);
                //table.Columns.Add("類別1加權平均排名(原始)母數" + Num);
                //table.Columns.Add("類別2加權平均排名(原始)" + Num);
                //table.Columns.Add("類別2加權平均排名(原始)母數" + Num);


                //foreach (string s1 in rankTypeMapDict.Values)
                //{
                //    foreach (string s2 in r2ListOld)
                //    {
                //        string ssKey = "學業" + s1 + "_" + s2 + "_" + Num;
                //        string ssKey1 = "學業(原始)" + s1 + "_" + s2 + "_" + Num;
                //        table.Columns.Add(ssKey);
                //        table.Columns.Add(ssKey1);
                //    }
                //}



            }

            #region 組距及分析


            //// 產生班級分項五標、組距合併欄位
            //foreach (string s1 in itemNameList)
            //{
            //    foreach (string s2 in rankTypeList)
            //    {
            //        foreach (string r in r2ListClass)
            //        {
            //            string key = s1 + "_" + s2 + "_" + r;
            //            string key1 = s1 + "(原始)_" + s2 + "_" + r;
            //            table.Columns.Add(key);
            //            table.Columns.Add(key1);
            //        }
            //    }
            //}


            #region 各科目組距及分析
            // 產生班級科目五標、組距合併欄位
            //for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
            //{
            //    foreach (string s2 in rankTypeList)
            //    {
            //        foreach (string r in r2ListClass)
            //        {
            //            string key = s2 + "科目_" + r + "_" + subjectIndex;
            //            string key1 = s2 + "科目(原始)_" + r + "_" + subjectIndex;
            //            table.Columns.Add(key);
            //            table.Columns.Add(key1);
            //        }
            //    }
            //}
            #endregion

            #region 各科目五標、標準差 
            //for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
            //{
            //    foreach (string s2 in rankTypeList)
            //    {
            //        foreach (string r in r2ListClass)
            //        {
            //            string rankType = s2;
            //            if (s2 == "年排名")
            //                rankType = "年排名";
            //            //科目成績1_班排名_avg_top_50
            //            string key = "科目成績" + subjectIndex + "_" + rankType + "_" + r;
            //            string key1 = "科目成績(原始)" + subjectIndex + "_" + rankType + "_" + r;
            //            table.Columns.Add(key);
            //            table.Columns.Add(key1);
            //        }
            //    }
            //}
            #endregion



            #endregion

            #endregion

            //宣告產生的報表
            Aspose.Words.Document document = new Aspose.Words.Document();

            // 有問題的班級提示
            if (overflowRecords.Count > 0)
            {
                foreach (ClassRecord classRec in overflowRecords)
                {
                    if (!errorClassList.Contains(classRec.ClassName))
                        errorClassList.Add(classRec.ClassName);
                }
            }

            if (exc != null)
            {
                FISCA.Presentation.Controls.MsgBox.Show(exc.Message, "產生班級學年成績單發生錯誤", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }

            Document docTemplate = _Configure.Template;
            if (docTemplate == null)
                docTemplate = new Document(new MemoryStream(Properties.Resources.DefaultTemplate));

            accessHelper.StudentHelper.FillSchoolYearEntryScore(true, selectedStudents);
            accessHelper.StudentHelper.FillSchoolYearSubjectScore(true, selectedStudents);
            accessHelper.StudentHelper.FillSemesterSubjectScore(true, selectedStudents);
            accessHelper.StudentHelper.FillField("SchoolYearEntryClassRating", selectedStudents);

            Dictionary<string, Dictionary<string, SchoolYearSubjectScoreInfo>> YearSubjectScoreInfoDic = new Dictionary<string, Dictionary<string, SchoolYearSubjectScoreInfo>>();

            //id,分項,Score
            Dictionary<string, Dictionary<string, decimal?>> YearEntryScoreDic = new Dictionary<string, Dictionary<string, decimal?>>();

            Dictionary<string, DataRow> StudYearSubjRankData = new Dictionary<string, DataRow>();

            //student id <subj_校部定_必選修, credit>
            Dictionary<string, Dictionary<string, decimal>> StudentYearSubjCreditBySemsSubjMappingDic = new Dictionary<string, Dictionary<string, decimal>>();

            List<string> studIDList = (from data in selectedStudents select data.StudentID).ToList();

            // 取得學年成績固定排名(舊)
            StudYearSubjRankData = DAO.QueryData.GetStudentYearSubjectScoreRowBySchoolYear(studIDList, _SchoolYear);
            Dictionary<string, DataRow> StudYearEntryRankData = DAO.QueryData.GetStudentYearEntryScoreRowBySchoolYear(studIDList, _SchoolYear);

            foreach (StudentRecord stud in selectedStudents)
            {
                foreach (SemesterSubjectScoreInfo_New smScore in stud.SemesterSubjectScoreList)
                {
                    //if (smScore == null)
                    //    continue;

                    //if (smScore.Detail.GetAttribute("不計學分") == "是")
                    //    continue;

                    if (smScore.Detail.GetAttribute("不計學分") != "是")
                    {
                        // 應得學分累計
                        if (!Global._StudCreditDict.ContainsKey(stud.StudentID))
                            Global._StudCreditDict.Add(stud.StudentID, new DAO.StudCredit());
                        Global._StudCreditDict[stud.StudentID].shouldGetTotalCredit += smScore.CreditDec;

                        // 實得學分累計
                        if (smScore.Pass)
                            Global._StudCreditDict[stud.StudentID].gotTotalCredit += smScore.CreditDec;
                    }

                    if (smScore.SchoolYear.ToString() == _SchoolYear)
                    {
                        if (!StudentYearSubjCreditBySemsSubjMappingDic.ContainsKey(stud.StudentID))
                        {
                            StudentYearSubjCreditBySemsSubjMappingDic.Add(stud.StudentID, new Dictionary<string, decimal>());
                        }
                        string specifySubjectName = smScore.Detail.GetAttribute("指定學年科目名稱");
                        string subjName = (specifySubjectName == "") ? smScore.Subject : specifySubjectName;
                        string subjectKey = subjName + "_" + smScore.Detail.GetAttribute("修課校部訂") + "_" + smScore.Detail.GetAttribute("修課必選修") + "_" + smScore.CreditDec;

                        if (!StudentYearSubjCreditBySemsSubjMappingDic[stud.StudentID].ContainsKey(subjectKey))
                            StudentYearSubjCreditBySemsSubjMappingDic[stud.StudentID].Add(subjectKey, smScore.CreditDec);
                        else
                            StudentYearSubjCreditBySemsSubjMappingDic[stud.StudentID][subjectKey] += smScore.CreditDec;

                        if (smScore.Detail.GetAttribute("不計學分") != "是")
                        {
                            // 學年應得學分
                            Global._StudCreditDict[stud.StudentID].shouldGetCredit += smScore.CreditDec;

                            // 學年實得學分
                            if (smScore.Pass)
                                Global._StudCreditDict[stud.StudentID].gotCredit += smScore.CreditDec;
                        }
                    }
                }

                foreach (SchoolYearSubjectScoreInfo yearScore in stud.SchoolYearSubjectScoreList)
                {
                    if (yearScore.SchoolYear.ToString() == _SchoolYear)
                    {
                        if (!YearSubjectScoreInfoDic.ContainsKey(stud.StudentID))
                            YearSubjectScoreInfoDic.Add(stud.StudentID, new Dictionary<string, SchoolYearSubjectScoreInfo>());
                        string subjectKey = yearScore.Subject + "_" + yearScore.Detail.GetAttribute("校部定") + "_" + yearScore.Detail.GetAttribute("必選修") + "_" + yearScore.Detail.GetAttribute("識別學分數");
                        if (!YearSubjectScoreInfoDic[stud.StudentID].ContainsKey(subjectKey))
                            YearSubjectScoreInfoDic[stud.StudentID].Add(subjectKey, yearScore);
                    }
                }

                foreach (SchoolYearEntryScoreInfo yearEntryScore in stud.SchoolYearEntryScoreList)
                {
                    if (yearEntryScore.SchoolYear.ToString() == _SchoolYear)
                    {
                        if (!YearEntryScoreDic.ContainsKey(stud.StudentID))
                            YearEntryScoreDic.Add(stud.StudentID, new Dictionary<string, decimal?>());
                        if (yearEntryScore.Entry == "學業")
                            YearEntryScoreDic[stud.StudentID].Add("學業", yearEntryScore.Score);
                        if (yearEntryScore.Entry == "專業科目")
                            YearEntryScoreDic[stud.StudentID].Add("專業科目", yearEntryScore.Score);
                        if (yearEntryScore.Entry == "實習科目")
                            YearEntryScoreDic[stud.StudentID].Add("實習科目", yearEntryScore.Score);
                    }
                }
            }

            foreach (ClassRecord classRec in selectedClasses)
            {
                List<string> classSubjectsKey = new List<string>();
                List<string> classSubjects = new List<string>();

                DataRow row = table.NewRow();

                row["學校名稱"] = SmartSchool.Customization.Data.SystemInformation.SchoolChineseName;
                row["學校地址"] = SmartSchool.Customization.Data.SystemInformation.Address;
                row["學校電話"] = SmartSchool.Customization.Data.SystemInformation.Telephone;
                row["科別名稱"] = classRec.Department;
                row["學年度"] = conf.SchoolYear;
                row["班級"] = classRec.ClassName;
                row["班導師"] = classRec.RefTeacher == null ? "" : classRec.RefTeacher.TeacherName;

                int ClassStuNum = 0;
                foreach (StudentRecord stuRec in classRec.Students)
                {
                    string studentID = stuRec.StudentID;

                    ClassStuNum++;
                    if (ClassStuNum > conf.StudentLimit)
                    {
                        if (!overflowRecords.Contains(classRec))
                            overflowRecords.Add(classRec);
                        break;
                    }
                    row["座號" + ClassStuNum] = stuRec.SeatNo;
                    row["學號" + ClassStuNum] = stuRec.StudentNumber;
                    row["姓名" + ClassStuNum] = stuRec.StudentName;
                    if (YearEntryScoreDic.ContainsKey(stuRec.StudentID))
                        foreach (string key in new string[] { "學業", "專業科目", "實習科目" })
                        {
                            if (YearEntryScoreDic[stuRec.StudentID].ContainsKey(key))
                                row["學年" + key + "成績" + ClassStuNum] = YearEntryScoreDic[stuRec.StudentID][key].Value;
                        }
                    if (Global._TempEntryClassRankDict.ContainsKey(stuRec.StudentID))
                        foreach (string key in new string[] { "學業", "專業科目", "實習科目" })
                        {
                            if (Global._TempEntryClassRankDict[stuRec.StudentID].ContainsKey(key))
                            {
                                row["學年" + key + "成績班排名" + ClassStuNum] = Global._TempEntryClassRankDict[stuRec.StudentID][key].Rank;
                                row["學年" + key + "成績班排名母數" + ClassStuNum] = Global._TempEntryClassRankDict[stuRec.StudentID][key].RankT;
                            }
                        }

                    if (Global._TempEntryGradeYearRankDict.ContainsKey(stuRec.StudentID))
                        foreach (string key in new string[] { "學業", "專業科目", "實習科" })
                        {
                            if (Global._TempEntryGradeYearRankDict[stuRec.StudentID].ContainsKey(key))
                            {
                                row["學年" + key + "成績年排名" + ClassStuNum] = Global._TempEntryGradeYearRankDict[stuRec.StudentID][key].Rank;
                                row["學年" + key + "成績年排名母數" + ClassStuNum] = Global._TempEntryGradeYearRankDict[stuRec.StudentID][key].RankT;
                            }
                        }

                    if (Global._TempEntryDeptRankDict.ContainsKey(stuRec.StudentID))
                        foreach (string key in new string[] { "學業", "專業科目", "實習科" })
                        {
                            if (Global._TempEntryDeptRankDict[stuRec.StudentID].ContainsKey(key))
                            {
                                row["學年" + key + "成績科排名" + ClassStuNum] = Global._TempEntryDeptRankDict[stuRec.StudentID][key].Rank;
                                row["學年" + key + "成績科排名母數" + ClassStuNum] = Global._TempEntryDeptRankDict[stuRec.StudentID][key].RankT;
                            }
                        }

                    if (Global._StudCreditDict.ContainsKey(stuRec.StudentID))
                    {
                        row["學年應得學分" + ClassStuNum] = Global._StudCreditDict[stuRec.StudentID].shouldGetCredit;
                        row["學年實得學分" + ClassStuNum] = Global._StudCreditDict[stuRec.StudentID].gotCredit;
                        row["應得學分累計" + ClassStuNum] = Global._StudCreditDict[stuRec.StudentID].shouldGetTotalCredit;
                        row["實得學分累計" + ClassStuNum] = Global._StudCreditDict[stuRec.StudentID].gotTotalCredit;
                    }

                    #region 整理班級中列印科目
                    if (YearSubjectScoreInfoDic.ContainsKey(stuRec.StudentID))
                    {
                        foreach (var subjectKey in YearSubjectScoreInfoDic[studentID].Keys)
                        {
                            string[] subjArray = subjectKey.Split('_');
                            if (!classSubjectsKey.Contains(subjectKey))
                                classSubjectsKey.Add(subjectKey);

                            if (!classSubjects.Contains(subjArray[0]))
                                classSubjects.Add(subjArray[0]);
                        }
                    }

                    #endregion
                }
                classSubjects.Sort(new StringComparer(Global._SubjectListOrder.ToArray()));
                classSubjectsKey.Sort(new StringComparer(Global._SubjectListOrder.ToArray()));

                int subjectIndex = 1;
                foreach (string subjectNameKey in classSubjectsKey)//classSubjects
                {
                    string[] subjectInfoArray = subjectNameKey.Split('_');
                    if (subjectIndex <= conf.SubjectLimit)
                    {
                        row["科目名稱" + subjectIndex] = subjectInfoArray[0];

                        ClassStuNum = 0;
                        foreach (StudentRecord stuRec in classRec.Students)
                        {
                            ClassStuNum++;
                            if (ClassStuNum > conf.StudentLimit)
                                break;
                            string studentID = stuRec.StudentID;

                            if (YearSubjectScoreInfoDic.ContainsKey(studentID))
                            {
                                if (YearSubjectScoreInfoDic[studentID].ContainsKey(subjectNameKey))
                                {
                                    row["科目成績" + ClassStuNum + "-" + subjectIndex] = YearSubjectScoreInfoDic[studentID][subjectNameKey].Score;

                                    if (Global._TempSubjClassRankDict.ContainsKey(studentID))
                                    {
                                        if (Global._TempSubjClassRankDict[studentID].ContainsKey(subjectInfoArray[0]))
                                        {
                                            row["班排名" + ClassStuNum + "-" + subjectIndex] = Global._TempSubjClassRankDict[studentID][subjectInfoArray[0]].Rank;
                                            row["班排名母數" + ClassStuNum + "-" + subjectIndex] = Global._TempSubjClassRankDict[studentID][subjectInfoArray[0]].RankT;
                                        }
                                    }
                                    if (Global._TempSubjGradeYearRankDict.ContainsKey(studentID))
                                        if (Global._TempSubjGradeYearRankDict[studentID].ContainsKey(subjectInfoArray[0]))
                                        {
                                            row["年排名" + ClassStuNum + "-" + subjectIndex] = Global._TempSubjGradeYearRankDict[studentID][subjectInfoArray[0]].Rank;
                                            row["年排名母數" + ClassStuNum + "-" + subjectIndex] = Global._TempSubjGradeYearRankDict[studentID][subjectInfoArray[0]].RankT;
                                        }
                                    if (Global._TempSubjDeptRankDict.ContainsKey(studentID))
                                        if (Global._TempSubjDeptRankDict[studentID].ContainsKey(subjectInfoArray[0]))
                                        {
                                            row["科排名" + ClassStuNum + "-" + subjectIndex] = Global._TempSubjDeptRankDict[studentID][subjectInfoArray[0]].Rank;
                                            row["科排名母數" + ClassStuNum + "-" + subjectIndex] = Global._TempSubjDeptRankDict[studentID][subjectInfoArray[0]].RankT;
                                        }

                                }
                            }
                            if (StudentYearSubjCreditBySemsSubjMappingDic.ContainsKey(studentID))
                            {
                                if (StudentYearSubjCreditBySemsSubjMappingDic[studentID].ContainsKey(subjectNameKey))
                                    row["科目學年學分" + ClassStuNum + "-" + subjectIndex] = StudentYearSubjCreditBySemsSubjMappingDic[studentID][subjectNameKey];
                            }

                        }
                        subjectIndex++;
                    }
                    else
                    {
                        //科目數超過的班級
                        if (!errorClassList.Contains(classRec.ClassName))
                        {
                            errorClassList.Add(classRec.ClassName);
                        }
                    }
                }

                table.Rows.Add(row);
            }

            docTemplate.MailMerge.Execute(table);
            docTemplate.MailMerge.RemoveEmptyParagraphs = true;
            docTemplate.MailMerge.DeleteFields();

            #region Word 合併列印

            try
            {
                e.Result = new object[] { docTemplate };
            }
            catch (Exception exow)
            {
                MsgBox.Show(exow.Message);
                //throw exow;
            }

            #endregion
        }

        public string parseScore(string str)
        {
            string value = "";
            decimal dc;
            if (decimal.TryParse(str, out dc))
            {
                //value = Math.Round(dc, parseNumber).ToString();
            }

            return value;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // 檢查設定學年度學期與目前系統預設是否相同
            string s1 = cboSchoolYear.Text;
            string s2 = K12.Data.School.DefaultSchoolYear;

            string semester = K12.Data.School.DefaultSemester;
            if (s1 != s2)
                if (FISCA.Presentation.Controls.MsgBox.Show("所選學年度學期與系統學年度學期不相同，請問是否繼續?", "學年度學期不同", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.No)
                    return;

            if (s1 == s2 && semester == "1")
                if (FISCA.Presentation.Controls.MsgBox.Show("目前系統為第一學期，請問是否繼續?", "學期提醒", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.No)
                    return;
            _SchoolYear = s1;
            SaveTemplate(null, null);

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();

            // 列印報表
            bgWorkerReport.RunWorkerAsync();
        }

        private void MainForm_Load(object sender, EventArgs e)
        {
            picLoding.Visible = true;
            //預設學年度學期
            _DefalutSchoolYear = "" + K12.Data.School.DefaultSchoolYear;
            _DefaultSemester = "" + K12.Data.School.DefaultSemester;

            _ConfigureList = _AccessHelper.Select<Configure>();

            cboConfigure.Items.Clear();
            foreach (var item in _ConfigureList)
            {
                cboConfigure.Items.Add(item);
            }
            cboConfigure.Items.Add(new Configure() { Name = "新增" });
            if (_ConfigureList.Count > 0)
                cboConfigure.SelectedIndex = 0;

            int i;
            if (int.TryParse(_DefalutSchoolYear, out i))
            {
                for (int j = 0; j < 5; j++)
                {
                    cboSchoolYear.Items.Add("" + (i - j));
                }
                cboSchoolYear.SelectedIndex = 0;
            }
            picLoding.Visible = false;
        }

        private void SaveTemplate(object sender, EventArgs e)
        {
            if (_Configure == null) return;
            _Configure.SchoolYear = cboSchoolYear.Text;


            _Configure.Encode();
            _Configure.Save();
        }

        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
            picLoding.Visible = true;
            if (cboConfigure.SelectedIndex == cboConfigure.Items.Count - 1)
            {
                //新增
                btnSaveConfig.Enabled = btnPrint.Enabled = false;
                NewConfigure dialog = new NewConfigure();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    _Configure = new Configure();
                    _Configure.Name = dialog.ConfigName;
                    _Configure.Template = dialog.Template;
                    _Configure.SubjectLimit = dialog.SubjectLimit;
                    _Configure.StudentLimit = dialog.StudentLimit;
                    _Configure.SchoolYear = _DefalutSchoolYear;
                    _Configure.Semester = _DefaultSemester;


                    _ConfigureList.Add(_Configure);
                    cboConfigure.Items.Insert(cboConfigure.SelectedIndex, _Configure);
                    cboConfigure.SelectedIndex = cboConfigure.SelectedIndex - 1;

                    _Configure.Encode();
                    _Configure.Save();
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            }
            else
            {
                if (cboConfigure.SelectedIndex >= 0)
                {
                    btnSaveConfig.Enabled = btnPrint.Enabled = true;
                    _Configure = _ConfigureList[cboConfigure.SelectedIndex];
                    if (_Configure.Template == null)
                        _Configure.Decode();
                    if (!cboSchoolYear.Items.Contains(_Configure.SchoolYear))
                        cboSchoolYear.Items.Add(_Configure.SchoolYear);
                    cboSchoolYear.Text = _Configure.SchoolYear;
                }
                else
                {
                    _Configure = null;
                    cboSchoolYear.SelectedIndex = -1;
                }
            }

            picLoding.Visible = false;
        }

        private void lnkViewMapping_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            picLoding.Visible = true;
            lnkViewMapping.Enabled = false;
            Global.CreateFieldTemplate();
            lnkViewMapping.Enabled = true;
            picLoding.Visible = false;
        }

        private void lnkDownloadTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this._Configure == null) return;
            picLoding.Visible = true;
            lnkDownloadTemplate.Enabled = false;
            #region 儲存檔案
            string inputReportName = "班級學年成績單樣板(" + this._Configure.Name + ").docx";
            string reportName = Global.ParseFileName(inputReportName);

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".docx");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                System.IO.FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                this._Configure.Template.Save(stream, Aspose.Words.SaveFormat.Docx);
                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".docx";
                sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.DefaultTemplate, 0, Properties.Resources.DefaultTemplate.Length);
                        stream.Flush();
                        stream.Close();

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
            lnkDownloadTemplate.Enabled = true;
            picLoding.Visible = false;
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (_Configure == null) return;

            if (MessageBox.Show("樣板刪除後將無法回復，確定刪除樣板?", "刪除樣板", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
            {
                _ConfigureList.Remove(_Configure);
                if (_Configure.UID != "")
                {
                    _Configure.Deleted = true;
                    _Configure.Save();
                }
                var conf = _Configure;
                cboConfigure.SelectedIndex = -1;
                cboConfigure.Items.Remove(conf);
            }

        }

        private void lnkUploadTenplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            picLoding.Visible = true;
            if (_Configure == null) return;
            lnkUploadTenplate.Enabled = false;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    this._Configure.Template = new Aspose.Words.Document(dialog.FileName);
                    List<string> fields = new List<string>(this._Configure.Template.MailMerge.GetFieldNames());
                    this._Configure.SubjectLimit = 0;
                    while (fields.Contains("科目名稱" + (this._Configure.SubjectLimit + 1)))
                    {
                        this._Configure.SubjectLimit++;
                    }
                    this._Configure.StudentLimit = 0;
                    while (fields.Contains("姓名" + (this._Configure.StudentLimit + 1)))
                    {
                        this._Configure.StudentLimit++;
                    }
                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗");
                }
            }
            lnkUploadTenplate.Enabled = true;
            picLoding.Visible = false;
        }
    }
}
