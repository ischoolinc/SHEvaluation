using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SmartSchool.Customization.Data;
using System.Data;
using System.IO;
using System.Threading;
using FISCA.Permission;
using SmartSchool.Customization.Data.StudentExtension;

namespace ClassSemesterScoreReportFixed_SH
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            string regCode = "SH.Class.ClassSemesterScoreReportFixed_SH";
            var btn = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["班級學期成績單(固定排名)"];

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                btn.Enable = false;
                if (UserAcl.Current[regCode].Executable && K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                    btn.Enable = true;

            };
            btn.Click += new EventHandler(Program_Click);

            Catalog cat1 = RoleAclSource.Instance["班級"]["功能按鈕"];
            cat1.Add(new RibbonFeature(regCode, "班級學期成績單(固定排名)"));

        }

        private static string GetNumber(decimal? p)
        {
            if (p == null) return "";
            string levelNumber;
            switch (((int)p.Value))
            {
                #region 對應levelNumber
                case 0:
                    levelNumber = "";
                    break;
                case 1:
                    levelNumber = "Ⅰ";
                    break;
                case 2:
                    levelNumber = "Ⅱ";
                    break;
                case 3:
                    levelNumber = "Ⅲ";
                    break;
                case 4:
                    levelNumber = "Ⅳ";
                    break;
                case 5:
                    levelNumber = "Ⅴ";
                    break;
                case 6:
                    levelNumber = "Ⅵ";
                    break;
                case 7:
                    levelNumber = "Ⅶ";
                    break;
                case 8:
                    levelNumber = "Ⅷ";
                    break;
                case 9:
                    levelNumber = "Ⅸ";
                    break;
                case 10:
                    levelNumber = "Ⅹ";
                    break;
                default:
                    levelNumber = "" + (p);
                    break;
                    #endregion
            }
            return levelNumber;
        }

        private static void Program_Click(object sender_, EventArgs e_)
        {
            ConfigForm form = new ConfigForm();
            if (form.ShowDialog() == DialogResult.OK)
            {
                AccessHelper accessHelper = new AccessHelper();
                List<ClassRecord> overflowRecords = new List<ClassRecord>();
                Exception exc = null;
                //取得列印設定
                Configure conf = form.Configure;
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
                table.Columns.Add("學期");
                table.Columns.Add("類別排名1");
                table.Columns.Add("類別排名2");
                //«通訊地址»«通訊地址郵遞區號»«通訊地址內容»
                //«戶籍地址»«戶籍地址郵遞區號»«戶籍地址內容»
                //«監護人»«父親»«母親»«科別名稱»
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
                    table.Columns.Add("學生類別排名1名稱" + i);
                    table.Columns.Add("學生類別排名2名稱" + i);
                }

                List<string> itemNameList = new List<string>();
                List<string> itemTypeList = new List<string>();
                List<string> rankTypeList = new List<string>();

                Dictionary<string, string> rankTypeMapDict = new Dictionary<string, string>();
                rankTypeMapDict.Add("班排名", "班排名");
                rankTypeMapDict.Add("科排名", "科排名");
                rankTypeMapDict.Add("年排名", "全校排名");
                rankTypeMapDict.Add("類別1排名", "類別1");
                rankTypeMapDict.Add("類別2排名", "類別2");

                itemNameList.Add("學業");
                itemNameList.Add("實習科目");
                itemNameList.Add("專業科目");


                itemTypeList.Add("學期/科目成績(原始)");
                itemTypeList.Add("學期/分項成績");
                itemTypeList.Add("學期/科目成績");
                itemTypeList.Add("學期/分項成績(原始)");

                rankTypeList.Add("班排名");
                rankTypeList.Add("科排名");
                rankTypeList.Add("年排名");
                rankTypeList.Add("類別1排名");
                rankTypeList.Add("類別2排名");

                // 排名、五標、組距
                List<string> r2List = new List<string>();
                r2List.Add("rank");
                r2List.Add("matrix_count");
                r2List.Add("pr");
                r2List.Add("percentile");
                r2List.Add("avg_top_25");
                r2List.Add("avg_top_50");
                r2List.Add("avg");
                r2List.Add("avg_bottom_50");
                r2List.Add("avg_bottom_25");
                r2List.Add("level_gte100");
                r2List.Add("level_90");
                r2List.Add("level_80");
                r2List.Add("level_70");
                r2List.Add("level_60");
                r2List.Add("level_50");
                r2List.Add("level_40");
                r2List.Add("level_30");
                r2List.Add("level_20");
                r2List.Add("level_10");
                r2List.Add("level_lt10");

                List<string> r2ListOld = new List<string>();
                r2ListOld.Add("pr");
                r2ListOld.Add("percentile");
                r2ListOld.Add("avg_top_25");
                r2ListOld.Add("avg_top_50");
                r2ListOld.Add("avg");
                r2ListOld.Add("avg_bottom_50");
                r2ListOld.Add("avg_bottom_25");
                r2ListOld.Add("level_gte100");
                r2ListOld.Add("level_90");
                r2ListOld.Add("level_80");
                r2ListOld.Add("level_70");
                r2ListOld.Add("level_60");
                r2ListOld.Add("level_50");
                r2ListOld.Add("level_40");
                r2ListOld.Add("level_30");
                r2ListOld.Add("level_20");
                r2ListOld.Add("level_10");
                r2ListOld.Add("level_lt10");

                // 班級五標與組距
                List<string> r2ListClass = new List<string>();
                r2ListClass.Add("matrix_count");
                r2ListClass.Add("avg_top_25");
                r2ListClass.Add("avg_top_50");
                r2ListClass.Add("avg");
                r2ListClass.Add("avg_bottom_50");
                r2ListClass.Add("avg_bottom_25");
                r2ListClass.Add("level_gte100");
                r2ListClass.Add("level_90");
                r2ListClass.Add("level_80");
                r2ListClass.Add("level_70");
                r2ListClass.Add("level_60");
                r2ListClass.Add("level_50");
                r2ListClass.Add("level_40");
                r2ListClass.Add("level_30");
                r2ListClass.Add("level_20");
                r2ListClass.Add("level_10");
                r2ListClass.Add("level_lt10");

                for (int Num = 1; Num <= conf.StudentLimit; Num++)
                {
                    for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
                    {

                        table.Columns.Add("科目成績" + Num + "-" + subjectIndex);
                        table.Columns.Add("班排名" + Num + "-" + subjectIndex);
                        table.Columns.Add("班排名母數" + Num + "-" + subjectIndex);
                        table.Columns.Add("科排名" + Num + "-" + subjectIndex);
                        table.Columns.Add("科排名母數" + Num + "-" + subjectIndex);
                        table.Columns.Add("類別1排名" + Num + "-" + subjectIndex);
                        table.Columns.Add("類別1排名母數" + Num + "-" + subjectIndex);
                        table.Columns.Add("類別2排名" + Num + "-" + subjectIndex);
                        table.Columns.Add("類別2排名母數" + Num + "-" + subjectIndex);
                        table.Columns.Add("全校排名" + Num + "-" + subjectIndex);
                        table.Columns.Add("全校排名母數" + Num + "-" + subjectIndex);

                        table.Columns.Add("科目成績(原始)" + Num + "-" + subjectIndex);
                        table.Columns.Add("班排名(原始)" + Num + "-" + subjectIndex);
                        table.Columns.Add("班排名(原始)母數" + Num + "-" + subjectIndex);
                        table.Columns.Add("科排名(原始)" + Num + "-" + subjectIndex);
                        table.Columns.Add("科排名(原始)母數" + Num + "-" + subjectIndex);
                        table.Columns.Add("全校排名(原始)" + Num + "-" + subjectIndex);
                        table.Columns.Add("全校排名(原始)母數" + Num + "-" + subjectIndex);
                        table.Columns.Add("類別1排名(原始)" + Num + "-" + subjectIndex);
                        table.Columns.Add("類別1排名(原始)母數" + Num + "-" + subjectIndex);
                        table.Columns.Add("類別2排名(原始)" + Num + "-" + subjectIndex);
                        table.Columns.Add("類別2排名(原始)母數" + Num + "-" + subjectIndex);

                        foreach (string s1 in rankTypeMapDict.Values)
                        {
                            foreach (string s2 in r2ListOld)
                            {
                                string ssKey = "科目" + s1 + "_" + s2 + "_" + Num + "-" + subjectIndex;
                                string ssKey1 = "科目(原始)" + s1 + "_" + s2 + "_" + Num + "-" + subjectIndex;
                                table.Columns.Add(ssKey);
                                table.Columns.Add(ssKey1);
                            }
                        }

                    }


                    // 學生學業
                    table.Columns.Add("加權平均" + Num);
                    table.Columns.Add("加權平均班排名" + Num);
                    table.Columns.Add("加權平均班排名母數" + Num);
                    table.Columns.Add("加權平均科排名" + Num);
                    table.Columns.Add("加權平均科排名母數" + Num);
                    table.Columns.Add("加權平均全校排名" + Num);
                    table.Columns.Add("加權平均全校排名母數" + Num);
                    table.Columns.Add("類別排名1" + Num);
                    table.Columns.Add("類別1加權平均" + Num);
                    table.Columns.Add("類別1加權平均排名" + Num);
                    table.Columns.Add("類別1加權平均排名母數" + Num);

                    table.Columns.Add("類別排名2" + Num);
                    table.Columns.Add("類別2加權平均" + Num);
                    table.Columns.Add("類別2加權平均排名" + Num);
                    table.Columns.Add("類別2加權平均排名母數" + Num);

                    // 學生學業(原始)
                    table.Columns.Add("加權平均(原始)" + Num);
                    table.Columns.Add("加權平均班排名(原始)" + Num);
                    table.Columns.Add("加權平均班排名(原始)母數" + Num);
                    table.Columns.Add("加權平均科排名(原始)" + Num);
                    table.Columns.Add("加權平均科排名(原始)母數" + Num);
                    table.Columns.Add("加權平均全校排名(原始)" + Num);
                    table.Columns.Add("加權平均全校排名(原始)母數" + Num);
                    table.Columns.Add("類別1加權平均排名(原始)" + Num);
                    table.Columns.Add("類別1加權平均排名(原始)母數" + Num);
                    table.Columns.Add("類別2加權平均排名(原始)" + Num);
                    table.Columns.Add("類別2加權平均排名(原始)母數" + Num);


                    foreach (string s1 in rankTypeMapDict.Values)
                    {
                        foreach (string s2 in r2ListOld)
                        {
                            string ssKey = "學業" + s1 + "_" + s2 + "_" + Num;
                            string ssKey1 = "學業(原始)" + s1 + "_" + s2 + "_" + Num;
                            table.Columns.Add(ssKey);
                            table.Columns.Add(ssKey1);
                        }
                    }



                    // 學生學分
                    table.Columns.Add("應得學分" + Num);
                    table.Columns.Add("實得學分" + Num);
                    table.Columns.Add("應得學分累計" + Num);
                    table.Columns.Add("實得學分累計" + Num);
                }

                #region 瘋狂的組距及分析


                // 產生班級分項五標、組距合併欄位
                foreach (string s1 in itemNameList)
                {
                    foreach (string s2 in rankTypeList)
                    {
                        foreach (string r in r2ListClass)
                        {
                            string key = s1 + "_" + s2 + "_" + r;
                            string key1 = s1 + "(原始)_" + s2 + "_" + r;
                            table.Columns.Add(key);
                            table.Columns.Add(key1);
                        }
                    }
                }


                #region 各科目組距及分析
                // 產生班級科目五標、組距合併欄位
                for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
                {
                    foreach (string s2 in rankTypeList)
                    {
                        foreach (string r in r2ListClass)
                        {
                            string key = s2 + "科目_" + r + "_" + subjectIndex;
                            string key1 = s2 + "科目(原始)_" + r + "_" + subjectIndex;
                            table.Columns.Add(key);
                            table.Columns.Add(key1);
                        }
                    }
                }
                #endregion

                #endregion
                #endregion


                //StreamWriter sw1 = new StreamWriter(Application.StartupPath + "\\table欄位.txt");
                //foreach (DataColumn dc in table.Columns)
                //    sw1.WriteLine(dc.Caption);
                //sw1.Close();

                //宣告產生的報表
                Aspose.Words.Document document = new Aspose.Words.Document();
                //用一個BackgroundWorker包起來
                System.ComponentModel.BackgroundWorker bkw = new System.ComponentModel.BackgroundWorker();
                bkw.WorkerReportsProgress = true;
                System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 班級學期成績單(固定排名)產生 S");
                bkw.ProgressChanged += delegate (object sender, System.ComponentModel.ProgressChangedEventArgs e)
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage(" 班級學期成績單產生中", e.ProgressPercentage);
                    System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + "班級學期成績單(固定排名)產生 " + e.ProgressPercentage);
                };
                bkw.RunWorkerCompleted += delegate
                {
                    System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 班級學期成績單(固定排名)產生 E");
                    string err = "下列班級因成績項目超過樣板支援上限，\n或者班級學生數超過樣板支援上限，\n超出部分科目成績無法印出，建議調整樣板內容。";
                    if (overflowRecords.Count > 0)
                    {
                        foreach (ClassRecord classRec in overflowRecords)
                        {
                            err += "\n" + classRec.ClassName;
                        }
                    }


                    if (exc != null)
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("產生班級學期成績單(固定排名)發生錯誤", exc.Message);

                        // throw new Exception("產生班級學期成績單發生錯誤", exc);
                    }

                    #region 儲存檔案
                    string inputReportName = "班級學期成績單（固定排名）";
                    string reportName = inputReportName;

                    string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    path = Path.Combine(path, reportName + ".doc");

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
                        document.Save(path, Aspose.Words.SaveFormat.Doc);
                        System.Diagnostics.Process.Start(path);
                    }
                    catch (Exception ex)
                    {
                        System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                        sd.Title = "另存新檔";
                        sd.FileName = reportName + ".doc";
                        sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
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
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("班級學期成績單(固定排名)產生完成。", 100);
                    if (overflowRecords.Count > 0)
                        MessageBox.Show(err);
                };

                Dictionary<string, DataRow> StudSubjRankData = new Dictionary<string, DataRow>();
                Dictionary<string, DataRow> StudSemsRankData = new Dictionary<string, DataRow>();
                Dictionary<string, Dictionary<string, Dictionary<string, string>>> SemsScoreRankMatrixDataDict = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

                // 班五標、組距
                Dictionary<string, Dictionary<string, Dictionary<string, string>>> SemsScoreRankMatrixDataClassDict = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

                bkw.DoWork += delegate (object sender, System.ComponentModel.DoWorkEventArgs e)
                {
                    #region 偷跑取得考試成績


                    // 取得學期科目成績
                    Dictionary<string, Dictionary<string, Dictionary<string, SemesterSubjectScoreInfo>>> studentSubjectScoreDict = new Dictionary<string, Dictionary<string, Dictionary<string, SemesterSubjectScoreInfo>>>();

                    ManualResetEvent scoreReady = new ManualResetEvent(false);
                    ManualResetEvent elseReady = new ManualResetEvent(false);
                    int sSchoolYear = 0, sSemester = 0;
                    bkw.ReportProgress(3);
                    new Thread(new ThreadStart(delegate
                    {
                        // 取得學生學期科目成績

                        int.TryParse(conf.SchoolYear, out sSchoolYear);
                        int.TryParse(conf.Semester, out sSemester);
                        List<string> studIDList = (from data in selectedStudents select data.StudentID).ToList();
                        // 取得學期成績固定排名
                        StudSubjRankData = DAO.QueryData.GetStudentSemesterSubjectScoreRowBySchoolYearSemester(studIDList, sSchoolYear, sSemester);
                        StudSemsRankData = DAO.QueryData.GetStudentSemesterScoreRowBySchoolYearSemester(studIDList, sSchoolYear, sSemester);


                        // 取得新版固定排名、PR、百分比、五標、組距
                        SemsScoreRankMatrixDataDict = DAO.QueryData.GetSemsScoreRankMatrixDataDict(sSchoolYear + "", sSemester + "", studIDList);

                        List<string> classIDList = new List<string>();

                        foreach (ClassRecord classRec in selectedClasses)
                        {
                            classIDList.Add(classRec.ClassID);
                        }

                        SemsScoreRankMatrixDataClassDict = DAO.QueryData.GetSemsScoreRankMatrixDataByClassIDDict(sSchoolYear + "", sSemester + "", classIDList);



                        List<string> tmpSS = new List<string>();
                        foreach (string s1 in SemsScoreRankMatrixDataDict.Keys)
                        {
                            foreach (string s2 in SemsScoreRankMatrixDataDict[s1].Keys)
                            {
                                if (!tmpSS.Contains(s2))
                                    tmpSS.Add(s2);
                            }
                        }

                        //StreamWriter sw = new StreamWriter(Application.StartupPath + "\\學期固定排名變數.txt");
                        //foreach (string ss in tmpSS)
                        //    sw.WriteLine(ss);
                        //sw.Close();


                        scoreReady.Set();
                        elseReady.Set();
                    })).Start();
                    #endregion
                    try
                    {
                        string key = "";
                        bkw.ReportProgress(10);
                        #region 整理同年級學生
                        //整理選取學生的年級
                        Dictionary<string, List<StudentRecord>> gradeyearStudents = new Dictionary<string, List<StudentRecord>>();
                        foreach (ClassRecord classRec in selectedClasses)
                        {
                            string grade = "";
                            grade = "" + classRec.GradeYear;
                            if (!gradeyearStudents.ContainsKey(grade))
                                gradeyearStudents.Add(grade, new List<StudentRecord>());
                            foreach (var studentRec in classRec.Students)
                            {
                                gradeyearStudents[grade].Add(studentRec);
                            }
                        }
                        foreach (var classRec in accessHelper.ClassHelper.GetAllClass())
                        {
                            if (!selectedClasses.Contains(classRec) && gradeyearStudents.ContainsKey("" + classRec.GradeYear))
                            {
                                //用班級去取出可能有相關的學生
                                foreach (var studentRec in classRec.Students)
                                {
                                    string grade = "";
                                    if (studentRec.RefClass != null)
                                        grade = "" + studentRec.RefClass.GradeYear;
                                    if (!gradeyearStudents[grade].Contains(studentRec))
                                        gradeyearStudents[grade].Add(studentRec);
                                }
                            }
                        }
                        #endregion
                        bkw.ReportProgress(15);
                        #region 取得學生類別
                        Dictionary<string, List<K12.Data.StudentTagRecord>> studentTags = new Dictionary<string, List<K12.Data.StudentTagRecord>>();
                        List<string> list = new List<string>();
                        foreach (var sRecs in gradeyearStudents.Values)
                        {
                            foreach (var stuRec in sRecs)
                            {
                                list.Add(stuRec.StudentID);
                            }
                        }
                        foreach (var tag in K12.Data.StudentTag.SelectByStudentIDs(list))
                        {
                            if (!studentTags.ContainsKey(tag.RefStudentID))
                                studentTags.Add(tag.RefStudentID, new List<K12.Data.StudentTagRecord>());
                            studentTags[tag.RefStudentID].Add(tag);
                        }
                        #endregion
                        bkw.ReportProgress(20);
                        //等到成績載完
                        scoreReady.WaitOne();
                        bkw.ReportProgress(35);
                        int progressCount = 0;

                      
                        Dictionary<string, Dictionary<string, SemesterSubjectScoreInfo>> SemesterSubjectScoreInfoDict = new Dictionary<string, Dictionary<string, SemesterSubjectScoreInfo>>();

                        #region 學生學期科目成績並整理
                        try
                        {
                            // 學生學分
                            Global._StudCreditDict.Clear();

                            // 取得學生學期科目成績並整理
                            accessHelper.StudentHelper.FillSemesterSubjectScore(true, selectedStudents);
                            // 整理符合畫面所選學年度學期科目成績
                            foreach (StudentRecord stud in selectedStudents)
                            {
                                if (!Global._StudCreditDict.ContainsKey(stud.StudentID))
                                {
                                    Global._StudCreditDict.Add(stud.StudentID, new DAO.StudCredit());
                                    Global._StudCreditDict[stud.StudentID].StudentID = stud.StudentID;
                                }

                                foreach (SemesterSubjectScoreInfo smScore in stud.SemesterSubjectScoreList)
                                {

                                    if (smScore == null)
                                        continue;

                                    if (smScore.Detail.GetAttribute("不計學分") == "是")
                                        continue;                                  

                                    // 累計應得學分
                                    Global._StudCreditDict[stud.StudentID].shouldGetTotalCredit += smScore.Credit;

                                    // 累計實得學分
                                    if (smScore.Pass)
                                        Global._StudCreditDict[stud.StudentID].gotTotalCredit += smScore.Credit;

                                    if (smScore.SchoolYear == sSchoolYear && smScore.Semester == sSemester)
                                    {
                                        // 建立學生學期成績索引
                                        if (!SemesterSubjectScoreInfoDict.ContainsKey(stud.StudentID))
                                            SemesterSubjectScoreInfoDict.Add(stud.StudentID, new Dictionary<string, SemesterSubjectScoreInfo>());

                                        string subjKey = smScore.Subject + smScore.Level;
                                        if (!SemesterSubjectScoreInfoDict[stud.StudentID].ContainsKey(subjKey))
                                            SemesterSubjectScoreInfoDict[stud.StudentID].Add(subjKey, smScore);

                                        // 學期應得學分
                                        Global._StudCreditDict[stud.StudentID].shouldGetCredit += smScore.Credit;

                                        // 學期實得學分
                                        if (smScore.Pass)
                                            Global._StudCreditDict[stud.StudentID].gotCredit += smScore.Credit;

                                        if (!studentSubjectScoreDict.ContainsKey(stud.StudentID))
                                            studentSubjectScoreDict.Add(stud.StudentID, new Dictionary<string, Dictionary<string, SemesterSubjectScoreInfo>>());

                                        if (!studentSubjectScoreDict[stud.StudentID].ContainsKey(smScore.Subject))
                                            studentSubjectScoreDict[stud.StudentID].Add(smScore.Subject, new Dictionary<string, SemesterSubjectScoreInfo>());

                                        // 科目名稱+級別是key
                                        string key11 = smScore.Subject + "^^^" + smScore.Level;
                                        if (!studentSubjectScoreDict[stud.StudentID][smScore.Subject].ContainsKey(key11))
                                            studentSubjectScoreDict[stud.StudentID][smScore.Subject].Add(key11, smScore);
                                    }
                                }
                            }
                        }
                        catch (Exception exception)
                        {
                            exc = exception;
                        }
                        finally
                        { }
                        #endregion



                        bkw.ReportProgress(70);
                        progressCount = 0;
                        List<string> tagName1List = new List<string>();
                        List<string> tagName2List = new List<string>();

                        #region 填入資料表
                        foreach (ClassRecord classRec in selectedClasses)
                        {
                            DataRow row = table.NewRow();
                            List<string> tag1List = new List<string>();
                            List<string> tag2List = new List<string>();
                            Dictionary<string, Dictionary<string, List<string>>> classSubjects = new Dictionary<string, Dictionary<string, List<string>>>();
                            string gradeYear = classRec.GradeYear;
                            #region 基本資料
                            row["學校名稱"] = SmartSchool.Customization.Data.SystemInformation.SchoolChineseName;
                            row["學校地址"] = SmartSchool.Customization.Data.SystemInformation.Address;
                            row["學校電話"] = SmartSchool.Customization.Data.SystemInformation.Telephone;
                            row["科別名稱"] = classRec.Department;


                            row["學年度"] = conf.SchoolYear;
                            row["學期"] = conf.Semester;
                            row["班級科別名稱"] = classRec.Department;
                            row["班級"] = classRec.ClassName;


                            row["類別排名1"] = "";
                            row["類別排名2"] = "";
                            row["班導師"] = classRec.RefTeacher == null ? "" : classRec.RefTeacher.TeacherName;
                            int ClassStuNum = 0;
                            tagName1List.Clear();
                            tagName2List.Clear();

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


                                #region 整理班級中列印科目
                                if (studentSubjectScoreDict.ContainsKey(stuRec.StudentID))
                                {
                                    foreach (var subjectName in studentSubjectScoreDict[studentID].Keys)
                                    {
                                        foreach (SemesterSubjectScoreInfo subjScore in studentSubjectScoreDict[studentID][subjectName].Values)
                                        {
                                            // 過濾畫面上選的學年度學期
                                            if (subjScore.SchoolYear == sSchoolYear && subjScore.Semester == sSemester)
                                            {
                                                string subjectLevel = subjScore.Level;
                                                string credit = "" + subjScore.Credit;
                                                if (conf.PrintSubjectList.Contains(subjectName))
                                                {
                                                    if (!classSubjects.ContainsKey(subjectName))
                                                        classSubjects.Add(subjectName, new Dictionary<string, List<string>>());
                                                    if (!classSubjects[subjectName].ContainsKey(subjectLevel))
                                                        classSubjects[subjectName].Add(subjectLevel, new List<string>());
                                                    if (!classSubjects[subjectName][subjectLevel].Contains(credit))
                                                        classSubjects[subjectName][subjectLevel].Add(credit);
                                                }
                                            }
                                            
                                        }
                                    }
                                }
                                #endregion
                            }
                            #endregion
                            #region 各科成績資料
                            #region 整理列印順序
                            List<string> subjectNameList = new List<string>(classSubjects.Keys);
                            subjectNameList.Sort(new StringComparer("國文"
                                            , "英文"
                                            , "數學"
                                            , "理化"
                                            , "生物"
                                            , "社會"
                                            , "物理"
                                            , "化學"
                                            , "歷史"
                                            , "地理"
                                            , "公民"));
                            #endregion
                            int subjectIndex = 1;
                            // 學期科目與定期評量
                            foreach (string subjectName in subjectNameList)
                            {
                                foreach (string subjectLevel in classSubjects[subjectName].Keys)
                                {
                                    decimal level;
                                    decimal? subjectNumber = null;
                                    subjectNumber = decimal.TryParse(subjectLevel, out level) ? (decimal?)level : null;
                                    if (subjectIndex <= conf.SubjectLimit)
                                    {
                                        row["科目名稱" + subjectIndex] = subjectName + GetNumber(subjectNumber);
                                        row["學分數" + subjectIndex] = "";
                                        foreach (string credit in classSubjects[subjectName][subjectLevel])
                                        {
                                            row["學分數" + subjectIndex] += (("" + row["學分數" + subjectIndex]) == "" ? "" : ",") + credit;
                                        }
                                        // 檢查畫面上定期評量列印科目
                                        if (conf.PrintSubjectList.Contains(subjectName))
                                        {
                                            ClassStuNum = 0;
                                            foreach (StudentRecord stuRec in classRec.Students)
                                            {
                                                ClassStuNum++;
                                                if (ClassStuNum > conf.StudentLimit)
                                                    break;
                                                string studentID = stuRec.StudentID;
                                                string key1 = subjectName + subjectNumber;
                                                if (SemesterSubjectScoreInfoDict.ContainsKey(studentID))
                                                {
                                                    if (SemesterSubjectScoreInfoDict[studentID].ContainsKey(key1))
                                                    {
                                                        row["科目成績" + ClassStuNum + "-" + subjectIndex] = SemesterSubjectScoreInfoDict[studentID][key1].Score;

                                                        if (SemsScoreRankMatrixDataDict.ContainsKey(studentID))
                                                        {
                                                            string sk1 = "";

                                                            foreach (string rt in rankTypeMapDict.Keys)
                                                            {
                                                                sk1 = "學期/科目成績_" + subjectName + "_" + rt;
                                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(sk1))
                                                                {
                                                                    if (rt.Contains("類"))
                                                                    {
                                                                        if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("rank"))
                                                                            row[rankTypeMapDict[rt] + "排名" + ClassStuNum + "-" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][sk1]["rank"];
                                                                        if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("matrix_count"))
                                                                            row[rankTypeMapDict[rt] + "排名母數" + ClassStuNum + "-" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][sk1]["matrix_count"];
                                                                    }
                                                                    else
                                                                    {
                                                                        if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("rank"))
                                                                            row[rankTypeMapDict[rt] + ClassStuNum + "-" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][sk1]["rank"];
                                                                        if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("matrix_count"))
                                                                            row[rankTypeMapDict[rt] + "母數" + ClassStuNum + "-" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][sk1]["matrix_count"];
                                                                    }

                                                                    foreach (string rr in r2ListOld)
                                                                    {
                                                                        if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey(rr))
                                                                        {
                                                                            string rkey = "科目" + rankTypeMapDict[rt] + "_" + rr + "_" + ClassStuNum + "-" + subjectIndex;
                                                                            row[rkey] = SemsScoreRankMatrixDataDict[studentID][sk1][rr];
                                                                        }
                                                                    }



                                                                }


                                                                sk1 = "學期/科目成績(原始)_" + subjectName + "_" + rt;
                                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(sk1))
                                                                {
                                                                    if (rt.Contains("類"))
                                                                    {
                                                                        if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("rank"))
                                                                            row[rankTypeMapDict[rt] + "排名(原始)" + ClassStuNum + "-" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][sk1]["rank"];
                                                                        if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("matrix_count"))
                                                                            row[rankTypeMapDict[rt] + "排名(原始)母數" + ClassStuNum + "-" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][sk1]["matrix_count"];
                                                                    }
                                                                    else
                                                                    {
                                                                        if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("rank"))
                                                                            row[rankTypeMapDict[rt] + "(原始)" + ClassStuNum + "-" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][sk1]["rank"];
                                                                        if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("matrix_count"))
                                                                            row[rankTypeMapDict[rt] + "(原始)母數" + ClassStuNum + "-" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][sk1]["matrix_count"];
                                                                    }

                                                                    foreach (string rr in r2ListOld)
                                                                    {
                                                                        if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey(rr))
                                                                        {
                                                                            string rkey = "科目(原始)" + rankTypeMapDict[rt] + "_" + rr + "_" + ClassStuNum + "-" + subjectIndex;
                                                                            row[rkey] = SemsScoreRankMatrixDataDict[studentID][sk1][rr];
                                                                        }
                                                                    }



                                                                }

                                                            }


                                                        }
                                                    }
                                                }

                                            }


                                            // dev..
                                            if (SemsScoreRankMatrixDataClassDict.ContainsKey(classRec.ClassID))
                                            {
                                                foreach (string s2 in rankTypeList)
                                                {
                                                    string s1key = "學期/科目成績_" + subjectName + "_" + s2;
                                                    if (SemsScoreRankMatrixDataClassDict[classRec.ClassID].ContainsKey(s1key))
                                                    {
                                                        foreach (string r in r2ListClass)
                                                        {
                                                            if (SemsScoreRankMatrixDataClassDict[classRec.ClassID][s1key].ContainsKey(r))
                                                            {
                                                                string sbkey = s2 + "科目_" + r + "_" + subjectIndex;
                                                                row[sbkey] = SemsScoreRankMatrixDataClassDict[classRec.ClassID][s1key][r];
                                                            }
                                                        }
                                                    }

                                                    s1key = "學期/科目成績(原始)_" + subjectName + "_" + s2;
                                                    if (SemsScoreRankMatrixDataClassDict[classRec.ClassID].ContainsKey(s1key))
                                                    {
                                                        foreach (string r in r2ListClass)
                                                        {
                                                            if (SemsScoreRankMatrixDataClassDict[classRec.ClassID][s1key].ContainsKey(r))
                                                            {
                                                                string sbkey = s2 + "科目(原始)_" + r + "_" + subjectIndex;
                                                                row[sbkey] = SemsScoreRankMatrixDataClassDict[classRec.ClassID][s1key][r];
                                                            }
                                                        }
                                                    }
                                                }
                                            }


                                            subjectIndex++;
                                        }
                                        else
                                        {
                                            //重要!!發現資料在樣板中印不下時一定要記錄起來，否則使用者自己不會去發現的
                                            if (!overflowRecords.Contains(classRec))
                                                overflowRecords.Add(classRec);
                                        }
                                    }
                                }
                                #endregion


                            }


                            #region 加權平均
                            ClassStuNum = 0;
                            foreach (StudentRecord stuRec in classRec.Students)
                            {
                                ClassStuNum++;
                                if (ClassStuNum > conf.StudentLimit)
                                    break;
                                string studentID = stuRec.StudentID;

                                // 處理學生學分
                                if (Global._StudCreditDict.ContainsKey(studentID))
                                {
                                    row["應得學分" + ClassStuNum] = Global._StudCreditDict[studentID].shouldGetCredit;
                                    row["實得學分" + ClassStuNum] = Global._StudCreditDict[studentID].gotCredit;
                                    row["應得學分累計" + ClassStuNum] = Global._StudCreditDict[studentID].shouldGetTotalCredit;
                                    row["實得學分累計" + ClassStuNum] = Global._StudCreditDict[studentID].gotTotalCredit;
                                }

                                if (Global._TempStudentSemesScoreDict.ContainsKey(studentID))
                                {
                                    if (Global._TempStudentSemesScoreDict[studentID].ContainsKey("學業"))
                                    {
                                        row["加權平均" + ClassStuNum] = Global._TempStudentSemesScoreDict[studentID]["學業"];
                                    }

                                    if (Global._TempStudentSemesScoreDict[studentID].ContainsKey("學業(原始)"))
                                    {
                                        row["加權平均(原始)" + ClassStuNum] = Global._TempStudentSemesScoreDict[studentID]["學業(原始)"];
                                    }
                                }



                                if (SemsScoreRankMatrixDataDict.ContainsKey(studentID))
                                {

                                    foreach (string rt in rankTypeMapDict.Keys)
                                    {
                                        string sk1 = "";
                                        sk1 = "學期/分項成績_學業_" + rt;
                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(sk1))
                                        {
                                            if (rt.Contains("類"))
                                            {
                                                // 類1,類2名稱
                                                if (DAO.QueryData.StudentIDTag1Dict.ContainsKey(studentID))
                                                {
                                                    row["學生類別排名1名稱" + ClassStuNum] = DAO.QueryData.StudentIDTag1Dict[studentID];
                                                }

                                                if (DAO.QueryData.StudentIDTag2Dict.ContainsKey(studentID))
                                                {
                                                    row["學生類別排名2名稱" + ClassStuNum] = DAO.QueryData.StudentIDTag2Dict[studentID];
                                                }

                                                if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("rank"))
                                                    row[rankTypeMapDict[rt] + "加權平均排名" + ClassStuNum] = SemsScoreRankMatrixDataDict[studentID][sk1]["rank"];
                                                if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("matrix_count"))
                                                    row[rankTypeMapDict[rt] + "加權平均排名母數" + ClassStuNum] = SemsScoreRankMatrixDataDict[studentID][sk1]["matrix_count"];
                                            }
                                            else
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("rank"))
                                                    row["加權平均" + rankTypeMapDict[rt] + ClassStuNum] = SemsScoreRankMatrixDataDict[studentID][sk1]["rank"];
                                                if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("matrix_count"))
                                                    row["加權平均" + rankTypeMapDict[rt] + "母數" + ClassStuNum] = SemsScoreRankMatrixDataDict[studentID][sk1]["matrix_count"];
                                            }

                                            foreach (string rr in r2ListOld)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey(rr))
                                                {
                                                    string rkey = "學業" + rankTypeMapDict[rt] + "_" + rr + "_" + ClassStuNum;
                                                    row[rkey] = SemsScoreRankMatrixDataDict[studentID][sk1][rr];
                                                }
                                            }

                                        }

                                        sk1 = "學期/分項成績(原始)_學業_" + rt;

                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(sk1))
                                        {
                                            if (rt.Contains("類"))
                                            {

                                                if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("rank"))
                                                    row[rankTypeMapDict[rt] + "加權平均排名(原始)" + ClassStuNum] = SemsScoreRankMatrixDataDict[studentID][sk1]["rank"];
                                                if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("matrix_count"))
                                                    row[rankTypeMapDict[rt] + "加權平均排名(原始)母數" + ClassStuNum] = SemsScoreRankMatrixDataDict[studentID][sk1]["matrix_count"];
                                            }
                                            else
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(sk1))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("rank"))
                                                        row["加權平均" + rankTypeMapDict[rt] + "(原始)" + ClassStuNum] = SemsScoreRankMatrixDataDict[studentID][sk1]["rank"];
                                                    if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey("matrix_count"))
                                                        row["加權平均" + rankTypeMapDict[rt] + "(原始)母數" + ClassStuNum] = SemsScoreRankMatrixDataDict[studentID][sk1]["matrix_count"];
                                                }
                                            }

                                            foreach (string rr in r2ListOld)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][sk1].ContainsKey(rr))
                                                {
                                                    string rkey = "學業(原始)" + rankTypeMapDict[rt] + "_" + rr + "_" + ClassStuNum;
                                                    row[rkey] = SemsScoreRankMatrixDataDict[studentID][sk1][rr];
                                                }
                                            }

                                        }

                                    }

                                }

                            }
                            #endregion

                            // dev..
                            // 班級五標、組距(分項成績)
                            if (SemsScoreRankMatrixDataClassDict.ContainsKey(classRec.ClassID))
                            {

                                if (DAO.QueryData.StudentClassIDTag1Dict.ContainsKey(classRec.ClassID))
                                {
                                    row["類別排名1"] = DAO.QueryData.StudentClassIDTag1Dict[classRec.ClassID];
                                }

                                if (DAO.QueryData.StudentClassIDTag2Dict.ContainsKey(classRec.ClassID))
                                {
                                    row["類別排名2"] = DAO.QueryData.StudentClassIDTag2Dict[classRec.ClassID];
                                }

                                foreach (string s1 in itemNameList)
                                {
                                    foreach (string s2 in rankTypeList)
                                    {
                                        string ssk1 = "學期/分項成績_" + s1 + "_" + s2;
                                        if (SemsScoreRankMatrixDataClassDict[classRec.ClassID].ContainsKey(ssk1))
                                        {
                                            foreach (string r in r2ListClass)
                                            {
                                                if (SemsScoreRankMatrixDataClassDict[classRec.ClassID][ssk1].ContainsKey(r))
                                                {
                                                    string rr = s1 + "_" + s2 + "_" + r;
                                                    row[rr] = SemsScoreRankMatrixDataClassDict[classRec.ClassID][ssk1][r];
                                                }
                                            }
                                        }

                                        ssk1 = "學期/分項成績(原始)_" + s1 + "_" + s2;
                                        if (SemsScoreRankMatrixDataClassDict[classRec.ClassID].ContainsKey(ssk1))
                                        {
                                            foreach (string r in r2ListClass)
                                            {
                                                if (SemsScoreRankMatrixDataClassDict[classRec.ClassID][ssk1].ContainsKey(r))
                                                {
                                                    string rr = s1 + "(原始)_" + s2 + "_" + r;
                                                    row[rr] = SemsScoreRankMatrixDataClassDict[classRec.ClassID][ssk1][r];
                                                }
                                            }
                                        }
                                    }
                                }
                            }



                            table.Rows.Add(row);
                            progressCount++;
                            bkw.ReportProgress(70 + progressCount * 20 / selectedClasses.Count);

                        }
                        #endregion

                        //// debug 用
                        //table.TableName = "debug";
                        //table.WriteXml(Application.StartupPath + "\\debug.xml");

                        bkw.ReportProgress(90);
                        document = conf.Template;
                        document.MailMerge.Execute(table);

                    }
                    catch (Exception exception)
                    {
                        exc = exception;
                    }
                };
                bkw.RunWorkerAsync();
            }
        }

        internal static void CreateFieldTemplate()
        {
            #region 產生欄位表

            Aspose.Words.Document doc = new Aspose.Words.Document(new System.IO.MemoryStream(Properties.Resources.Template));
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
            int maxSubjectNum = 15;
            int maxStuNum = 60;

            builder.Font.Size = 8;
            #region 基本欄位
            builder.Writeln("基本欄位");
            builder.StartTable();
            foreach (string field in new string[] { "學年度", "學期", "學校名稱", "學校地址", "學校電話", "科別名稱", "班級", "班導師", "類別排名1", "類別排名2" })
            {
                builder.InsertCell();
                builder.Write(field);
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + field + " \\* MERGEFORMAT ", "«" + field + "»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 科目成績
            builder.Writeln("科目成績");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("科目名稱");
            for (int i = 1; i <= maxSubjectNum; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目名稱" + i + " \\* MERGEFORMAT ", "«科目名稱" + i + "»");
            }
            builder.EndRow();

            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("學分數");
            for (int i = 1; i <= maxSubjectNum; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學分數" + i + " \\* MERGEFORMAT ", "«C" + i + "»");
            }
            builder.EndRow();

            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 科目成績" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«S" + i + "»");
                }
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 科目成績排名
            foreach (string key in new string[] { "班", "科", "全校", "類別1", "類別2" })
            {
                builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
                builder.Writeln("科目成績" + key + "排名");
                builder.StartTable();
                builder.InsertCell();
                builder.InsertCell();
                builder.InsertCell();
                builder.Write("科目名稱");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 科目名稱" + i + " \\* MERGEFORMAT ", "«科目名稱" + i + "»");
                }
                builder.EndRow();

                builder.InsertCell();
                builder.InsertCell();
                builder.InsertCell();
                builder.Write("學分數");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 學分數" + i + " \\* MERGEFORMAT ", "«C" + i + "»");
                }
                builder.EndRow();

                for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 學號" + stuIndex + "\\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 座號" + stuIndex + "\\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 姓名" + stuIndex + "\\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                    for (int i = 1; i <= maxSubjectNum; i++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + key + "排名" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«RS»");
                        builder.InsertField("MERGEFIELD " + key + "排名母數" + stuIndex + "-" + i + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
            }
            #endregion

            #region 科目成績排名(原始)
            foreach (string key in new string[] { "班", "科", "全校", "類別1", "類別2" })
            {
                builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
                builder.Writeln("科目成績" + key + "排名(原始)");
                builder.StartTable();
                builder.InsertCell();
                builder.InsertCell();
                builder.InsertCell();
                builder.Write("科目名稱");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 科目名稱" + i + " \\* MERGEFORMAT ", "«科目名稱" + i + "»");
                }
                builder.EndRow();

                builder.InsertCell();
                builder.InsertCell();
                builder.InsertCell();
                builder.Write("學分數");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 學分數" + i + " \\* MERGEFORMAT ", "«C" + i + "»");
                }
                builder.EndRow();

                for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 學號" + stuIndex + "\\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 座號" + stuIndex + "\\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 姓名" + stuIndex + "\\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                    for (int i = 1; i <= maxSubjectNum; i++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + key + "排名(原始)" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«RS»");
                        builder.InsertField("MERGEFIELD " + key + "排名(原始)母數" + stuIndex + "-" + i + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
            }
            #endregion

            #region 加權總分、加權平均及排名
            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("加權平均及排名");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();

            builder.Write("加權平均");
            builder.InsertCell();
            builder.Write("加權平均班排名");
            builder.InsertCell();
            builder.Write("加權平均科排名");
            builder.InsertCell();
            builder.Write("加權平均校排名");
            builder.EndRow();
            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均" + stuIndex + " \\* MERGEFORMAT ", "«平均»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均班排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 加權平均班排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均科排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 加權平均科排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均全校排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 加權平均全校排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 加權總分、加權平均及排名(原始)
            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("加權平均及排名(原始)");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();

            builder.Write("加權平均(原始)");
            builder.InsertCell();
            builder.Write("加權平均班排名(原始)");
            builder.InsertCell();
            builder.Write("加權平均科排名(原始)");
            builder.InsertCell();
            builder.Write("加權平均校排名(原始)");
            builder.EndRow();
            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均(原始)" + stuIndex + " \\* MERGEFORMAT ", "«平均»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均班排名(原始)" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 加權平均班排名(原始)母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均科排名(原始)" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 加權平均科排名(原始)母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均全校排名(原始)" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 加權平均全校排名(原始)母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion



            #region 應得與實得學分

            builder.Writeln("應得與實得學分");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("應得學分");
            builder.InsertCell();
            builder.Write("實得學分");
            builder.InsertCell();
            builder.Write("應得學分累計");
            builder.InsertCell();
            builder.Write("實得學分累計");
            builder.EndRow();
            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

                builder.InsertCell();
                builder.InsertField("MERGEFIELD 應得學分" + stuIndex + " \\* MERGEFORMAT ", "«SC»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 實得學分" + stuIndex + " \\* MERGEFORMAT ", "«GC»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 應得學分累計" + stuIndex + " \\* MERGEFORMAT ", "«ST»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 實得學分累計" + stuIndex + " \\* MERGEFORMAT ", "«GT»");

                builder.EndRow();
            }
            builder.EndTable();

            #endregion

            #region 類別1排名
            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("類別1排名");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();

            //builder.Write("加權平均");
            //builder.InsertCell();
            builder.Write("加權平均排名");
            builder.EndRow();
            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

                //builder.InsertCell();
                //builder.InsertField("MERGEFIELD 類別1加權平均" + stuIndex + " \\* MERGEFORMAT ", "«類1加均»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別1加權平均排名" + stuIndex + " \\* MERGEFORMAT ", "«RP»");
                builder.InsertField("MERGEFIELD 類別1加權平均排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TP»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 類別2排名
            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("類別2排名");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();

            //builder.Write("加權平均");
            //builder.InsertCell();
            builder.Write("加權平均排名");
            builder.EndRow();
            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

                //builder.InsertCell();
                //builder.InsertField("MERGEFIELD 類別2加權平均" + stuIndex + " \\* MERGEFORMAT ", "«類2加均»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別2加權平均排名" + stuIndex + " \\* MERGEFORMAT ", "«RP»");
                builder.InsertField("MERGEFIELD 類別2加權平均排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TP»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            List<string> r1 = new List<string>();
            List<string> r2 = new List<string>();

            r1.Add("pr");
            r1.Add("percentile");
            r1.Add("avg_top_25");
            r1.Add("avg_top_50");
            r1.Add("avg");
            r1.Add("avg_bottom_50");
            r1.Add("avg_bottom_25");

            r2.Add("level_gte100");
            r2.Add("level_90");
            r2.Add("level_80");
            r2.Add("level_70");
            r2.Add("level_60");
            r2.Add("level_50");
            r2.Add("level_40");
            r2.Add("level_30");
            r2.Add("level_20");
            r2.Add("level_10");
            r2.Add("level_lt10");

            List<string> ta1 = new List<string>();
            ta1.Add("學業");
            ta1.Add("學業(原始)");

            List<string> ta2 = new List<string>();
            ta2.Add("班排名");
            ta2.Add("科排名");
            ta2.Add("全校排名");
            ta2.Add("類別1");
            ta2.Add("類別2");

            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);

            foreach (string s1 in ta1)
            {
                builder.Writeln();
                builder.Writeln(s1 + " 百分比、五標");

                builder.StartTable();

                builder.InsertCell();
                builder.Write("名稱_序號");
                builder.InsertCell();
                builder.Write("PR");
                builder.InsertCell();
                builder.Write("百分比");
                builder.InsertCell();
                builder.Write("頂標");
                builder.InsertCell();
                builder.Write("高標");
                builder.InsertCell();
                builder.Write("均標");
                builder.InsertCell();
                builder.Write("低標");
                builder.InsertCell();
                builder.Write("底標");
                builder.EndRow();
                foreach (string s2 in ta2)
                {

                    for (int i = 1; i <= maxStuNum; i++)
                    {

                        builder.InsertCell();
                        builder.Write(s2 + "_" + i);
                        foreach (string s3 in r1)
                        {
                            builder.InsertCell();
                            string nn = s1 + s2 + "_" + s3 + "_" + i;
                            builder.InsertField("MERGEFIELD " + nn + " \\* MERGEFORMAT ", "«R»");
                        }
                        builder.EndRow();
                    }

                }
                builder.EndTable();

            }

            foreach (string s1 in ta1)
            {
                builder.Writeln();
                builder.Writeln(s1 + " 組距");

                builder.StartTable();

                builder.InsertCell();
                builder.Write("名稱_序號");
                builder.InsertCell();
                builder.Write("100以上");
                builder.InsertCell();
                builder.Write("90以上小於100");
                builder.InsertCell();
                builder.Write("80以上小於90");
                builder.InsertCell();
                builder.Write("70以上小於80");
                builder.InsertCell();
                builder.Write("60以上小於70");
                builder.InsertCell();
                builder.Write("50以上小於60");
                builder.InsertCell();
                builder.Write("40以上小於50");
                builder.InsertCell();
                builder.Write("30以上小於40");
                builder.InsertCell();
                builder.Write("20以上小於30");
                builder.InsertCell();
                builder.Write("10以上小於20");
                builder.InsertCell();
                builder.Write("10以下");
                builder.EndRow();
                foreach (string s2 in ta2)
                {
                    for (int i = 1; i <= maxStuNum; i++)
                    {

                        builder.InsertCell();
                        builder.Write(s2 + "_" + i);
                        foreach (string s3 in r2)
                        {
                            builder.InsertCell();
                            string nn = s1 + s2 + "_" + s3 + "_" + i;
                            builder.InsertField("MERGEFIELD " + nn + " \\* MERGEFORMAT ", "«R»");
                        }
                        builder.EndRow();

                    }
                }
                builder.EndTable();
                builder.Writeln();

            }



            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            #region 學業分項 百分比、五標、組距
            List<string> t1 = new List<string>();
            t1.Add("學業");
            t1.Add("學業(原始)");
            t1.Add("實習科目");
            t1.Add("專業科目(原始)");
            t1.Add("專業科目");
            t1.Add("實習科目(原始)");

            List<string> t2 = new List<string>();
            t2.Add("班排名");
            t2.Add("科排名");
            t2.Add("年排名");
            t2.Add("類別1排名");
            t2.Add("類別2排名");




            foreach (string s1 in t1)
            {
                builder.Writeln();
                builder.Writeln(s1 + " 百分比、五標");

                builder.StartTable();
                builder.InsertCell();
                builder.Write("分類");
                builder.InsertCell();
                builder.Write("頂標");
                builder.InsertCell();
                builder.Write("高標");
                builder.InsertCell();
                builder.Write("均標");
                builder.InsertCell();
                builder.Write("低標");
                builder.InsertCell();
                builder.Write("底標");
                builder.EndRow();
                foreach (string s2 in t2)
                {
                    builder.InsertCell();
                    builder.Write(s2);
                    foreach (string s3 in r1)
                    {
                        if (s3 == "pr" || s3 == "percentile")
                            continue;

                        builder.InsertCell();
                        string nn = s1 + "_" + s2 + "_" + s3;
                        builder.InsertField("MERGEFIELD " + nn + " \\* MERGEFORMAT ", "«R»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();

            }


            foreach (string s1 in t1)
            {
                builder.Writeln();
                builder.Writeln(s1 + " 組距");

                builder.StartTable();

                builder.InsertCell();
                builder.Write("分類");
                builder.InsertCell();
                builder.Write("100以上");
                builder.InsertCell();
                builder.Write("90以上小於100");
                builder.InsertCell();
                builder.Write("80以上小於90");
                builder.InsertCell();
                builder.Write("70以上小於80");
                builder.InsertCell();
                builder.Write("60以上小於70");
                builder.InsertCell();
                builder.Write("50以上小於60");
                builder.InsertCell();
                builder.Write("40以上小於50");
                builder.InsertCell();
                builder.Write("30以上小於40");
                builder.InsertCell();
                builder.Write("20以上小於30");
                builder.InsertCell();
                builder.Write("10以上小於20");
                builder.InsertCell();
                builder.Write("10以下");
                builder.EndRow();
                foreach (string s2 in t2)
                {
                    builder.InsertCell();
                    builder.Write(s2);
                    foreach (string s3 in r2)
                    {
                        builder.InsertCell();
                        string nn = s1 + "_" + s2 + "_" + s3;
                        builder.InsertField("MERGEFIELD " + nn + " \\* MERGEFORMAT ", "«R»");
                    }
                    builder.EndRow();
                }
                builder.EndTable();
                builder.Writeln();

            }
            #endregion


            #region 學生學業 類別1,2名稱 

            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("學生類別排名1名稱,類別排名2名稱");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("類別排名1名稱");
            builder.InsertCell();
            builder.Write("類別排名2名稱");         
            builder.EndRow();

            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學生類別排名1名稱" + stuIndex + " \\* MERGEFORMAT ", "«類1_" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學生類別排名2名稱" + stuIndex + " \\* MERGEFORMAT ", "«類2_" + stuIndex + "»");

                builder.EndRow();
            }
            builder.EndTable();

            #endregion




            #endregion

            #region 儲存檔案
            string inputReportName = "班級學期成績單(固定排名)合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

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
                doc.Save(path, Aspose.Words.SaveFormat.Doc);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        doc.Save(path, Aspose.Words.SaveFormat.Doc);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
        }

    }
}
