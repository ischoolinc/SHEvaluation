﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Threading;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using SmartSchool;
using FISCA.Permission;
using 班級定期評量成績單_固定排名.Service;
using 班級定期評量成績單_固定排名.Model;
using FISCA.Data;

namespace 班級定期評量成績單_固定排名
{

    /// <summary>
    ///  此功能直接重原始定期評量成績單(及時排名)修改來。
    /// </summary>
    public class Program
    {
        // 四捨五入位數

        public static int AvgRd = 2;
        [FISCA.MainMethod]
        public static void Main()
        {

            var btn = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["班級定期評量成績單(固定排名)"];

            RoleAclSource.Instance["班級"]["功能按鈕"].Add(new RibbonFeature("41B01543-4815-4110-B4A7-FAFA3FCBED14", "班級定期評量成績單(固定排名)"));

            K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate
            {
                btn.Enable = K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0 &&
                FISCA.Permission.UserAcl.Current["41B01543-4815-4110-B4A7-FAFA3FCBED14"].Executable;
            };
            btn.Click += new EventHandler(Program_Click);
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


        // 按下列印後
        private static void Program_Click(object sender_, EventArgs e_)
        {
            AccessHelper accessHelper = new AccessHelper();
            // 基本資料放在這裡 
            List<ClassRecord> selectedClasses = accessHelper.ClassHelper.GetSelectedClass();
            List<StudentRecord> selectedStudents = new List<StudentRecord>();
            List<string> selectGrades = new List<string>();//選取之班級的年級
            List<string> selectClass = selectedClasses.Select(x => x.ClassID).ToList();
            foreach (ClassRecord classRecord in selectedClasses)
            {
                if (!selectGrades.Contains(classRecord.GradeYear))
                    selectGrades.Add(classRecord.GradeYear);
            }


            ConfigForm form = new ConfigForm(selectClass); ;
            if (form.ShowDialog() == DialogResult.OK)
            {
                Configure conf = form.Configure; // 列印設定相關變數(學年度 學期 )


                Dictionary<string, Dictionary<string, FixRankIntervalInfo>> IntervalInfos;// 組距 by 班級 
                Dictionary<string, Dictionary<string, StudentFixedRankInfo>> dicFixedRankData;// 組距 by 班級 學生 

                List<ClassRecord> overflowRecords = new List<ClassRecord>();
                Exception exc = null;
                //取得列印設定
                // 1.處理類別排名 
                // 1.0 取得現在班級下有哪一些類別 及類別(類別1 類別2)下有 哪些項目( 自然組 社會組)
                Dictionary<string, Dictionary<string, List<string>>> FixRankTags = Utility.GetTagInfoFromRankMatrixByClass(conf.SchoolYear, conf.Semester, conf.ExamRecord.ID, selectedClasses);

                //  

                // 把所選取 班級下的學生存起來
                foreach (ClassRecord classRec in selectedClasses)
                {
                    foreach (StudentRecord stuRec in classRec.Students)
                    {
                        if (!selectedStudents.Contains(stuRec))
                            selectedStudents.Add(stuRec);
                    }
                }

                // 建立合併欄位總表
                DataTable table = Utility.CreateMergeDataTabel(conf, selectedClasses);

                //宣告產生的報表
                Aspose.Words.Document document = new Aspose.Words.Document();
                //用一個BackgroundWorker包起來
                System.ComponentModel.BackgroundWorker bkw = new System.ComponentModel.BackgroundWorker();
                bkw.WorkerReportsProgress = true;
                System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 班級定期評量成績單(固定排名)產生 S");
                bkw.ProgressChanged += delegate (object sender, System.ComponentModel.ProgressChangedEventArgs e)
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage(" 班級評量成績單產生中", e.ProgressPercentage);
                    System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 班級定期評量成績單(固定排名)產生 " + e.ProgressPercentage);
                };
                bkw.RunWorkerCompleted += delegate
                {
                    System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 班級定期評量成績單(固定排名)產生 E");
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
                        throw new Exception("產生班級定期評量成績單(固定排名)發生錯誤", exc);
                    }
                    #region 儲存檔案
                    string inputReportName = "班級評量成績單";
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
                    catch
                    {
                        System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                        sd.Title = "另存新檔";
                        sd.FileName = reportName + ".doc";
                        sd.Filter = "Excel檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                        if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            try
                            {
                                document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);

                            }
                            catch (Exception ex)
                            {
                                FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }
                    #endregion
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("班級定期評量成績單(固定排名)產生完成。", 100);
                    if (overflowRecords.Count > 0)
                        MessageBox.Show(err);
                };
                bkw.DoWork += delegate (object sender, System.ComponentModel.DoWorkEventArgs e)
                {
                    #region 偷跑取得考試成績
                    Dictionary<string, Dictionary<string, Dictionary<string, ExamScoreInfo>>> studentExamSores = new Dictionary<string, Dictionary<string, Dictionary<string, ExamScoreInfo>>>();
                    Dictionary<string, Dictionary<string, ExamScoreInfo>> studentRefExamSores = new Dictionary<string, Dictionary<string, ExamScoreInfo>>();
                    ManualResetEvent scoreReady = new ManualResetEvent(false);
                    ManualResetEvent elseReady = new ManualResetEvent(false);

                    bkw.ReportProgress(3);
                    new Thread(new ThreadStart(delegate
                        {
                            // 取得學生學期科目成績
                            int sSchoolYear, sSemester;
                            int.TryParse(conf.SchoolYear, out sSchoolYear);
                            int.TryParse(conf.Semester, out sSemester);

                            // 20191018  Jean 新增固定排名邏輯 需要在classID的回圈下

                            #region 整理學生定期評量成績
                            #region 篩選課程學年度、學期、科目取得有可能有需要的資料
                            List<CourseRecord> targetCourseList = new List<CourseRecord>();
                            try
                            {
                                foreach (var courseRecord in accessHelper.CourseHelper.GetAllCourse(sSchoolYear, sSemester))
                                {
                                    //用科目濾出可能有用到的課程
                                    if (conf.PrintSubjectList.Contains(courseRecord.Subject)
                                        || conf.TagRank1SubjectList.Contains(courseRecord.Subject)
                                        || conf.TagRank2SubjectList.Contains(courseRecord.Subject))

                                        //Jean
                                        targetCourseList.Add(courseRecord);
                                }
                            }
                            catch (Exception exception)
                            {
                                exc = exception;
                            }
                            #endregion`
                            try
                            {
                                if (conf.ExamRecord != null || conf.RefenceExamRecord != null)
                                {
                                    accessHelper.CourseHelper.FillExam(targetCourseList);
                                    var tcList = new List<CourseRecord>();
                                    var totalList = new List<CourseRecord>();
                                    foreach (var courseRec in targetCourseList)
                                    {
                                        //Jean

                                        if (conf.ExamRecord != null && courseRec.ExamList.Contains(conf.ExamRecord.Name))
                                        {
                                            tcList.Add(courseRec);
                                            totalList.Add(courseRec);
                                        }
                                        if (tcList.Count == 180)
                                        {
                                            accessHelper.CourseHelper.FillStudentAttend(tcList);
                                            accessHelper.CourseHelper.FillExamScore(tcList);
                                            tcList.Clear();
                                        }
                                    }
                                    accessHelper.CourseHelper.FillStudentAttend(tcList);
                                    accessHelper.CourseHelper.FillExamScore(tcList);

                                    //Jean
                                    foreach (var courseRecord in totalList)
                                    {
                                        #region 整理本次定期評量成績
                                        if (conf.ExamRecord != null && courseRecord.ExamList.Contains(conf.ExamRecord.Name))
                                        {
                                            foreach (var attendStudent in courseRecord.StudentAttendList)
                                            {
                                                if (!studentExamSores.ContainsKey(attendStudent.StudentID)) studentExamSores.Add(attendStudent.StudentID, new Dictionary<string, Dictionary<string, ExamScoreInfo>>());
                                                if (!studentExamSores[attendStudent.StudentID].ContainsKey(courseRecord.Subject)) studentExamSores[attendStudent.StudentID].Add(courseRecord.Subject, new Dictionary<string, ExamScoreInfo>());
                                                studentExamSores[attendStudent.StudentID][courseRecord.Subject].Add("" + attendStudent.CourseID, null);
                                            }
                                            foreach (var examScoreRec in courseRecord.ExamScoreList)//
                                            {
                                                if (examScoreRec.ExamName == conf.ExamRecord.Name)
                                                {
                                                    studentExamSores[examScoreRec.StudentID][courseRecord.Subject]["" + examScoreRec.CourseID] = examScoreRec;
                                                }
                                            }
                                        }
                                        #endregion
                                        #region 整理前次定期評量成績
                                        if (conf.RefenceExamRecord != null && courseRecord.ExamList.Contains(conf.RefenceExamRecord.Name))
                                        {
                                            foreach (var examScoreRec in courseRecord.ExamScoreList)
                                            {
                                                if (examScoreRec.ExamName == conf.RefenceExamRecord.Name)
                                                {
                                                    if (!studentRefExamSores.ContainsKey(examScoreRec.StudentID))
                                                        studentRefExamSores.Add(examScoreRec.StudentID, new Dictionary<string, ExamScoreInfo>());
                                                    studentRefExamSores[examScoreRec.StudentID].Add("" + examScoreRec.CourseID, examScoreRec);
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                }
                            }
                            catch (Exception exception)
                            {
                                exc = exception;
                            }
                            finally
                            {
                                scoreReady.Set();
                            }
                            #endregion
                            #region 整理學生學期、學年成績
                            try
                            {
                                accessHelper.StudentHelper.FillAttendance(selectedStudents);
                                accessHelper.StudentHelper.FillReward(selectedStudents);
                            }
                            catch (Exception exception)
                            {
                                exc = exception;
                            }
                            finally
                            {
                                elseReady.Set();
                            }
                            #endregion
                        })).Start();
                    #endregion
                    try
                    {
                        string key = "";
                        bkw.ReportProgress(10);
                        #region 整理同年級學生
                        // 整理所選班級有哪一些年級 

                        //整理選取學生的年級
                        Dictionary<string, List<StudentRecord>> gradeyearStudents = new Dictionary<string, List<StudentRecord>>();

                        //Jean 增加固定排名邏輯 
                        #region 固定排名
                        // 當年度學期班級定排名級距
                        // 取得固定排名下類1排名及類2 資訊

                        FixedRankIntervalService fixedRankService = new FixedRankIntervalService(conf.SchoolYear, conf.Semester, conf.ExamRecord.ID);

                        // 取得 本次固定排名 Tag 相關資訊 
                        List<string> classIDs = selectedClasses.Select(x => x.ClassID).ToList();
                        List<string> gradeYears = selectedClasses.Select(x => x.GradeYear).Distinct().ToList();
                        Dictionary<string, List<string>> tagRankInfoFromRankMatrix = Utility.GetTagInfoFromRankMatrix(conf.SchoolYear, conf.Semester, conf.ExamRecord.ID, selectedClasses);

                        IntervalInfos = fixedRankService.GetAllClassInterval(selectedClasses);
                        //當年度學生固定排名
                        FixedRankGetService rankService = new FixedRankGetService(conf.SchoolYear, conf.Semester, conf.ExamRecord.ID);
                        dicFixedRankData = rankService.GetFixedRankManyClass(classIDs);
                        // 取得班級下學生類別1 類別2(固定排名結算時)

                        Dictionary<string, Dictionary<string, string>> studentTags1FixRankSByClass = Utility.GetStudentTagSByClass(conf.SchoolYear, conf.Semester, conf.ExamRecord.ID, selectedClasses, "類別1排名");
                        Dictionary<string, Dictionary<string, string>> studentTags2FixRankSByClass = Utility.GetStudentTagSByClass(conf.SchoolYear, conf.Semester, conf.ExamRecord.ID, selectedClasses, "類別2排名");
                        Dictionary<string, Dictionary<string, List<string>>> subjextTagFromFixRankByClass = Utility.GetTagsSubjextFromRankMatrixByClass(conf.SchoolYear, conf.Semester, conf.ExamRecord.ID, selectedClasses, "類別1排名");


                        #endregion 不知道是不是用來計算 即時排名

                        foreach (ClassRecord classRec in selectedClasses) // 1.取得選擇班級
                        {
                            string grade = "";
                            grade = "" + classRec.GradeYear;
                            if (!gradeyearStudents.ContainsKey(grade))
                                gradeyearStudents.Add(grade, new List<StudentRecord>());
                            foreach (var studentRec in classRec.Students)
                            {
                                gradeyearStudents[grade].Add(studentRec); // 1.1 以年級下去抓學生資料
                            }
                        }

                        foreach (var classRec in accessHelper.ClassHelper.GetAllClass()) // 取得所班級
                        {
                            if (!selectedClasses.Contains(classRec) && gradeyearStudents.ContainsKey("" + classRec.GradeYear))
                            {
                                // 用班級去取出可能有相關的學生
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

                        #region 計算總分及各項目排名
                        Dictionary<string, string> studentTag1Group = new Dictionary<string, string>();
                        Dictionary<string, string> studentTag2Group = new Dictionary<string, string>();
                        // Dictionary<string, List<decimal>> ranks = new Dictionary<string, List<decimal>>();
                        // Dictionary<string, List<string>> rankStudents = new Dictionary<string, List<string>>();
                        Dictionary<string, decimal> studentPrintSubjectSum = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag1SubjectSum = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag2SubjectSum = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentPrintSubjectAvg = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag1SubjectAvg = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag2SubjectAvg = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentPrintSubjectSumW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag1SubjectSumW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag2SubjectSumW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentPrintSubjectAvgW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag1SubjectAvgW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> studentTag2SubjectAvgW = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> analytics = new Dictionary<string, decimal>();
                        int total = 0;
                        foreach (var gss in gradeyearStudents.Values)
                        {
                            total += gss.Count;
                        }
                        bkw.ReportProgress(40);
                        //foreach (string gradeyear in gradeyearStudents.Keys)
                        // 根據所選取得學生
                        foreach (ClassRecord classRecord in selectedClasses)
                        {
                            //找出全年級學生
                            //foreach (var studentRec in gradeyearStudents[gradeyear])
                            foreach (var studentRec in classRecord.Students)
                            {
                                string studentID = studentRec.StudentID;
                                bool rank = true;
                                string tag1ID = "";
                                string tag2ID = "";

                                #region 分析學生所屬類別
                                if (studentTags.ContainsKey(studentID))
                                {
                                    foreach (var tag in studentTags[studentID])
                                    {
                                        #region 判斷學生是否屬於不排名類別
                                        //if (conf.RankFilterTagList.Contains(tag.RefTagID))
                                        //{
                                        //    rank = false;
                                        //}
                                        #endregion
                                        #region 判斷學生在類別排名1中所屬的類別
                                        //if (tag1ID == "" && conf.TagRank1TagList.Contains(tag.RefTagID))
                                        //{
                                        //    tag1ID = tag.RefTagID;
                                        //    studentTag1Group.Add(studentID, tag1ID);
                                        //}
                                        #endregion
                                        #region 判斷學生在類別排名2中所屬的類別
                                        //if (tag2ID == "" && conf.TagRank2TagList.Contains(tag.RefTagID))
                                        //{
                                        //    tag2ID = tag.RefTagID;
                                        //    studentTag2Group.Add(studentID, tag2ID);
                                        //}
                                        #endregion
                                    }
                                }
                                #endregion
                                bool summaryRank = true;
                                bool tag1SummaryRank = true;
                                bool tag2SummaryRank = true;
                                if (studentExamSores.ContainsKey(studentID))
                                {
                                    decimal printSubjectSum = 0;
                                    int printSubjectCount = 0;
                                    decimal tag1SubjectSum = 0;
                                    int tag1SubjectCount = 0;
                                    decimal tag2SubjectSum = 0;
                                    int tag2SubjectCount = 0;
                                    decimal printSubjectSumW = 0;
                                    decimal printSubjectCreditSum = 0;
                                    decimal tag1SubjectSumW = 0;
                                    decimal tag1SubjectCreditSum = 0;
                                    decimal tag2SubjectSumW = 0;
                                    decimal tag2SubjectCreditSum = 0;
                                    foreach (var subjectName in studentExamSores[studentID].Keys)
                                    {
                                        if (conf.PrintSubjectList.Contains(subjectName))
                                        {
                                            #region 是列印科目
                                            foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                            {
                                                if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                {
                                                    printSubjectSum += sceTakeRecord.ExamScore;//計算總分
                                                    printSubjectCount++;
                                                    //計算加權總分
                                                    printSubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                    printSubjectCreditSum += sceTakeRecord.CreditDec();
                                                    if (rank && sceTakeRecord.Status == "一般")// 不在過濾名單且為一般生才做排名
                                                    {
                                                        if (sceTakeRecord.RefClass != null)
                                                        {
                                                            //各科目班排名
                                                            key = "班排名" + sceTakeRecord.RefClass.ClassID + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                            //  if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            //  if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            //  ranks[key].Add(sceTakeRecord.ExamScore);
                                                            //  rankStudents[key].Add(studentID);
                                                        }
                                                        if (sceTakeRecord.Department != "")
                                                        {
                                                            //各科目科排名
                                                            //   key = "科排名" + sceTakeRecord.Department + "^^^" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                            //    if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            //    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            //  ranks[key].Add(sceTakeRecord.ExamScore);
                                                            //    /rankStudents[key].Add(studentID);
                                                        }
                                                        //各科目全校排名
                                                        // key = "年排名" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                        //  if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                        //  if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                        //  ranks[key].Add(sceTakeRecord.ExamScore);
                                                        //  rankStudents[key].Add(studentID);
                                                    }
                                                }
                                                else
                                                {
                                                    summaryRank = false;
                                                }
                                            }
                                            #endregion
                                        }
                                        // if (tag1ID != "" && conf.TagRank1SubjectList.Contains(subjectName))



                                        if (subjextTagFromFixRankByClass.ContainsKey(studentRec.RefClass.ClassID) && subjextTagFromFixRankByClass[studentRec.RefClass.ClassID].ContainsKey("類別1排名"))
                                        {
                                            if (subjextTagFromFixRankByClass[studentRec.RefClass.ClassID]["類別1排名"].Contains(subjectName))
                                            {
                                                #region 有Tag1且是排名科目
                                                foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                                {
                                                    if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                    {
                                                        tag1SubjectSum += sceTakeRecord.ExamScore;//計算總分
                                                        tag1SubjectCount++;
                                                        //計算加權總分
                                                        tag1SubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                        tag1SubjectCreditSum += sceTakeRecord.CreditDec();
                                                        //各科目類別1排名
                                                        if (rank && sceTakeRecord.Status == "一般")//不在過濾名單且為一般生才做排名
                                                        {
                                                            if (conf.PrintSubjectList.Contains(subjectName))//是列印科目才算科目排名                                                
                                                            {
                                                                // key = "類別1排名" + tag1ID + "^^^" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                // if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                                // if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                                // ranks[key].Add(sceTakeRecord.ExamScore);
                                                                // rankStudents[key].Add(studentID);
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        tag1SummaryRank = false;
                                                    }
                                                }
                                                #endregion
                                            }
                                        }

                                        if (subjextTagFromFixRankByClass.ContainsKey(studentRec.RefClass.ClassID) && subjextTagFromFixRankByClass[studentRec.RefClass.ClassID].ContainsKey("類別2排名"))
                                        {
                                            if (subjextTagFromFixRankByClass[studentRec.RefClass.ClassID]["類別2排名"].Contains(subjectName))
                                            {
                                                #region 有Tag2且是排名科目
                                                foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                                {
                                                    if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                    {
                                                        tag2SubjectSum += sceTakeRecord.ExamScore;//計算總分
                                                        tag2SubjectCount++;
                                                        //計算加權總分
                                                        tag2SubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                        tag2SubjectCreditSum += sceTakeRecord.CreditDec();
                                                        //各科目類別2排名
                                                        if (rank && sceTakeRecord.Status == "一般")//不在過濾名單且為一般生才做排名
                                                        {
                                                            if (conf.PrintSubjectList.Contains(subjectName))//是列印科目才算科目排名                                                
                                                            {
                                                                //    key = "類別2排名" + tag2ID + "^^^" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                //  if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                                //  if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                                //  ranks[key].Add(sceTakeRecord.ExamScore);
                                                                //   rankStudents[key].Add(studentID);
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        tag2SummaryRank = false;
                                                    }
                                                }
                                                #endregion
                                            }
                                        }
                                    }
                                    if (printSubjectCount > 0)
                                    {
                                        #region 有列印科目處理加總成績
                                        //總分
                                        studentPrintSubjectSum.Add(studentID, printSubjectSum);
                                        //平均四捨五入至小數點第二位
                                        studentPrintSubjectAvg.Add(studentID, Math.Round(printSubjectSum / printSubjectCount, AvgRd, MidpointRounding.AwayFromZero));
                                        if (rank && studentRec.Status == "一般" && summaryRank == true)//不在過濾名單且沒有特殊成績狀況且為一般生才做排名
                                        {
                                            //總分班排名
                                            key = "總分班排名" + studentRec.RefClass.ClassID;
                                            // if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            // if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            //  ranks[key].Add(printSubjectSum);
                                            //  rankStudents[key].Add(studentID);
                                            //總分科排名
                                            //  key = "總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                            // if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            // if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            //   ranks[key].Add(printSubjectSum);
                                            //   rankStudents[key].Add(studentID);
                                            //總分全校排名
                                            //    key = "總分全校排名" + gradeyear;
                                            // if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            // if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            //  ranks[key].Add(printSubjectSum);
                                            //   rankStudents[key].Add(studentID);
                                            //平均班排名
                                            //  key = "平均班排名" + studentRec.RefClass.ClassID;
                                            // if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            //  if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            //  ranks[key].Add(Math.Round(printSubjectSum / printSubjectCount, AvgRd, MidpointRounding.AwayFromZero));
                                            //   rankStudents[key].Add(studentID);
                                            //平均科排名
                                            //  key = "平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                            //  if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            //   if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            //   ranks[key].Add(Math.Round(printSubjectSum / printSubjectCount, AvgRd, MidpointRounding.AwayFromZero));
                                            //   rankStudents[key].Add(studentID);
                                            //平均全校排名
                                            //  key = "平均年排名" + gradeyear;
                                            //  if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            //  if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            //  ranks[key].Add(Math.Round(printSubjectSum / printSubjectCount, AvgRd, MidpointRounding.AwayFromZero));
                                            //   rankStudents[key].Add(studentID);
                                        }
                                        #endregion
                                        if (printSubjectCreditSum > 0)
                                        {
                                            #region 有總學分數處理加總
                                            //加權總分
                                            studentPrintSubjectSumW.Add(studentID, printSubjectSumW);
                                            //加權平均四捨五入至小數點第二位
                                            studentPrintSubjectAvgW.Add(studentID, Math.Round(printSubjectSumW / printSubjectCreditSum, AvgRd, MidpointRounding.AwayFromZero));
                                            if (rank && studentRec.Status == "一般" && summaryRank == true)//不在過濾名單且為一般生才做排名
                                            {
                                                //加權總分班排名
                                                key = "加權總分班排名" + studentRec.RefClass.ClassID;
                                                //  if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                //   if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                //   ranks[key].Add(printSubjectSumW);
                                                //     rankStudents[key].Add(studentID);
                                                //加權總分科排名
                                                //   key = "加權總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                                //     if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                //   if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                //    ranks[key].Add(printSubjectSumW);
                                                //   rankStudents[key].Add(studentID);
                                                //加權總分全校排名
                                                //   key = "加權總分全校排名" + gradeyear;
                                                //  if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                //   if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                //   ranks[key].Add(printSubjectSumW);
                                                // rankStudents[key].Add(studentID);
                                                //加權平均班排名
                                                key = "加權平均班排名" + studentRec.RefClass.ClassID;
                                                // if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                //   if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                //   ranks[key].Add(Math.Round(printSubjectSumW / printSubjectCreditSum, AvgRd, MidpointRounding.AwayFromZero));
                                                //   rankStudents[key].Add(studentID);
                                                //加權平均科排名
                                                //  key = "加權平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                                //  if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                //    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                // ranks[key].Add(Math.Round(printSubjectSumW / printSubjectCreditSum, AvgRd, MidpointRounding.AwayFromZero));
                                                //    rankStudents[key].Add(studentID);
                                                //加權平均全校排名
                                                //   key = "加權平均年排名" + gradeyear;
                                                // if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                //   if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                // ranks[key].Add(Math.Round(printSubjectSumW / printSubjectCreditSum, AvgRd, MidpointRounding.AwayFromZero));
                                                //     rankStudents[key].Add(studentID);
                                            }
                                            #endregion
                                        }
                                    }
                                    //類別1總分平均排名
                                    if (tag1SubjectCount > 0)
                                    {
                                        //總分
                                        studentTag1SubjectSum.Add(studentID, tag1SubjectSum);
                                        //平均四捨五入至小數點第二位
                                        studentTag1SubjectAvg.Add(studentID, Math.Round(tag1SubjectSum / tag1SubjectCount, AvgRd, MidpointRounding.AwayFromZero));
                                        if (rank && studentRec.Status == "一般" && tag1SummaryRank == true)//不在過濾名單且為一般生才做排名
                                        {
                                            //  key = "類別1總分排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                            //   if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            //  if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            //  ranks[key].Add(tag1SubjectSum);
                                            //  rankStudents[key].Add(studentID);

                                            // key = "類別1平均排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                            // if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            // if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            //  ranks[key].Add(Math.Round(tag1SubjectSum / tag1SubjectCount, AvgRd, MidpointRounding.AwayFromZero));
                                            // rankStudents[key].Add(studentID);
                                        }
                                        //類別1加權總分平均排名
                                        if (tag1SubjectCreditSum > 0)
                                        {
                                            studentTag1SubjectSumW.Add(studentID, tag1SubjectSumW);
                                            studentTag1SubjectAvgW.Add(studentID, Math.Round(tag1SubjectSumW / tag1SubjectCreditSum, AvgRd, MidpointRounding.AwayFromZero));
                                            if (rank && studentRec.Status == "一般" && tag1SummaryRank == true)//不在過濾名單且為一般生才做排名
                                            {
                                                //    key = "類別1加權總分排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                                //   if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                //    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                //   ranks[key].Add(tag1SubjectSumW);
                                                //    rankStudents[key].Add(studentID);

                                                //      key = "類別1加權平均排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                                //    if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                //    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                //     ranks[key].Add(Math.Round(tag1SubjectSumW / tag1SubjectCreditSum, AvgRd, MidpointRounding.AwayFromZero));
                                                //     rankStudents[key].Add(studentID);
                                            }
                                        }
                                    }
                                    //類別2總分平均排名
                                    if (tag2SubjectCount > 0)
                                    {
                                        //總分
                                        studentTag2SubjectSum.Add(studentID, tag2SubjectSum);
                                        //平均四捨五入至小數點第二位
                                        studentTag2SubjectAvg.Add(studentID, Math.Round(tag2SubjectSum / tag2SubjectCount, AvgRd, MidpointRounding.AwayFromZero));
                                        if (rank && studentRec.Status == "一般" && tag2SummaryRank == true)//不在過濾名單且為一般生才做排名
                                        {
                                            //   key = "類別2總分排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                            //  if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            // if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            //  ranks[key].Add(tag2SubjectSum);
                                            //   rankStudents[key].Add(studentID);
                                            //key = "類別2平均排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                            //  if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                            // if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                            //  ranks[key].Add(Math.Round(tag2SubjectSum / tag2SubjectCount, AvgRd, MidpointRounding.AwayFromZero));
                                            //  rankStudents[key].Add(studentID);
                                        }
                                        //類別2加權總分平均排名
                                        if (tag2SubjectCreditSum > 0)
                                        {
                                            studentTag2SubjectSumW.Add(studentID, tag2SubjectSumW);
                                            studentTag2SubjectAvgW.Add(studentID, Math.Round(tag2SubjectSumW / tag2SubjectCreditSum, AvgRd, MidpointRounding.AwayFromZero));
                                            if (rank && studentRec.Status == "一般" && tag2SummaryRank == true)//不在過濾名單且為一般生才做排名
                                            {
                                                //     key = "類別2加權總分排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                                //   if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                //    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                //    ranks[key].Add(tag2SubjectSumW);
                                                //    rankStudents[key].Add(studentID);

                                                //    key = "類別2加權平均排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                                //    if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                //     if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                //    ranks[key].Add(Math.Round(tag2SubjectSumW / tag2SubjectCreditSum, AvgRd, MidpointRounding.AwayFromZero));
                                                //  rankStudents[key].Add(studentID);
                                            }
                                        }
                                    }
                                }
                                progressCount++;
                                bkw.ReportProgress(40 + progressCount * 30 / total);
                            }
                        }
                        //   foreach (var k in ranks.Keys)
                        {

                            //      var rankscores = ranks[k];
                            //排序
                            //  rankscores.Sort();
                            //   rankscores.Reverse();
                            //高均標、組距
                            //   if (rankscores.Count > 0)
                            {
                                #region 算高標的中點
                                int middleIndex = 0;
                                int count = 1;
                                //    var score = rankscores[0];
                                //    while (rankscores.Count > middleIndex)
                                {
                                    //     if (score != rankscores[middleIndex])
                                    {
                                        //        if (count * 2 >= rankscores.Count) break;
                                        //          score = rankscores[middleIndex];
                                    }
                                    middleIndex++;
                                    count++;
                                }
                                //  if (rankscores.Count == middleIndex)
                                {
                                    middleIndex--;
                                    count--;
                                }
                                #endregion
                                // analytics.Add(k + "^^^高標", Math.Round(rankscores.GetRange(0, count).Average(), AvgRd, MidpointRounding.AwayFromZero));
                                // analytics.Add(k + "^^^均標", Math.Round(rankscores.Average(), AvgRd, MidpointRounding.AwayFromZero));
                                #region 算低標的中點
                                //  middleIndex = rankscores.Count - 1;
                                count = 1;
                                // score = rankscores[middleIndex];
                                while (middleIndex >= 0)
                                {
                                    //     if (score != rankscores[middleIndex])
                                    {
                                        //        if (count * 2 >= rankscores.Count) break;
                                        //        score = rankscores[middleIndex];
                                    }
                                    middleIndex--;
                                    count++;
                                }
                                if (middleIndex < 0)
                                {
                                    middleIndex++;
                                    count--;
                                }
                                #endregion
                                //    analytics.Add(k + "^^^低標", Math.Round(rankscores.GetRange(middleIndex, count).Average(), AvgRd, MidpointRounding.AwayFromZero));
                                //Compute the Average      
                                //  var avg = (double)rankscores.Average();
                                //Perform the Sum of (value-avg)_2_2      
                                //   var sum = (double)rankscores.Sum(d => Math.Pow((double)d - avg, 2));
                                //Put it all together      
                                //   analytics.Add(k + "^^^標準差", Math.Round((decimal)Math.Sqrt((sum) / rankscores.Count()), AvgRd, MidpointRounding.AwayFromZero));
                            }
                            #region 計算級距
                            int count90 = 0, count80 = 0, count70 = 0, count60 = 0, count50 = 0, count40 = 0, count30 = 0, count20 = 0, count10 = 0;
                            int count100Up = 0, count90Up = 0, count80Up = 0, count70Up = 0, count60Up = 0, count50Up = 0, count40Up = 0, count30Up = 0, count20Up = 0, count10Up = 0;
                            int count90Down = 0, count80Down = 0, count70Down = 0, count60Down = 0, count50Down = 0, count40Down = 0, count30Down = 0, count20Down = 0, count10Down = 0;
                            //  foreach (var score in rankscores)
                            {
                                //if (score >= 100)
                                //    count100Up++;
                                //else if (score >= 90)
                                //    count90++;
                                //else if (score >= 80)
                                //    count80++;
                                //else if (score >= 70)
                                //    count70++;
                                //else if (score >= 60)
                                //    count60++;
                                //else if (score >= 50)
                                //    count50++;
                                //else if (score >= 40)
                                //    count40++;
                                //else if (score >= 30)
                                //    count30++;
                                //else if (score >= 20)
                                //    count20++;
                                //else if (score >= 10)
                                //    count10++;
                                //else
                                //    count10Down++;
                            }
                            count90Up = count100Up + count90;
                            count80Up = count90Up + count80;
                            count70Up = count80Up + count70;
                            count60Up = count70Up + count60;
                            count50Up = count60Up + count50;
                            count40Up = count50Up + count40;
                            count30Up = count40Up + count30;
                            count20Up = count30Up + count20;
                            count10Up = count20Up + count10;

                            count20Down = count10Down + count10;
                            count30Down = count20Down + count20;
                            count40Down = count30Down + count30;
                            count50Down = count40Down + count40;
                            count60Down = count50Down + count50;
                            count70Down = count60Down + count60;
                            count80Down = count70Down + count70;
                            count90Down = count80Down + count80;

                            //analytics.Add(k + "^^^count90", count90);
                            //analytics.Add(k + "^^^count80", count80);
                            //analytics.Add(k + "^^^count70", count70);
                            //analytics.Add(k + "^^^count60", count60);
                            //analytics.Add(k + "^^^count50", count50);
                            //analytics.Add(k + "^^^count40", count40);
                            //analytics.Add(k + "^^^count30", count30);
                            //analytics.Add(k + "^^^count20", count20);
                            //analytics.Add(k + "^^^count10", count10);
                            //analytics.Add(k + "^^^count100Up", count100Up);
                            //analytics.Add(k + "^^^count90Up", count90Up);
                            //analytics.Add(k + "^^^count80Up", count80Up);
                            //analytics.Add(k + "^^^count70Up", count70Up);
                            //analytics.Add(k + "^^^count60Up", count60Up);
                            //analytics.Add(k + "^^^count50Up", count50Up);
                            //analytics.Add(k + "^^^count40Up", count40Up);
                            //analytics.Add(k + "^^^count30Up", count30Up);
                            //analytics.Add(k + "^^^count20Up", count20Up);
                            //analytics.Add(k + "^^^count10Up", count10Up);
                            //analytics.Add(k + "^^^count90Down", count90Down);
                            //analytics.Add(k + "^^^count80Down", count80Down);
                            //analytics.Add(k + "^^^count70Down", count70Down);
                            //analytics.Add(k + "^^^count60Down", count60Down);
                            //analytics.Add(k + "^^^count50Down", count50Down);
                            //analytics.Add(k + "^^^count40Down", count40Down);
                            //analytics.Add(k + "^^^count30Down", count30Down);
                            //analytics.Add(k + "^^^count20Down", count20Down);
                            //analytics.Add(k + "^^^count10Down", count10Down);
                            #endregion
                        }
                        #endregion
                        //參考成績載完
                        elseReady.WaitOne();
                        bkw.ReportProgress(70);
                        progressCount = 0;
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
                            row["試別"] = conf.ExamRecord.Name;

                            row["學年度"] = conf.SchoolYear;
                            row["學期"] = conf.Semester;
                            row["班級科別名稱"] = classRec.Department;
                            row["班級"] = classRec.ClassName;
                            row["定期評量"] = conf.ExamRecord.Name;
                            row["班導師"] = classRec.RefTeacher == null ? "" : classRec.RefTeacher.TeacherName;

                            //整理 組距名稱
                            foreach (string item in new string[] { "類別1", "類別2" })
                            {
                                try
                                {
                                    int tagItemNum = 1;
                                    foreach (string itemName in FixRankTags[classRec.ClassID][$"{item}排名"])
                                    {
                                        row[$"{item}_分組{tagItemNum}名稱"] = itemName;
                                        tagItemNum++;
                                    }
                                }
                                catch (Exception ex)
                                {

                                    Console.WriteLine("錯誤發生:" + ex.Message);
                                }

                            }


                            // 處理學生類別1 類別2 之 類別名稱比如 

                            foreach (StudentRecord stuRec in classRec.Students)
                            {
                                // 


                            }

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

                                // 處理學生所屬類別
                                if (studentTags1FixRankSByClass.ContainsKey(classRec.ClassID))
                                {
                                    if (studentTags1FixRankSByClass[classRec.ClassID].ContainsKey(studentID))
                                    {
                                        row[$"類別1_分組名稱{ClassStuNum}"] = studentTags1FixRankSByClass[classRec.ClassID][studentID];
                                    }
                                }

                                if (studentTags2FixRankSByClass.ContainsKey(classRec.ClassID))
                                {
                                    if (studentTags2FixRankSByClass[classRec.ClassID].ContainsKey(studentID))
                                    {
                                        row[$"類別2_分組名稱{ClassStuNum}"] = studentTags2FixRankSByClass[classRec.ClassID][studentID];
                                    }
                                }


                                row["座號" + ClassStuNum] = stuRec.SeatNo;
                                row["學號" + ClassStuNum] = stuRec.StudentNumber;
                                row["姓名" + ClassStuNum] = stuRec.StudentName;
                                //整理這個班的學生所屬類別
                                if (studentTag1Group.ContainsKey(studentID) && !tag1List.Contains(studentTag1Group[studentID]))
                                {
                                    tag1List.Add(studentTag1Group[studentID]);
                                    foreach (var tag in studentTags[studentID])
                                    {
                                        if (tag.RefTagID == studentTag1Group[studentID])
                                        {
                                            row["類別排名1"] += (("" + row["類別排名1"]) == "" ? "" : "/") + tag.Name;
                                        }
                                    }
                                }
                                if (studentTag2Group.ContainsKey(studentID) && !tag2List.Contains(studentTag2Group[studentID]))
                                {
                                    tag2List.Add(studentTag2Group[studentID]);
                                    foreach (var tag in studentTags[studentID])
                                    {
                                        if (tag.RefTagID == studentTag2Group[studentID])
                                        {
                                            row["類別排名2"] += (("" + row["類別排名2"]) == "" ? "" : "/") + tag.Name;
                                        }
                                    }
                                }

                                #region 整理班級中列印科目
                                if (studentExamSores.ContainsKey(stuRec.StudentID))
                                {
                                    foreach (var subjectName in studentExamSores[studentID].Keys)
                                    {
                                        foreach (var courseID in studentExamSores[studentID][subjectName].Keys)
                                        {
                                            string subjectLevel = accessHelper.CourseHelper.GetCourse(courseID)[0].SubjectLevel;
                                            string credit = "" + accessHelper.CourseHelper.GetCourse(courseID)[0].CreditDec();
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
                                                if (studentExamSores.ContainsKey(studentID))
                                                {
                                                    if (studentExamSores[studentID].ContainsKey(subjectName))
                                                    {
                                                        foreach (var courseID in studentExamSores[studentID][subjectName].Keys)
                                                        {
                                                            if (accessHelper.CourseHelper.GetCourse(courseID)[0].SubjectLevel == subjectLevel)//studentExamSores[studentID][subjectName][courseID]可能是null
                                                            {
                                                                #region 評量成績
                                                                var sceTakeRecord = studentExamSores[studentID][subjectName][courseID];
                                                                if (sceTakeRecord != null)
                                                                {
                                                                    row["科目成績" + ClassStuNum + "-" + subjectIndex] = sceTakeRecord.SpecialCase == "" ? ("" + sceTakeRecord.ExamScore) : sceTakeRecord.SpecialCase;
                                                                    #region 班排名及落點分析
                                                                    if (stuRec.RefClass != null)
                                                                    {
                                                                        //Jean 改固定  

                                                                        //***對照Dictionary 之 key = $"{itemType}^^^{intervalInfo.RankType}^^^grade{intervalInfo.GradeYear}^^^{intervalInfo.RankName}^^^{intervalInfo.ItemName}";

                                                                        key = "科目成績" + "^^^" + "班排名" + "^^^" + "grade" + stuRec.RefClass.GradeYear + "^^^" + stuRec.RefClass.ClassName + "^^^" + sceTakeRecord.Subject;//@@@@@ 取固定排名的key值


                                                                        if (dicFixedRankData.ContainsKey(stuRec.RefClass.ClassID))
                                                                        {
                                                                            if (dicFixedRankData[stuRec.RefClass.ClassID].ContainsKey(studentID))
                                                                                if (dicFixedRankData[stuRec.RefClass.ClassID][studentID].DicSubjectFixRank.ContainsKey(key))
                                                                                {
                                                                                    row["班排名" + ClassStuNum + "-" + subjectIndex] = dicFixedRankData[stuRec.RefClass.ClassID][studentID].DicSubjectFixRank[key].Rank;
                                                                                    row["班排名母數" + ClassStuNum + "-" + subjectIndex] = dicFixedRankData[stuRec.RefClass.ClassID][studentID].DicSubjectFixRank[key].MatrixCount;
                                                                                }
                                                                                else
                                                                                {
                                                                                    Console.WriteLine("oop~  You should go to see the Dictionary Key ....can't find the right key.");
                                                                                }
                                                                        }
                                                                    }
                                                                    #endregion
                                                                    #region 科排名及落點分析
                                                                    if (stuRec.Department != "")
                                                                    {

                                                                        key = "科目成績" + "^^^" + "科排名" + "^^^" + "grade" + stuRec.RefClass.GradeYear + "^^^" + stuRec.RefClass.Department + "^^^" + sceTakeRecord.Subject;//@@@@@ 取固定排名的key值

                                                                        // Jean 改固定排名
                                                                        if (dicFixedRankData.ContainsKey(stuRec.RefClass.ClassID)) //如果固定排名裡有學生的ID      
                                                                        {
                                                                            if (dicFixedRankData[stuRec.RefClass.ClassID].ContainsKey(studentID))
                                                                            {
                                                                                if (dicFixedRankData[stuRec.RefClass.ClassID][studentID].DicSubjectFixRank.ContainsKey(key))
                                                                                {
                                                                                    row["科排名" + ClassStuNum + "-" + subjectIndex] = dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectFixRank[key].Rank;
                                                                                    row["科排名母數" + ClassStuNum + "-" + subjectIndex] = dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectFixRank[key].MatrixCount;
                                                                                }
                                                                            }
                                                                        }
                                                                    }

                                                                    #endregion
                                                                    #region 全校(年)排名及落點分析

                                                                    key = "科目成績" + "^^^" + "年排名" + "^^^" + "grade" + stuRec.RefClass.GradeYear + "^^^" + stuRec.RefClass.GradeYear + "年級" + "^^^" + sceTakeRecord.Subject;//@@@@@ 取固定排名的key值


                                                                    if (dicFixedRankData.ContainsKey(stuRec.RefClass.ClassID)) //如果固定排名裡有學生的ID      
                                                                    {
                                                                        if (dicFixedRankData[stuRec.RefClass.ClassID].ContainsKey(stuRec.StudentID))

                                                                        {
                                                                            if (dicFixedRankData[stuRec.RefClass.ClassID][studentID].DicSubjectFixRank.ContainsKey(key))
                                                                            {
                                                                                row["年排名" + ClassStuNum + "-" + subjectIndex] = dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectFixRank[key].Rank;
                                                                                row["年排名母數" + ClassStuNum + "-" + subjectIndex] = dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectFixRank[key].MatrixCount;
                                                                            }
                                                                        }
                                                                    }

                                                                    #endregion
                                                                    #region 類別1排名及落點分析
                                                                    // if (studentTag1Group.ContainsKey(studentID) && conf.TagRank1SubjectList.Contains(subjectName))
                                                                    {
                                                                        //key = "類別1排名" + studentTag1Group[studentID] + "^^^" + gradeYear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;

                                                                        //如果有

                                                                        if (studentTags.ContainsKey(studentID))
                                                                        {
                                                                            foreach (K12.Data.StudentTagRecord stuTag in studentTags[studentID])
                                                                            {
                                                                                string studentTag = stuTag.Name;
                                                                                key = "科目成績" + "^^^" + "類別1排名" + "^^^" + "grade" + stuRec.RefClass.GradeYear + "^^^" + studentTag + "^^^" + sceTakeRecord.Subject;//@@@@@ 取固定排名的key值

                                                                                if (dicFixedRankData.ContainsKey(stuRec.RefClass.ClassID)) //如果固定排名裡有學生的ID      
                                                                                {
                                                                                    if (dicFixedRankData[stuRec.RefClass.ClassID].ContainsKey(stuRec.StudentID))
                                                                                    {
                                                                                        if (dicFixedRankData[stuRec.RefClass.ClassID][studentID].DicSubjectFixRank.ContainsKey(key))
                                                                                        {
                                                                                            row["類別1排名" + ClassStuNum + "-" + subjectIndex] = dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectFixRank[key].Rank;
                                                                                            row["類別1排名母數" + ClassStuNum + "-" + subjectIndex] = dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectFixRank[key].MatrixCount;
                                                                                        }
                                                                                    }
                                                                                }
                                                                            }



                                                                        }




                                                                    }
                                                                    #endregion
                                                                    #region 類別2排名及落點分析

                                                                    if (studentTags.ContainsKey(studentID))
                                                                    {
                                                                        foreach (K12.Data.StudentTagRecord stuTag in studentTags[studentID])
                                                                        {
                                                                            string studentTag = stuTag.Name;
                                                                            key = "科目成績" + "^^^" + "類別2排名" + "^^^" + "grade" + stuRec.RefClass.GradeYear + "^^^" + studentTag + "^^^" + sceTakeRecord.Subject;//@@@@@ 取固定排名的key值
                                                                            if (dicFixedRankData.ContainsKey(stuRec.RefClass.ClassID)) //如果固定排名裡有學生的ID      
                                                                            {
                                                                                if (dicFixedRankData[stuRec.RefClass.ClassID].ContainsKey(stuRec.StudentID))
                                                                                {
                                                                                    if (dicFixedRankData[stuRec.RefClass.ClassID][studentID].DicSubjectFixRank.ContainsKey(key))
                                                                                    {
                                                                                        row["類別2排名" + ClassStuNum + "-" + subjectIndex] = dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectFixRank[key].Rank;
                                                                                        row["類別2排名母數" + ClassStuNum + "-" + subjectIndex] = dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectFixRank[key].MatrixCount;
                                                                                    }
                                                                                }
                                                                            }
                                                                        }
                                                                    }
                                                                    #endregion
                                                                }
                                                                else
                                                                {//修課有該考試但沒有成績資料
                                                                    row["科目成績" + ClassStuNum + "-" + subjectIndex] = "未輸入";
                                                                }
                                                                #endregion
                                                                #region 參考成績
                                                                if (studentRefExamSores.ContainsKey(studentID) && studentRefExamSores[studentID].ContainsKey(courseID))
                                                                {
                                                                    row["前次成績" + ClassStuNum + "-" + subjectIndex] =
                                                                            studentRefExamSores[studentID][courseID].SpecialCase == ""
                                                                            ? ("" + studentRefExamSores[studentID][courseID].ExamScore)
                                                                            : studentRefExamSores[studentID][courseID].SpecialCase;
                                                                }
                                                                #endregion
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            #region 【固定排名】【班】【科目】五標 及 組距 (1081017 Jean 改成固定排名)

                                            //科目班標

                                            key = $"科目成績^^^班排名^^^grade{classRec.GradeYear}^^^{classRec.ClassName}^^^{subjectName}";

                                            if (IntervalInfos.ContainsKey(classRec.ClassID) && IntervalInfos[classRec.ClassID].ContainsKey(key))
                                            {

                                                row["班頂標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_top_25;
                                                row["班高標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_top_50;
                                                row["班均標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg;
                                                row["班低標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_bottom_50;
                                                row["班底標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_bottom_25;

                                                row["班組距" + subjectIndex + "count90"] = IntervalInfos[classRec.ClassID][key].Level_90;
                                                row["班組距" + subjectIndex + "count80"] = IntervalInfos[classRec.ClassID][key].Level_80;
                                                row["班組距" + subjectIndex + "count70"] = IntervalInfos[classRec.ClassID][key].Level_70;
                                                row["班組距" + subjectIndex + "count60"] = IntervalInfos[classRec.ClassID][key].Level_60;
                                                row["班組距" + subjectIndex + "count50"] = IntervalInfos[classRec.ClassID][key].Level_50;
                                                row["班組距" + subjectIndex + "count40"] = IntervalInfos[classRec.ClassID][key].Level_40;
                                                row["班組距" + subjectIndex + "count30"] = IntervalInfos[classRec.ClassID][key].Level_30;
                                                row["班組距" + subjectIndex + "count20"] = IntervalInfos[classRec.ClassID][key].Level_20;
                                                row["班組距" + subjectIndex + "count10"] = IntervalInfos[classRec.ClassID][key].Level_10;
                                                row["班組距" + subjectIndex + "count100Up"] = IntervalInfos[classRec.ClassID][key].Level_gte100;
                                                row["班組距" + subjectIndex + "count90Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[90];
                                                row["班組距" + subjectIndex + "count80Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[80];
                                                row["班組距" + subjectIndex + "count70Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[70];
                                                row["班組距" + subjectIndex + "count60Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[60];
                                                row["班組距" + subjectIndex + "count50Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[50];
                                                row["班組距" + subjectIndex + "count40Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[40];
                                                row["班組距" + subjectIndex + "count30Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[30];
                                                row["班組距" + subjectIndex + "count20Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[20];
                                                row["班組距" + subjectIndex + "count10Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[10];
                                                row["班組距" + subjectIndex + "count90Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[90];
                                                row["班組距" + subjectIndex + "count80Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[80];
                                                row["班組距" + subjectIndex + "count70Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[70];
                                                row["班組距" + subjectIndex + "count60Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[60];
                                                row["班組距" + subjectIndex + "count50Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[50];
                                                row["班組距" + subjectIndex + "count40Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[40];
                                                row["班組距" + subjectIndex + "count30Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[30];
                                                row["班組距" + subjectIndex + "count20Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[20];
                                                row["班組距" + subjectIndex + "count10Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[10];

                                            }

                                            if (IntervalInfos.ContainsKey(classRec.ClassID))
                                            {
                                                //需替換固定排名 by Jean
                                                try
                                                {
                                                    if (IntervalInfos[classRec.ClassID].ContainsKey(subjectName))
                                                    {
                                                    }
                                                }
                                                catch (Exception ex)
                                                {
                                                    Console.WriteLine(ex);
                                                }

                                            }

                                            #endregion
                                            #region 【固定排名】【科】【科目】五標 及 組距

                                            key = $"科目成績^^^科排名^^^grade{classRec.GradeYear}^^^{classRec.Department}^^^{subjectName}";

                                            if (IntervalInfos.ContainsKey(classRec.ClassID) && IntervalInfos[classRec.ClassID].ContainsKey(key))
                                            {
                                                //Jean 固定排名 (科組距)

                                                row["科頂標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_top_25;
                                                row["科高標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_top_50;
                                                row["科均標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg;
                                                row["科低標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_bottom_50;
                                                row["科底標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_bottom_25;

                                                row["科組距" + subjectIndex + "count90"] = IntervalInfos[classRec.ClassID][key].Level_90;
                                                row["科組距" + subjectIndex + "count80"] = IntervalInfos[classRec.ClassID][key].Level_80;
                                                row["科組距" + subjectIndex + "count70"] = IntervalInfos[classRec.ClassID][key].Level_70;
                                                row["科組距" + subjectIndex + "count60"] = IntervalInfos[classRec.ClassID][key].Level_60;
                                                row["科組距" + subjectIndex + "count50"] = IntervalInfos[classRec.ClassID][key].Level_50;
                                                row["科組距" + subjectIndex + "count40"] = IntervalInfos[classRec.ClassID][key].Level_40;
                                                row["科組距" + subjectIndex + "count30"] = IntervalInfos[classRec.ClassID][key].Level_30;
                                                row["科組距" + subjectIndex + "count20"] = IntervalInfos[classRec.ClassID][key].Level_20;
                                                row["科組距" + subjectIndex + "count10"] = IntervalInfos[classRec.ClassID][key].Level_10;
                                                row["科組距" + subjectIndex + "count100Up"] = IntervalInfos[classRec.ClassID][key].Level_gte100;
                                                row["科組距" + subjectIndex + "count90Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[90];
                                                row["科組距" + subjectIndex + "count80Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[80];
                                                row["科組距" + subjectIndex + "count70Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[70];
                                                row["科組距" + subjectIndex + "count60Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[60];
                                                row["科組距" + subjectIndex + "count50Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[50];
                                                row["科組距" + subjectIndex + "count40Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[40];
                                                row["科組距" + subjectIndex + "count30Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[30];
                                                row["科組距" + subjectIndex + "count20Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[20];
                                                row["科組距" + subjectIndex + "count10Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[10];
                                                row["科組距" + subjectIndex + "count90Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[90];
                                                row["科組距" + subjectIndex + "count80Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[80];
                                                row["科組距" + subjectIndex + "count70Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[70];
                                                row["科組距" + subjectIndex + "count60Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[60];
                                                row["科組距" + subjectIndex + "count50Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[50];
                                                row["科組距" + subjectIndex + "count40Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[40];
                                                row["科組距" + subjectIndex + "count30Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[30];
                                                row["科組距" + subjectIndex + "count20Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[20];
                                                row["科組距" + subjectIndex + "count10Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[10];

                                            }
                                            #endregion
                                            #region 【固定排名】【年】【科目】五標 及 組距 

                                            key = $"科目成績^^^年排名^^^grade{classRec.GradeYear}^^^{classRec.GradeYear}年級^^^{subjectName}";

                                            if (IntervalInfos.ContainsKey(classRec.ClassID) && IntervalInfos[classRec.ClassID].ContainsKey(key))
                                            {
                                                row["年頂標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_top_25;
                                                row["年高標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_top_50;
                                                row["年均標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg;
                                                row["年低標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_bottom_50;
                                                row["年底標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_bottom_25;

                                                row["年組距" + subjectIndex + "count90"] = IntervalInfos[classRec.ClassID][key].Level_90;
                                                row["年組距" + subjectIndex + "count80"] = IntervalInfos[classRec.ClassID][key].Level_80;
                                                row["年組距" + subjectIndex + "count70"] = IntervalInfos[classRec.ClassID][key].Level_70;
                                                row["年組距" + subjectIndex + "count60"] = IntervalInfos[classRec.ClassID][key].Level_60;
                                                row["年組距" + subjectIndex + "count50"] = IntervalInfos[classRec.ClassID][key].Level_50;
                                                row["年組距" + subjectIndex + "count40"] = IntervalInfos[classRec.ClassID][key].Level_40;
                                                row["年組距" + subjectIndex + "count30"] = IntervalInfos[classRec.ClassID][key].Level_30;
                                                row["年組距" + subjectIndex + "count20"] = IntervalInfos[classRec.ClassID][key].Level_20;
                                                row["年組距" + subjectIndex + "count10"] = IntervalInfos[classRec.ClassID][key].Level_10;
                                                row["年組距" + subjectIndex + "count100Up"] = IntervalInfos[classRec.ClassID][key].Level_gte100;
                                                row["年組距" + subjectIndex + "count90Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[90];
                                                row["年組距" + subjectIndex + "count80Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[80];
                                                row["年組距" + subjectIndex + "count70Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[70];
                                                row["年組距" + subjectIndex + "count60Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[60];
                                                row["年組距" + subjectIndex + "count50Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[50];
                                                row["年組距" + subjectIndex + "count40Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[40];
                                                row["年組距" + subjectIndex + "count30Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[30];
                                                row["年組距" + subjectIndex + "count20Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[20];
                                                row["年組距" + subjectIndex + "count10Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[10];
                                                row["年組距" + subjectIndex + "count90Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[90];
                                                row["年組距" + subjectIndex + "count80Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[80];
                                                row["年組距" + subjectIndex + "count70Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[70];
                                                row["年組距" + subjectIndex + "count60Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[60];
                                                row["年組距" + subjectIndex + "count50Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[50];
                                                row["年組距" + subjectIndex + "count40Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[40];
                                                row["年組距" + subjectIndex + "count30Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[30];
                                                row["年組距" + subjectIndex + "count20Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[20];
                                                row["年組距" + subjectIndex + "count10Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[10];
                                            }

                                            #endregion

                                            if (tagRankInfoFromRankMatrix.Count > 0)
                                            {
                                                #region 【固定排名】【類別1】【科目】五標 及 組距  
                                                // 【說明】
                                                // todo 
                                                // 先取得有類1的名稱並且類1下有哪些分類

                                                if (FixRankTags.ContainsKey(classRec.ClassID))  //如果本次有計算類別1
                                                {
                                                    int number = 1;
                                                    if (FixRankTags[classRec.ClassID].ContainsKey("類別1排名"))
                                                    {
                                                        foreach (string rankName in FixRankTags[classRec.ClassID]["類別1排名"])
                                                        {
                                                            // row[$"類1_分組{number}名稱"] = rankName;

                                                            key = $"科目成績^^^類別1排名^^^grade{classRec.GradeYear}^^^{rankName}^^^{subjectName}";

                                                            if (IntervalInfos[classRec.ClassID].ContainsKey(key))
                                                            {
                                                                row[$"類別1_分組{number}頂標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_top_25;
                                                                row[$"類別1_分組{number}高標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_top_50;
                                                                row[$"類別1_分組{number}均標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg;
                                                                row[$"類別1_分組{number}低標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_bottom_50;
                                                                row[$"類別1_分組{number}底標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_bottom_25;

                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count90"] = IntervalInfos[classRec.ClassID][key].Level_90;
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count80"] = IntervalInfos[classRec.ClassID][key].Level_80;
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count70"] = IntervalInfos[classRec.ClassID][key].Level_70;
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count60"] = IntervalInfos[classRec.ClassID][key].Level_60;
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count50"] = IntervalInfos[classRec.ClassID][key].Level_50;
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count40"] = IntervalInfos[classRec.ClassID][key].Level_40;
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count30"] = IntervalInfos[classRec.ClassID][key].Level_30;
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count20"] = IntervalInfos[classRec.ClassID][key].Level_20;
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count10"] = IntervalInfos[classRec.ClassID][key].Level_10;
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count100Up"] = IntervalInfos[classRec.ClassID][key].Level_gte100;
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count90Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[90];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count80Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[80];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count70Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[70];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count60Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[60];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count50Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[50];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count40Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[40];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count30Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[30];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count20Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[20];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count10Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[10];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count90Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[90];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count80Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[80];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count70Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[70];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count60Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[60];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count50Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[50];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count40Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[40];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count30Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[30];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count20Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[20];
                                                                row[$"類別1_分組{number}組距" + subjectIndex + "count10Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[10];
                                                                number++;
                                                                // subjectIndex++;

                                                            }
                                                            else
                                                            {
                                                                Console.WriteLine(key + "不包含在dic 裡面");
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {

                                                        Console.WriteLine(classRec.ClassID + " 沒有 固定排名 類別1 之 排名 紀錄 ");

                                                    }

                                                }

                                                #endregion
                                                #region 【固定排名】【類別2】【科目】五標 及 組距  
                                                // 【說明】
                                                // todo 
                                                // 先取得有類1的名稱並且類1下有哪些分類

                                                if (FixRankTags.ContainsKey(classRec.ClassID))  //如果本次有計算類別1
                                                {
                                                    int number = 1;
                                                    if (FixRankTags[classRec.ClassID].ContainsKey("類別2排名"))
                                                    {
                                                        foreach (string rankName in FixRankTags[classRec.ClassID]["類別2排名"])
                                                        {
                                                            // row[$"類1_分組{number}名稱"] = rankName;

                                                            key = $"科目成績^^^類別2排名^^^grade{classRec.GradeYear}^^^{rankName}^^^{subjectName}";

                                                            if (IntervalInfos[classRec.ClassID].ContainsKey(key))
                                                            {
                                                                row[$"類別2_分組{number}頂標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_top_25;
                                                                row[$"類別2_分組{number}高標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_top_50;
                                                                row[$"類別2_分組{number}均標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg;
                                                                row[$"類別2_分組{number}低標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_bottom_50;
                                                                row[$"類別2_分組{number}底標" + subjectIndex] = IntervalInfos[classRec.ClassID][key].Avg_bottom_25;

                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count90"] = IntervalInfos[classRec.ClassID][key].Level_90;
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count80"] = IntervalInfos[classRec.ClassID][key].Level_80;
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count70"] = IntervalInfos[classRec.ClassID][key].Level_70;
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count60"] = IntervalInfos[classRec.ClassID][key].Level_60;
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count50"] = IntervalInfos[classRec.ClassID][key].Level_50;
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count40"] = IntervalInfos[classRec.ClassID][key].Level_40;
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count30"] = IntervalInfos[classRec.ClassID][key].Level_30;
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count20"] = IntervalInfos[classRec.ClassID][key].Level_20;
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count10"] = IntervalInfos[classRec.ClassID][key].Level_10;
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count100Up"] = IntervalInfos[classRec.ClassID][key].Level_gte100;
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count90Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[90];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count80Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[80];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count70Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[70];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count60Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[60];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count50Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[50];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count40Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[40];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count30Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[30];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count20Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[20];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count10Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[10];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count90Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[90];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count80Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[80];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count70Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[70];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count60Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[60];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count50Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[50];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count40Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[40];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count30Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[30];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count20Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[20];
                                                                row[$"類別2_分組{number}組距" + subjectIndex + "count10Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[10];
                                                                number++;
                                                                // subjectIndex++;

                                                            }
                                                            else
                                                            {
                                                                Console.WriteLine(key + "不包含");
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {

                                                        Console.WriteLine(classRec.ClassID + " 沒有 固定排名 類別2 之 排名 紀錄 ");

                                                    }
                                                }

                                                #endregion
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
                            #region 總分
                            ClassStuNum = 0;
                            string _studentID = "";


                            foreach (StudentRecord stuRec in classRec.Students)
                            {
                                ClassStuNum++;
                                if (ClassStuNum > conf.StudentLimit)
                                    break;
                                _studentID = stuRec.StudentID;
                                if (studentPrintSubjectSum.ContainsKey(_studentID))
                                {
                                    row["總分" + ClassStuNum] = studentPrintSubjectSum[_studentID];

                                    key = "總計成績" + "^^^" + "班排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.ClassName + "^^^" + "總分";

                                    row = PutRankValueToTable(dicFixedRankData, row, stuRec, "總分", "班", ClassStuNum, key);

                                    key = "總計成績" + "^^^" + "科排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.Department + "^^^" + "總分";

                                    row = PutRankValueToTable(dicFixedRankData, row, stuRec, "總分", "科", ClassStuNum, key);

                                    key = "總計成績" + "^^^" + "年排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.GradeYear + "年級" + "^^^" + "總分";

                                    row = PutRankValueToTable(dicFixedRankData, row, stuRec, "總分", "年", ClassStuNum, key);
                                }
                            }


                            #region 總分班組距及高低標 【總分】【班】【級距】

                            key = $"總計成績^^^班排名^^^grade{classRec.GradeYear}^^^{classRec.ClassName}^^^總分";
                            if (IntervalInfos.ContainsKey(classRec.ClassID) && IntervalInfos[classRec.ClassID].ContainsKey(key))
                            {
                                key = $"總計成績^^^班排名^^^grade{classRec.GradeYear}^^^{classRec.ClassName}^^^總分";

                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "總分", "班", key);

                            }
                            #endregion
                            #region 總分科組距及高低標

                            key = $"總計成績^^^科排名^^^grade{classRec.GradeYear}^^^{classRec.Department}^^^總分";
                            if (IntervalInfos.ContainsKey(classRec.ClassID) && IntervalInfos[classRec.ClassID].ContainsKey(key))
                            {

                                key = $"總計成績^^^科排名^^^grade{classRec.GradeYear}^^^{classRec.Department}^^^總分";

                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "總分", "科", key);

                            }

                            #endregion
                            #region 總分校組距及高低標


                            key = $"總計成績^^^年排名^^^grade{classRec.GradeYear}^^^{classRec.GradeYear}年級^^^總分";
                            if (IntervalInfos.ContainsKey(classRec.ClassID) && IntervalInfos[classRec.ClassID].ContainsKey(key))
                            {


                                key = $"總計成績^^^年排名^^^grade{classRec.GradeYear}^^^{classRec.GradeYear}年級^^^總分";

                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "總分", "年", key);
                            }

                            #endregion
                            #endregion
                            #region 平均(算數)---排名
                            ClassStuNum = 0;
                            foreach (StudentRecord stuRec in classRec.Students)
                            {
                                ClassStuNum++;
                                if (ClassStuNum > conf.StudentLimit)
                                    break;
                                string studentID = stuRec.StudentID;
                                if (studentPrintSubjectAvg.ContainsKey(studentID))
                                {
                                    row["平均" + ClassStuNum] = studentPrintSubjectAvg[studentID];

                                    key = "總計成績" + "^^^" + "班排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.ClassName + "^^^" + "平均";

                                    row = PutRankValueToTable(dicFixedRankData, row, stuRec, "平均", "班", ClassStuNum, key);

                                    key = "總計成績" + "^^^" + "科排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.Department + "^^^" + "平均";

                                    row = PutRankValueToTable(dicFixedRankData, row, stuRec, "平均", "科", ClassStuNum, key);

                                    key = "總計成績" + "^^^" + "年排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.GradeYear + "年級" + "^^^" + "平均";

                                    row = PutRankValueToTable(dicFixedRankData, row, stuRec, "平均", "年", ClassStuNum, key);
                                }
                            }
                            #region 平均(算數)【班】【組距】

                            if (IntervalInfos.ContainsKey(classRec.ClassID) && IntervalInfos[classRec.ClassID].ContainsKey(key))
                            {
                                key = $"總計成績^^^班排名^^^grade{classRec.GradeYear}^^^{classRec.ClassName}^^^平均";

                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "平均", "班", key);
                            }

                            #endregion
                            #region 平均(算數) 【科】【組距】
                            if (IntervalInfos.ContainsKey(classRec.ClassID) && IntervalInfos[classRec.ClassID].ContainsKey(key))
                            {
                                key = $"總計成績^^^科排名^^^grade{classRec.GradeYear}^^^{classRec.Department}^^^平均";

                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "平均", "科", key);

                                #endregion
                                #region 平均(算數)【年】【組距】

                                key = $"總計成績^^^年排名^^^grade{classRec.GradeYear}^^^{classRec.GradeYear}年級^^^平均";

                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "平均", "年", key);

                                #endregion
                                #endregion

                                #region 加權總分 
                                ClassStuNum = 0;
                                foreach (StudentRecord stuRec in classRec.Students)
                                {
                                    ClassStuNum++;
                                    if (ClassStuNum > conf.StudentLimit)
                                        break;
                                    string studentID = stuRec.StudentID;
                                    if (studentPrintSubjectSumW.ContainsKey(studentID))
                                    {
                                        row["加權總分" + ClassStuNum] = studentPrintSubjectSumW[studentID];


                                        key = "總計成績" + "^^^" + "班排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.ClassName + "^^^" + "加權總分";

                                        row = PutRankValueToTable(dicFixedRankData, row, stuRec, "加權總分", "班", ClassStuNum, key);

                                        key = "總計成績" + "^^^" + "科排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.Department + "^^^" + "加權總分";

                                        row = PutRankValueToTable(dicFixedRankData, row, stuRec, "加權總分", "科", ClassStuNum, key);

                                        key = "總計成績" + "^^^" + "年排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.GradeYear + "年級" + "^^^" + "加權總分";

                                        row = PutRankValueToTable(dicFixedRankData, row, stuRec, "加權總分", "年", ClassStuNum, key);


                                    }
                                }
                                #region 加權總分 -【班】【組距】
                                key = $"總計成績^^^班排名^^^grade{classRec.GradeYear}^^^{classRec.ClassName}^^^加權總分";
                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "加權總分", "班", key);
                                #endregion

                                #region 加權總分-【科】【組距】

                                key = $"總計成績^^^科排名^^^grade{classRec.GradeYear}^^^{classRec.Department}^^^加權總分";
                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "加權總分", "科", key);

                                #endregion

                                #region 加權總分-【年】【組距】
                                key = $"總計成績^^^年排名^^^grade{classRec.GradeYear}^^^{classRec.GradeYear}年級^^^加權總分";
                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "加權總分", "年", key);

                                #endregion


                                #endregion

                                #region 加權平均 
                                ClassStuNum = 0;
                                foreach (StudentRecord stuRec in classRec.Students)
                                {
                                    ClassStuNum++;
                                    if (ClassStuNum > conf.StudentLimit)
                                        break;
                                    string studentID = stuRec.StudentID;
                                    if (studentPrintSubjectAvgW.ContainsKey(studentID))
                                    {
                                        row["加權平均" + ClassStuNum] = studentPrintSubjectAvgW[studentID];
                                        //key = "班排名" + stuRec.RefClass.ClassName + "^^^" + "加權平均";

                                        key = "總計成績" + "^^^" + "班排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.ClassName + "^^^" + "加權平均";

                                        row = PutRankValueToTable(dicFixedRankData, row, stuRec, "加權平均", "班", ClassStuNum, key);

                                        key = "總計成績" + "^^^" + "科排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.Department + "^^^" + "加權平均";

                                        row = PutRankValueToTable(dicFixedRankData, row, stuRec, "加權平均", "科", ClassStuNum, key);

                                        key = "總計成績" + "^^^" + "年排名" + "^^^" + "grade" + gradeYear + "^^^" + stuRec.RefClass.GradeYear + "年級" + "^^^" + "加權平均";

                                        row = PutRankValueToTable(dicFixedRankData, row, stuRec, "加權平均", "年", ClassStuNum, key);

                                    }
                                }
                                #region 加權平均【加權平均】【班】 【組距】
                                key = $"總計成績^^^班排名^^^grade{classRec.GradeYear}^^^{classRec.ClassName}^^^加權平均";
                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "加權平均", "班", key);
                                #endregion

                                #region 加權平均【科】 【組距】
                                key = $"總計成績^^^科排名^^^grade{classRec.GradeYear}^^^{classRec.Department}^^^加權平均";
                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "加權平均", "科", key);
                                #endregion

                                #region 加權平均【年】 【組距】
                                key = $"總計成績^^^年排名^^^grade{classRec.GradeYear}^^^{classRec.GradeYear}年級^^^加權平均";
                                row = PutRankIntervalToTable(IntervalInfos, row, classRec, "加權平均", "年", key);
                                #endregion


                                // 總計成績 (類別1 下 類別2)

                                #region 總計成績 (類別1 下 類別2)
                                foreach (string totalScoreItemName in new string[] { "總分", "平均", "加權總分", "加權平均" })
                                {
                                    #region 加權平均【類別1】 【組距】

                                    if (FixRankTags.ContainsKey(classRec.ClassID))  //如果本次有計算類別1
                                    {
                                        int number = 1;
                                        if (FixRankTags[classRec.ClassID].ContainsKey("類別1排名"))
                                        {
                                            foreach (string rankName in FixRankTags[classRec.ClassID]["類別1排名"])
                                            {
                                                // row[$"類1_分組{number}名稱"] = rankName;
                                                //string totalScoreItemName = "加權平均";

                                                key = $"總計成績^^^類別1排名^^^grade{classRec.GradeYear}^^^{rankName}^^^{totalScoreItemName}";

                                                if (IntervalInfos[classRec.ClassID].ContainsKey(key))
                                                {
                                                    row[$"{totalScoreItemName}類別1_分組{number}頂標"] = IntervalInfos[classRec.ClassID][key].Avg_top_25;
                                                    row[$"{totalScoreItemName}類別1_分組{number}高標"] = IntervalInfos[classRec.ClassID][key].Avg_top_50;
                                                    row[$"{totalScoreItemName}類別1_分組{number}均標"] = IntervalInfos[classRec.ClassID][key].Avg;
                                                    row[$"{totalScoreItemName}類別1_分組{number}低標"] = IntervalInfos[classRec.ClassID][key].Avg_bottom_50;
                                                    row[$"{totalScoreItemName}類別1_分組{number}底標"] = IntervalInfos[classRec.ClassID][key].Avg_bottom_25;

                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count90"] = IntervalInfos[classRec.ClassID][key].Level_90;
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count80"] = IntervalInfos[classRec.ClassID][key].Level_80;
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count70"] = IntervalInfos[classRec.ClassID][key].Level_70;
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count60"] = IntervalInfos[classRec.ClassID][key].Level_60;
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count50"] = IntervalInfos[classRec.ClassID][key].Level_50;
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count40"] = IntervalInfos[classRec.ClassID][key].Level_40;
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count30"] = IntervalInfos[classRec.ClassID][key].Level_30;
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count20"] = IntervalInfos[classRec.ClassID][key].Level_20;
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count10"] = IntervalInfos[classRec.ClassID][key].Level_10;
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count100Up"] = IntervalInfos[classRec.ClassID][key].Level_gte100;
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count90Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[90];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count80Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[80];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count70Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[70];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count60Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[60];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count50Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[50];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count40Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[40];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count30Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[30];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count20Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[20];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count10Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[10];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count90Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[90];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count80Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[80];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count70Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[70];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count60Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[60];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count50Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[50];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count40Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[40];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count30Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[30];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count20Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[20];
                                                    row[$"{totalScoreItemName}類別1_分組{number}組距count10Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[10];

                                                    number++;
                                                    //subjectIndex++;

                                                }
                                                else
                                                {

                                                }
                                            }
                                        }

                                    }

                                    #endregion

                                    #region 加權平均【類別2】 【組距】

                                    if (FixRankTags.ContainsKey(classRec.ClassID))  //如果本次有計算類別1
                                    {
                                        int number = 1;
                                        if (FixRankTags[classRec.ClassID].ContainsKey("類別2排名"))
                                        {
                                            foreach (string rankName in FixRankTags[classRec.ClassID]["類別2排名"])
                                            {
                                                // row[$"類1_分組{number}名稱"] = rankName;
                                                // string totalScoreItemName = "加權平均";

                                                key = $"總計成績^^^類別2排名^^^grade{classRec.GradeYear}^^^{rankName}^^^{totalScoreItemName}";

                                                if (IntervalInfos[classRec.ClassID].ContainsKey(key))
                                                {

                                                    row[$"{totalScoreItemName}類別2_分組{number}頂標"] = IntervalInfos[classRec.ClassID][key].Avg_top_25;
                                                    row[$"{totalScoreItemName}類別2_分組{number}高標"] = IntervalInfos[classRec.ClassID][key].Avg_top_50;
                                                    row[$"{totalScoreItemName}類別2_分組{number}均標"] = IntervalInfos[classRec.ClassID][key].Avg;
                                                    row[$"{totalScoreItemName}類別2_分組{number}低標"] = IntervalInfos[classRec.ClassID][key].Avg_bottom_50;
                                                    row[$"{totalScoreItemName}類別2_分組{number}底標"] = IntervalInfos[classRec.ClassID][key].Avg_bottom_25;

                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count90"] = IntervalInfos[classRec.ClassID][key].Level_90;
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count80"] = IntervalInfos[classRec.ClassID][key].Level_80;
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count70"] = IntervalInfos[classRec.ClassID][key].Level_70;
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count60"] = IntervalInfos[classRec.ClassID][key].Level_60;
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count50"] = IntervalInfos[classRec.ClassID][key].Level_50;
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count40"] = IntervalInfos[classRec.ClassID][key].Level_40;
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count30"] = IntervalInfos[classRec.ClassID][key].Level_30;
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count20"] = IntervalInfos[classRec.ClassID][key].Level_20;
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count10"] = IntervalInfos[classRec.ClassID][key].Level_10;
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count100Up"] = IntervalInfos[classRec.ClassID][key].Level_gte100;
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count90Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[90];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count80Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[80];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count70Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[70];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count60Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[60];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count50Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[50];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count40Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[40];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count30Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[30];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count20Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[20];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count10Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[10];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count90Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[90];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count80Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[80];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count70Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[70];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count60Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[60];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count50Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[50];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count40Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[40];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count30Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[30];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count20Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[20];
                                                    row[$"{totalScoreItemName}類別2_分組{number}組距count10Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[10];
                                                    number++;
                                                    // subjectIndex++;
                                                }
                                                else
                                                {
                                                    Console.WriteLine(key + "");
                                                }
                                            }
                                        }
                                    }
                                }
                                #endregion 

                                #endregion

                                #endregion

                                #region 類別1綜合成績
                                ClassStuNum = 0;
                                foreach (StudentRecord stuRec in classRec.Students)
                                {
                                    ClassStuNum++;
                                    if (ClassStuNum > conf.StudentLimit)
                                        break;
                                    string studentID = stuRec.StudentID;
                                    // if (studentTag1Group.ContainsKey(studentID))
                                    {
                                        // todo
                                        if (studentTag1SubjectSum.ContainsKey(studentID))
                                        {
                                            row["類別1總分" + ClassStuNum] = studentTag1SubjectSum[studentID];

                                        }
                                        if (studentTag1SubjectAvg.ContainsKey(studentID))
                                        {
                                            row["類別1平均" + ClassStuNum] = studentTag1SubjectAvg[studentID];

                                        }
                                        if (studentTag1SubjectAvgW.ContainsKey(studentID))
                                        {
                                            row["類別1加權平均" + ClassStuNum] = studentTag1SubjectAvgW[studentID];

                                        }
                                        if (studentTag1SubjectSumW.ContainsKey(studentID))
                                        {
                                            row["類別1加權總分" + ClassStuNum] = studentTag1SubjectSumW[studentID];
                                        }

                                        //開始抓取排名

                                        if (studentTags.ContainsKey(stuRec.StudentID))
                                        {
                                            foreach (K12.Data.StudentTagRecord stuTag in studentTags[stuRec.StudentID])
                                            {
                                                try
                                                {
                                                    key = $"總計成績^^^類別1排名^^^grade{classRec.GradeYear}^^^{stuTag.Name}^^^總分";
                                                    row = PutRankValueToTable(dicFixedRankData, row, stuRec, "類別1", "總分", ClassStuNum, key);

                                                    key = $"總計成績^^^類別1排名^^^grade{classRec.GradeYear}^^^{stuTag.Name}^^^平均";
                                                    row = PutRankValueToTable(dicFixedRankData, row, stuRec, "類別1", "平均", ClassStuNum, key);

                                                    key = $"總計成績^^^類別1排名^^^grade{classRec.GradeYear}^^^{stuTag.Name}^^^加權總分";
                                                    row = PutRankValueToTable(dicFixedRankData, row, stuRec, "類別1", "加權總分", ClassStuNum, key);

                                                    key = $"總計成績^^^類別1排名^^^grade{classRec.GradeYear}^^^{stuTag.Name}^^^加權平均";
                                                    row = PutRankValueToTable(dicFixedRankData, row, stuRec, "類別1", "加權平均", ClassStuNum, key);

                                                }
                                                catch (Exception ex)
                                                {


                                                }
                                            }
                                        }

                                    }
                                }
                                for (int i = 0; i < tag1List.Count; i++)
                                {
                                    string tag = tag1List[i];
                                    #region 類別1總分組距及高低標

                                    #endregion
                                    #region 類別1平均組距及高低標

                                }
                                #endregion
                                #region 類別1加權總分組距及高低標

                                #endregion
                                #region 類別1加權平均組距及高低標

                                #endregion
                            }
                            #endregion

                            #region 類別2綜合成績
                            ClassStuNum = 0;
                            foreach (StudentRecord stuRec in classRec.Students)
                            {
                                ClassStuNum++;
                                if (ClassStuNum > conf.StudentLimit)
                                    break;
                                string studentID = stuRec.StudentID;
                                // 類別 2 總計成績

                                if (studentTag2SubjectSum.ContainsKey(studentID))
                                {
                                    row["類別2總分" + ClassStuNum] = studentTag2SubjectSum[studentID];

                                }
                                if (studentTag2SubjectAvg.ContainsKey(studentID))
                                {
                                    row["類別2平均" + ClassStuNum] = studentTag2SubjectAvg[studentID];

                                }
                                if (studentTag2SubjectAvgW.ContainsKey(studentID))
                                {
                                    row["類別2加權平均" + ClassStuNum] = studentTag2SubjectAvgW[studentID];

                                }
                                if (studentTag2SubjectSumW.ContainsKey(studentID))
                                {
                                    row["類別2加權總分" + ClassStuNum] = studentTag2SubjectSumW[studentID];
                                }








                                if (studentTags.ContainsKey(stuRec.StudentID))
                                {
                                    foreach (K12.Data.StudentTagRecord stuTag in studentTags[stuRec.StudentID])
                                    {
                                        try
                                        {
                                            key = $"總計成績^^^類別2排名^^^grade{classRec.GradeYear}^^^{stuTag.Name}^^^總分";
                                            row = PutRankValueToTable(dicFixedRankData, row, stuRec, "類別2", "總分", ClassStuNum, key);

                                            key = $"總計成績^^^類別2排名^^^grade{classRec.GradeYear}^^^{stuTag.Name}^^^平均";
                                            row = PutRankValueToTable(dicFixedRankData, row, stuRec, "類別2", "平均", ClassStuNum, key);

                                            key = $"總計成績^^^類別2排名^^^grade{classRec.GradeYear}^^^{stuTag.Name}^^^加權總分";
                                            row = PutRankValueToTable(dicFixedRankData, row, stuRec, "類別2", "加權總分", ClassStuNum, key);

                                            key = $"總計成績^^^類別2排名^^^grade{classRec.GradeYear}^^^{stuTag.Name}^^^加權平均";
                                            row = PutRankValueToTable(dicFixedRankData, row, stuRec, "類別2", "加權平均", ClassStuNum, key);
                                        }
                                        catch (Exception ex)
                                        {

                                        }
                                    }
                                }
                            }
                            for (int i = 0; i < tag2List.Count; i++)
                            {
                                string tag = tag2List[i];
                                #region 類別2總分組距及高低標
                                key = "類別2總分排名" + "^^^" + gradeYear + "^^^" + tag;
                                // if (rankStudents.ContainsKey(key))
                                {
                                    if (i == 0)
                                    {

                                    }
                                    else
                                    {

                                    }
                                }
                                #endregion
                                #region 類別2平均組距及高低標
                                key = "類別2平均排名" + "^^^" + gradeYear + "^^^" + tag;

                                #endregion
                                #region 類別2加權總分組距及高低標
                                key = "類別2加權總分排名" + "^^^" + gradeYear + "^^^" + tag;

                                #endregion
                                #region 類別2加權平均組距及高低標

                                #endregion
                            }
                            #endregion
                            //孟樺Jean 增加

                            table.Rows.Add(row);
                            progressCount++;
                            bkw.ReportProgress(70 + progressCount * 20 / selectedClasses.Count);

                        }
                        #endregion
                        bkw.ReportProgress(90);
                        document = conf.Template;


                        document.MailMerge.Execute(table);


                    }
                    catch (Exception exception)
                    {
                        exc = exception;
                    }

                    // aaa
                };

                bkw.RunWorkerAsync();
            }
        }
        /// <summary>
        ///  3.產生功能變數
        /// </summary>




        /// <summary>
        /// 總計成績 級距塞入
        /// </summary>
        /// <param name="dr">dataTable </param>
        /// <param name="itemName">總計成績OR科目</param>
        /// <param name="rankType">排名範圍(班、科、年)</param>
        /// <param name="key">Dictionary key 值</param>
        /// <returns></returns>
        private static DataRow PutRankIntervalToTable(Dictionary<string, Dictionary<string, FixRankIntervalInfo>> IntervalInfos, DataRow row, ClassRecord classRec, string itemName, string rankType, string key)
        {
            //總計成績班排名頂標

            if (IntervalInfos.ContainsKey(classRec.ClassID) && IntervalInfos[classRec.ClassID].ContainsKey(key))
            {
                row[$"{itemName}{rankType}頂標"] = IntervalInfos[classRec.ClassID][key].Avg_top_25;
                row[$"{itemName}{rankType}高標"] = IntervalInfos[classRec.ClassID][key].Avg_top_50;
                row[$"{itemName}{rankType}均標"] = IntervalInfos[classRec.ClassID][key].Avg;
                row[$"{itemName}{rankType}低標"] = IntervalInfos[classRec.ClassID][key].Avg_bottom_50;
                row[$"{itemName}{rankType}底標"] = IntervalInfos[classRec.ClassID][key].Avg_bottom_25;

                if (itemName != "總分" && itemName != "加權總分") //分數級距沒有意義
                {
                    row[$"{itemName}{rankType}組距count90"] = IntervalInfos[classRec.ClassID][key].Level_90;
                    row[$"{itemName}{rankType}組距count80"] = IntervalInfos[classRec.ClassID][key].Level_80;
                    row[$"{itemName}{rankType}組距count70"] = IntervalInfos[classRec.ClassID][key].Level_70;
                    row[$"{itemName}{rankType}組距count60"] = IntervalInfos[classRec.ClassID][key].Level_60;
                    row[$"{itemName}{rankType}組距count50"] = IntervalInfos[classRec.ClassID][key].Level_50;
                    row[$"{itemName}{rankType}組距count40"] = IntervalInfos[classRec.ClassID][key].Level_40;
                    row[$"{itemName}{rankType}組距count30"] = IntervalInfos[classRec.ClassID][key].Level_30;
                    row[$"{itemName}{rankType}組距count20"] = IntervalInfos[classRec.ClassID][key].Level_20;
                    row[$"{itemName}{rankType}組距count10"] = IntervalInfos[classRec.ClassID][key].Level_10;
                    row[$"{itemName}{rankType}組距count100Up"] = IntervalInfos[classRec.ClassID][key].Level_gte100;
                    row[$"{itemName}{rankType}組距count90Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[90];
                    row[$"{itemName}{rankType}組距count80Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[80];
                    row[$"{itemName}{rankType}組距count70Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[70];
                    row[$"{itemName}{rankType}組距count60Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[60];
                    row[$"{itemName}{rankType}組距count50Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[50];
                    row[$"{itemName}{rankType}組距count40Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[40];
                    row[$"{itemName}{rankType}組距count30Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[30];
                    row[$"{itemName}{rankType}組距count20Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[20];
                    row[$"{itemName}{rankType}組距count10Up"] = IntervalInfos[classRec.ClassID][key].DicCaculateUpResult[10];
                    row[$"{itemName}{rankType}組距count90Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[90];
                    row[$"{itemName}{rankType}組距count80Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[80];
                    row[$"{itemName}{rankType}組距count70Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[70];
                    row[$"{itemName}{rankType}組距count60Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[60];
                    row[$"{itemName}{rankType}組距count50Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[50];
                    row[$"{itemName}{rankType}組距count40Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[40];
                    row[$"{itemName}{rankType}組距count30Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[30];
                    row[$"{itemName}{rankType}組距count20Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[20];
                    row[$"{itemName}{rankType}組距count10Down"] = IntervalInfos[classRec.ClassID][key].DicCaculateDoownResult[10];
                }

                if (itemName == "總分" || itemName == "加權總分")
                {
                    row[$"{itemName}{rankType}組距count90"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count80"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count70"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count60"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count50"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count40"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count30"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count20"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count10"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count100Up"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count90Up"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count80Up"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count70Up"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count60Up"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count50Up"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count40Up"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count30Up"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count20Up"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count10Up"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count90Down"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count80Down"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count70Down"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count60Down"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count50Down"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count40Down"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count30Down"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count20Down"] = "總分不計算";
                    row[$"{itemName}{rankType}組距count10Down"] = "總分不計算";
                }
            }
            return row;
        }



        private static DataRow PutRankValueToTable(Dictionary<string, Dictionary<string, StudentFixedRankInfo>> dicFixedRankData, DataRow row, StudentRecord stuRec, string itemName, string rankType, int ClassStuNum, string key)
        {

            if (dicFixedRankData.ContainsKey(stuRec.RefClass.ClassID) && dicFixedRankData[stuRec.RefClass.ClassID].ContainsKey(stuRec.StudentID) && dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectTotalFixRank.ContainsKey(key))
            {
                // todo 
                try
                {
                    row[$"{itemName}{rankType}排名母數" + ClassStuNum] = dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectTotalFixRank[key].MatrixCount;
                    row[$"{itemName}{rankType}排名" + ClassStuNum] = dicFixedRankData[stuRec.RefClass.ClassID][stuRec.StudentID].DicSubjectTotalFixRank[key].Rank;
                }
                catch (Exception ex)
                {
                    Console.WriteLine("錯誤");
                }
            }
            return row;
        }
    }
}


