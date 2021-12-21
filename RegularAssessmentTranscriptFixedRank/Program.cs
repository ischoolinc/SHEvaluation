using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.IO;
using SmartSchool.Customization.Data;
using System.Threading;
using SmartSchool.Customization.Data.StudentExtension;
using FISCA.Data;
using FISCA.Permission;
using SmartSchool;
using Campus.ePaperCloud;

namespace RegularAssessmentTranscriptFixedRank
{
    public class Program
    {
        public static int AvgRd = 2;
        [FISCA.MainMethod]
        public static void Main()
        {
            var btn = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["定期評量成績單(固定排名)"];
            btn.Enable = false;
            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                btn.Enable = Permissions.定期評量成績單固定排名權限 && (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0);
            };
            btn.Click += new EventHandler(Program_Click);

            //權限設定
            Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.定期評量成績單固定排名, "定期評量成績單(固定排名)"));
        }

        // 學生清單暫存
        private static Aspose.Cells.Workbook _wbStudentList;

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

        /// <summary>
        /// 四捨五入至使用者指定位數
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string ParseScore(string str)
        {
            string value = "";
            decimal dc;
            if (decimal.TryParse(str, out dc))
            {
                value = Math.Round(dc, AvgRd, MidpointRounding.AwayFromZero).ToString();
            }

            return value;
        }

        static void Program_Click(object sender_, EventArgs e_)
        {
            ConfigForm form = new ConfigForm();
            if (form.ShowDialog() == DialogResult.OK)
            {
                AccessHelper accessHelper = new AccessHelper();
                //return;
                List<StudentRecord> overflowRecords = new List<StudentRecord>();
                //取得列印設定
                Configure conf = form.Configure;
                //建立測試的選取學生(先期不管怎麼選就是印這些人)
                List<string> selectedStudents = K12.Presentation.NLDPanels.Student.SelectedSource;
                //建立合併欄位總表
                DataTable table = new DataTable();
                #region 所有的合併欄位
                table.Columns.Add("學校名稱");
                table.Columns.Add("學校地址");
                table.Columns.Add("學校電話");
                table.Columns.Add("收件人地址");
                //«通訊地址»«通訊地址郵遞區號»«通訊地址內容»
                //«戶籍地址»«戶籍地址郵遞區號»«戶籍地址內容»
                //«監護人»«父親»«母親»«科別名稱»
                table.Columns.Add("通訊地址");
                table.Columns.Add("通訊地址郵遞區號");
                table.Columns.Add("通訊地址內容");
                table.Columns.Add("戶籍地址");
                table.Columns.Add("戶籍地址郵遞區號");
                table.Columns.Add("戶籍地址內容");
                table.Columns.Add("監護人");
                table.Columns.Add("父親");
                table.Columns.Add("母親");
                table.Columns.Add("科別名稱");
                table.Columns.Add("試別");
                table.Columns.Add("家長代碼");

                table.Columns.Add("收件人");
                table.Columns.Add("學年度");
                table.Columns.Add("學期");
                table.Columns.Add("班級科別名稱");
                table.Columns.Add("班級");
                table.Columns.Add("班導師");
                table.Columns.Add("座號");
                table.Columns.Add("學號");
                table.Columns.Add("姓名");
                table.Columns.Add("系統編號");
                table.Columns.Add("定期評量");
                table.Columns.Add("類別1名稱");
                table.Columns.Add("類別2名稱");

                // 新增合併欄位
                List<string> r1List = new List<string>();
                List<string> r2List = new List<string>();
                List<string> r3List = new List<string>();

                r1List.Add("總分班排名");
                r1List.Add("總分科排名");
                r1List.Add("總分全校排名");
                r1List.Add("平均班排名");
                r1List.Add("平均科排名");
                r1List.Add("平均全校排名");
                r1List.Add("加權總分班排名");
                r1List.Add("加權總分科排名");
                r1List.Add("加權總分全校排名");
                r1List.Add("加權平均班排名");
                r1List.Add("加權平均科排名");
                r1List.Add("加權平均全校排名");
                r1List.Add("類別1總分排名");
                r1List.Add("類別1平均排名");
                r1List.Add("類別1加權總分排名");
                r1List.Add("類別1加權平均排名");
                r1List.Add("類別2總分排名");
                r1List.Add("類別2平均排名");
                r1List.Add("類別2加權總分排名");
                r1List.Add("類別2加權平均排名");

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
                r2List.Add("level_60up");
                r2List.Add("level_60down");
                r2List.Add("std_dev_pop");
                r2List.Add("pr_88");
                r2List.Add("pr_75");
                r2List.Add("pr_50");
                r2List.Add("pr_25");
                r2List.Add("pr_12");


                r3List.Add("班排名");
                r3List.Add("科排名");
                r3List.Add("全校排名");
                r3List.Add("類別1排名");
                r3List.Add("類別2排名");

                foreach (string r1 in r1List)
                {
                    foreach (string r2 in r2List)
                    {
                        string it = r1 + "_" + r2;
                        table.Columns.Add(it);
                    }
                }

                for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
                {
                    table.Columns.Add("科目名稱" + subjectIndex);
                    table.Columns.Add("學分數" + subjectIndex);
                    table.Columns.Add("前次成績" + subjectIndex);
                    table.Columns.Add("科目成績" + subjectIndex);
                    table.Columns.Add("班排名" + subjectIndex);
                    table.Columns.Add("班排名母數" + subjectIndex);
                    table.Columns.Add("科排名" + subjectIndex);
                    table.Columns.Add("科排名母數" + subjectIndex);
                    table.Columns.Add("類別1排名" + subjectIndex);
                    table.Columns.Add("類別1排名母數" + subjectIndex);
                    table.Columns.Add("類別2排名" + subjectIndex);
                    table.Columns.Add("類別2排名母數" + subjectIndex);
                    table.Columns.Add("全校排名" + subjectIndex);
                    table.Columns.Add("全校排名母數" + subjectIndex);

                    foreach (string r3 in r3List)
                    {
                        foreach (string r2 in r2List)
                        {
                            string it = r3 + subjectIndex + "_" + r2;
                            table.Columns.Add(it);
                        }
                    }

                }
                table.Columns.Add("總分");
                table.Columns.Add("總分班排名");
                table.Columns.Add("總分班排名母數");
                table.Columns.Add("總分科排名");
                table.Columns.Add("總分科排名母數");
                table.Columns.Add("總分全校排名");
                table.Columns.Add("總分全校排名母數");
                table.Columns.Add("平均");
                table.Columns.Add("平均班排名");
                table.Columns.Add("平均班排名母數");
                table.Columns.Add("平均科排名");
                table.Columns.Add("平均科排名母數");
                table.Columns.Add("平均全校排名");
                table.Columns.Add("平均全校排名母數");
                table.Columns.Add("加權總分");
                table.Columns.Add("加權總分班排名");
                table.Columns.Add("加權總分班排名母數");
                table.Columns.Add("加權總分科排名");
                table.Columns.Add("加權總分科排名母數");
                table.Columns.Add("加權總分全校排名");
                table.Columns.Add("加權總分全校排名母數");
                table.Columns.Add("加權平均");
                table.Columns.Add("加權平均班排名");
                table.Columns.Add("加權平均班排名母數");
                table.Columns.Add("加權平均科排名");
                table.Columns.Add("加權平均科排名母數");
                table.Columns.Add("加權平均全校排名");
                table.Columns.Add("加權平均全校排名母數");

                table.Columns.Add("類別排名1");
                table.Columns.Add("類別1總分");
                table.Columns.Add("類別1總分排名");
                table.Columns.Add("類別1總分排名母數");
                table.Columns.Add("類別1平均");
                table.Columns.Add("類別1平均排名");
                table.Columns.Add("類別1平均排名母數");
                table.Columns.Add("類別1加權總分");
                table.Columns.Add("類別1加權總分排名");
                table.Columns.Add("類別1加權總分排名母數");
                table.Columns.Add("類別1加權平均");
                table.Columns.Add("類別1加權平均排名");
                table.Columns.Add("類別1加權平均排名母數");

                table.Columns.Add("類別排名2");
                table.Columns.Add("類別2總分");
                table.Columns.Add("類別2總分排名");
                table.Columns.Add("類別2總分排名母數");
                table.Columns.Add("類別2平均");
                table.Columns.Add("類別2平均排名");
                table.Columns.Add("類別2平均排名母數");
                table.Columns.Add("類別2加權總分");
                table.Columns.Add("類別2加權總分排名");
                table.Columns.Add("類別2加權總分排名母數");
                table.Columns.Add("類別2加權平均");
                table.Columns.Add("類別2加權平均排名");
                table.Columns.Add("類別2加權平均排名母數");
                // 獎懲統計 --
                table.Columns.Add("大功統計");
                table.Columns.Add("小功統計");
                table.Columns.Add("嘉獎統計");
                table.Columns.Add("大過統計");
                table.Columns.Add("小過統計");
                table.Columns.Add("警告統計");
                table.Columns.Add("留校察看");

                table.Columns.Add("成績校正年");
                table.Columns.Add("成績校正月");
                table.Columns.Add("成績校正日");
                table.Columns.Add("成績校正民國年");
                table.Columns.Add("成績校正日期");

                // 新增 columns name
                table.Columns.Add("開始日期");
                table.Columns.Add("結束日期");

                // 先固定8個
                for (int i = 1; i <= 8; i++)
                    table.Columns.Add("學習服務區間時數" + i);

                table.Columns.Add("大功區間統計");
                table.Columns.Add("小功區間統計");
                table.Columns.Add("嘉獎區間統計");
                table.Columns.Add("大過區間統計");
                table.Columns.Add("小過區間統計");
                table.Columns.Add("警告區間統計");
                table.Columns.Add("留校察看區間");

                // 動態新增缺曠統計，使用模式一般_曠課、一般_事假..
                foreach (string name in Utility.GetATMappingKey())
                    table.Columns.Add("區間" + name);

                // 動態資料新增
                for (int atIdx = 1; atIdx <= conf.AttendanceDetailLimit; atIdx++)
                {
                    // 缺曠區間明細 A:日期,B:
                    table.Columns.Add("缺曠區間明細日期" + atIdx);
                    table.Columns.Add("缺曠區間明細內容" + atIdx);
                    table.Columns.Add("缺曠區間明細C" + atIdx);
                }

                // 獎懲區間明細 A:日期,B:類別支數,C:事由
                for (int atIdx = 1; atIdx <= conf.DisciplineDetailLimit; atIdx++)
                {
                    table.Columns.Add("獎懲區間明細日期" + atIdx);
                    table.Columns.Add("獎懲區間明細類別支數" + atIdx);
                    table.Columns.Add("獎懲區間明細事由" + atIdx);
                }
                for (int atIdx = 1; atIdx <= conf.ServiceLearningDetailLimit; atIdx++)
                {
                    // 學習服務區間明細 A:日期,B:內容,C:時數
                    table.Columns.Add("學習服務區間明細日期" + atIdx);
                    table.Columns.Add("學習服務區間明細內容" + atIdx);
                    table.Columns.Add("學習服務區間明細時數" + atIdx);
                }

                // 新增服務學習欄位
                table.Columns.Add("學習服務區間加總");
                table.Columns.Add("學習服務當學期加總");

                // 標頭
                table.TableName = "ss";
                table.WriteXmlSchema(Application.StartupPath + "\\sS.xml");

                #endregion
                //宣告產生的報表
                Aspose.Words.Document document = new Aspose.Words.Document();
                //用一個BackgroundWorker包起來
                System.ComponentModel.BackgroundWorker bkw = new System.ComponentModel.BackgroundWorker();
                bkw.WorkerReportsProgress = true;
                System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 學生定期評量成績單(固定排名)產生 S");
                bkw.ProgressChanged += delegate (object sender, System.ComponentModel.ProgressChangedEventArgs e)
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("學生定期評量成績單(固定排名)產生中", e.ProgressPercentage);
                    System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 學生定期評量成績單(固定排名)產生 " + e.ProgressPercentage);
                };
                Exception exc = null;
                bkw.RunWorkerCompleted += delegate
                {
                    //#region 將 DataTable 內合併欄位產生出來
                    //StreamWriter sw = new StreamWriter(Application.StartupPath + "\\定期評量成績單合併欄位.txt");
                    //foreach (DataColumn dc in table.Columns)
                    //    sw.WriteLine(dc.Caption);

                    //sw.Close();
                    //#endregion


                    System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 學生定期評量成績單(固定排名)產生 E");
                    string err = "下列學生因成績項目超過樣板支援上限，\n超出部分科目成績無法印出，建議調整樣板內容。";
                    if (overflowRecords.Count > 0)
                    {
                        foreach (var stuRec in overflowRecords)
                        {
                            err += "\n" + (stuRec.RefClass == null ? "" : (stuRec.RefClass.ClassName + "班" + stuRec.SeatNo + "號")) + "[" + stuRec.StudentNumber + "]" + stuRec.StudentName;
                        }
                    }
                    #region 儲存檔案
                    string inputReportName = "定期評量成績單(固定排名)";
                    string reportName = inputReportName;

                    try
                    {
                        K12.Data.ExamRecord examRecord = conf.ExamRecord;
                        string newReportName = conf.SchoolYear + "學年度第" + conf.Semester + "學期" + examRecord.Name + "" + reportName;
                        MemoryStream memoryStream = new MemoryStream();
                        document.Save(memoryStream, Aspose.Words.SaveFormat.Doc);
                        ePaperCloud ePaperCloud = new ePaperCloud();
                        ePaperCloud.upload_ePaper(Convert.ToInt32(conf.SchoolYear), Convert.ToInt32(conf.Semester), newReportName, "", memoryStream, ePaperCloud.ViewerType.Student, ePaperCloud.FormatType.Docx);
                    }
                    catch (Exception ex)
                    {
                        throw ex;
                    }
                    #endregion

                    // 檢查是否需要產生學生清單
                    if (form.GetisExportStudentList())
                    {
                        string ExportReportName = "定期評量成績單(固定排名)學生清單";

                        string pathxls = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                        if (!Directory.Exists(pathxls))
                            Directory.CreateDirectory(pathxls);
                        pathxls = Path.Combine(pathxls, ExportReportName + ".xls");

                        if (File.Exists(pathxls))
                        {
                            int i = 1;
                            while (true)
                            {
                                string newPath = Path.GetDirectoryName(pathxls) + "\\" + Path.GetFileNameWithoutExtension(pathxls) + (i++) + Path.GetExtension(pathxls);
                                if (!File.Exists(newPath))
                                {
                                    pathxls = newPath;
                                    break;
                                }
                            }
                        }

                        try
                        {
                            _wbStudentList.Save(pathxls, Aspose.Cells.FileFormatType.Excel2003);
                            System.Diagnostics.Process.Start(pathxls);
                        }
                        catch
                        {
                            System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                            sd.Title = "另存新檔";
                            sd.FileName = ExportReportName + ".xls";
                            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                            {
                                try
                                {
                                    _wbStudentList.Save(sd.FileName, Aspose.Cells.FileFormatType.Excel2003);


                                }
                                catch
                                {
                                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                    return;
                                }
                            }
                        }
                    }

                    FISCA.Presentation.MotherForm.SetStatusBarMessage("學生定期評量成績單(固定排名)產生完成。", 100);
                    if (overflowRecords.Count > 0)
                        MessageBox.Show(err);
                    if (exc != null)
                    {
                        throw new Exception("產生學生定期評量成績單(固定排名)發生錯誤", exc);
                    }
                };
                bkw.DoWork += delegate (object sender, System.ComponentModel.DoWorkEventArgs e)
                {
                    //取得家長代碼
                    QueryHelper q = new QueryHelper();
                    string ids = string.Join("','", selectedStudents);

                    Dictionary<string, string> parentCode = new Dictionary<string, string>();

                    DataTable response = q.Select("SELECT id,parent_code FROM student WHERE id IN ('" + ids + "')");
                    foreach (DataRow row in response.Rows)
                    {
                        string id = row["id"].ToString();
                        string code = row["parent_code"].ToString();

                        if (!parentCode.ContainsKey(id))
                            parentCode.Add(id, code);
                    }

                    var studentRecords = accessHelper.StudentHelper.GetStudents(selectedStudents);
                    Dictionary<string, Dictionary<string, Dictionary<string, ExamScoreInfo>>> studentExamSores = new Dictionary<string, Dictionary<string, Dictionary<string, ExamScoreInfo>>>();
                    Dictionary<string, Dictionary<string, ExamScoreInfo>> studentRefExamSores = new Dictionary<string, Dictionary<string, ExamScoreInfo>>();
                    ManualResetEvent scoreReady = new ManualResetEvent(false);
                    ManualResetEvent elseReady = new ManualResetEvent(false);
                    #region 偷跑取得考試成績
                    new Thread(new ThreadStart(delegate
                    {
                        // 取得學生學期科目成績
                        int sSchoolYear, sSemester;
                        int.TryParse(conf.SchoolYear, out sSchoolYear);
                        int.TryParse(conf.Semester, out sSemester);
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
                                    targetCourseList.Add(courseRecord);
                            }
                        }
                        catch (Exception exception)
                        {
                            exc = exception;
                        }
                        #endregion
                        try
                        {
                            if (conf.ExamRecord != null || conf.RefenceExamRecord != null)
                            {
                                accessHelper.CourseHelper.FillExam(targetCourseList);
                                var tcList = new List<CourseRecord>();
                                var totalList = new List<CourseRecord>();
                                foreach (var courseRec in targetCourseList)
                                {
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
                                        foreach (var examScoreRec in courseRecord.ExamScoreList)
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
                            accessHelper.StudentHelper.FillAttendance(studentRecords);
                            accessHelper.StudentHelper.FillReward(studentRecords);
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
                        bkw.ReportProgress(0);
                        #region 缺曠對照表
                        List<K12.Data.PeriodMappingInfo> periodMappingInfos = K12.Data.PeriodMapping.SelectAll();
                        Dictionary<string, string> dicPeriodMappingType = new Dictionary<string, string>();
                        List<string> periodTypes = new List<string>();
                        foreach (K12.Data.PeriodMappingInfo periodMappingInfo in periodMappingInfos)
                        {
                            if (!dicPeriodMappingType.ContainsKey(periodMappingInfo.Name))
                                dicPeriodMappingType.Add(periodMappingInfo.Name, periodMappingInfo.Type);

                            if (!periodTypes.Contains(periodMappingInfo.Type))
                                periodTypes.Add(periodMappingInfo.Type);
                        }
                        foreach (var absence in K12.Data.AbsenceMapping.SelectAll())
                        {
                            foreach (var pt in periodTypes)
                            {
                                string attendanceKey = pt + "_" + absence.Name;
                                if (!table.Columns.Contains(attendanceKey))
                                {
                                    table.Columns.Add(attendanceKey);
                                }
                            }
                        }
                        #endregion
                        bkw.ReportProgress(3);
                        #region 整理學生住址
                        accessHelper.StudentHelper.FillContactInfo(studentRecords);
                        #endregion
                        #region 整理學生父母及監護人
                        accessHelper.StudentHelper.FillParentInfo(studentRecords);
                        #endregion
                        bkw.ReportProgress(10);
                        #region 整理同年級學生
                        //整理選取學生的年級
                        Dictionary<string, List<StudentRecord>> gradeyearStudents = new Dictionary<string, List<StudentRecord>>();
                        foreach (var studentRec in studentRecords)
                        {
                            string grade = "";
                            if (studentRec.RefClass != null)
                                grade = "" + studentRec.RefClass.GradeYear;
                            if (!gradeyearStudents.ContainsKey(grade))
                                gradeyearStudents.Add(grade, new List<StudentRecord>());
                            gradeyearStudents[grade].Add(studentRec);
                        }
                        foreach (var classRec in accessHelper.ClassHelper.GetAllClass())
                        {
                            if (gradeyearStudents.ContainsKey("" + classRec.GradeYear))
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

                        #region 計算總分及各項目排名
                        Dictionary<string, string> studentTag1Group = new Dictionary<string, string>();
                        Dictionary<string, string> studentTag2Group = new Dictionary<string, string>();
                        Dictionary<string, List<decimal>> ranks = new Dictionary<string, List<decimal>>();
                        Dictionary<string, List<string>> rankStudents = new Dictionary<string, List<string>>();
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
                        foreach (string gradeyear in gradeyearStudents.Keys)
                        {
                            //找出全年級學生
                            foreach (var studentRec in gradeyearStudents[gradeyear])
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
                                        if (conf.RankFilterTagList.Contains(tag.RefTagID))
                                        {
                                            rank = false;
                                        }
                                        #endregion
                                        #region 判斷學生在類別排名1中所屬的類別
                                        if (tag1ID == "" && conf.TagRank1TagList.Contains(tag.RefTagID))
                                        {
                                            tag1ID = tag.RefTagID;
                                            studentTag1Group.Add(studentID, tag1ID);
                                        }
                                        #endregion
                                        #region 判斷學生在類別排名2中所屬的類別
                                        if (tag2ID == "" && conf.TagRank2TagList.Contains(tag.RefTagID))
                                        {
                                            tag2ID = tag.RefTagID;
                                            studentTag2Group.Add(studentID, tag2ID);
                                        }
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
                                                    //計算加權總分 - 2014/10/9 改為decimal
                                                    printSubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                    printSubjectCreditSum += sceTakeRecord.CreditDec();
                                                }
                                                else
                                                {
                                                    summaryRank = false;
                                                }
                                            }
                                            #endregion
                                        }
                                        if (tag1ID != "" && conf.TagRank1SubjectList.Contains(subjectName))
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
                                                }
                                                else
                                                {
                                                    tag1SummaryRank = false;
                                                }
                                            }
                                            #endregion
                                        }
                                        if (tag2ID != "" && conf.TagRank2SubjectList.Contains(subjectName))
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
                                                }
                                                else
                                                {
                                                    tag2SummaryRank = false;
                                                }
                                            }
                                            #endregion
                                        }

                                    }
                                    if (printSubjectCount > 0)
                                    {
                                        #region 有列印科目處理加總成績
                                        //總分
                                        studentPrintSubjectSum.Add(studentID, printSubjectSum);
                                        //平均四捨五入至小數點第二位
                                        studentPrintSubjectAvg.Add(studentID, Math.Round(printSubjectSum / printSubjectCount, AvgRd, MidpointRounding.AwayFromZero));

                                        #endregion
                                        if (printSubjectCreditSum > 0)
                                        {
                                            #region 有總學分數處理加總
                                            //加權總分
                                            studentPrintSubjectSumW.Add(studentID, printSubjectSumW);
                                            //加權平均四捨五入至小數點第二位
                                            studentPrintSubjectAvgW.Add(studentID, Math.Round(printSubjectSumW / printSubjectCreditSum, AvgRd, MidpointRounding.AwayFromZero));

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

                                    }
                                    //類別2總分平均排名
                                    if (tag2SubjectCount > 0)
                                    {
                                        //總分
                                        studentTag2SubjectSum.Add(studentID, tag2SubjectSum);
                                        //平均四捨五入至小數點第二位
                                        studentTag2SubjectAvg.Add(studentID, Math.Round(tag2SubjectSum / tag2SubjectCount, AvgRd, MidpointRounding.AwayFromZero));

                                    }
                                }
                                progressCount++;
                                bkw.ReportProgress(40 + progressCount * 30 / total);
                            }
                        }
                        #endregion


                        // 先取得 K12 StudentRec,因為後面透過 k12.data 取資料有的傳入ID,有的傳入 Record 有點亂
                        List<K12.Data.StudentRecord> StudRecList = new List<K12.Data.StudentRecord>();
                        List<string> StudIDList = (from data in studentRecords select data.StudentID).ToList();
                        StudRecList = K12.Data.Student.SelectByIDs(StudIDList);

                        // 取得暫存資料 學習服務區間時數
                        Dictionary<string, Dictionary<string, decimal>> ServiceLearningByDateDict = Utility.GetServiceLearningByDate(StudIDList, form.GetBeginDate(), form.GetEndDate());
                        Dictionary<string, List<DataRow>> ServiceLearningDetailByDateDict = Utility.GetServiceLearningDetailByDate(StudIDList, form.GetBeginDate(), form.GetEndDate());

                        // 服務學習
                        Utility.SetServiceLearningSum(StudIDList, form.GetBeginDate(), form.GetEndDate(), conf.SchoolYear, conf.Semester);

                        Dictionary<string, decimal> ServiceLearningByDateDictSum = Utility.GetServiceLearningByDateDict();

                        Dictionary<string, decimal> ServiceLearningBySemesterDictSum = Utility.GetServiceLearningBySemesterDict();

                        // 取得排名、五標、分數區間
                        Dictionary<string, Dictionary<string, DataRow>> RankMatrixDataDict = Utility.GetRankMatrixData(conf.SchoolYear, conf.Semester, conf.ExamRecord.ID, StudIDList);


                        // 取得缺曠
                        Dictionary<string, Dictionary<string, int>> AttendanceCountDict = Utility.GetAttendanceCountByDate(StudRecList, form.GetBeginDate(), form.GetEndDate());
                        Dictionary<string, List<K12.Data.AttendanceRecord>> AttendanceDetailDict = Utility.GetAttendanceDetailByDate(StudRecList, form.GetBeginDate(), form.GetEndDate());

                        // 取得獎懲
                        Dictionary<string, Dictionary<string, int>> DisciplineCountDict = Utility.GetDisciplineCountByDate(StudIDList, form.GetBeginDate(), form.GetEndDate());
                        Dictionary<string, List<K12.Data.DisciplineRecord>> DisciplinedetailDict = Utility.GetDisciplineDetailByDate(StudIDList, form.GetBeginDate(), form.GetEndDate());

                        List<K12.Data.PeriodMappingInfo> PeriodMappingList = K12.Data.PeriodMapping.SelectAll();
                        // 節次>類別
                        Dictionary<string, string> PeriodMappingDict = new Dictionary<string, string>();
                        foreach (K12.Data.PeriodMappingInfo rec in PeriodMappingList)
                        {
                            if (!PeriodMappingDict.ContainsKey(rec.Name))
                                PeriodMappingDict.Add(rec.Name, rec.Type);
                        }

                        bkw.ReportProgress(70);
                        elseReady.WaitOne();
                        progressCount = 0;
                        #region 填入資料表
                        foreach (var stuRec in studentRecords)
                        {
                            string studentID = stuRec.StudentID;
                            string gradeYear = (stuRec.RefClass == null ? "" : "" + stuRec.RefClass.GradeYear);
                            DataRow row = table.NewRow();

                            // 這區段是新增功能資料
                            // 畫面上開始結束日期
                            row["開始日期"] = form.GetBeginDate().ToShortDateString();
                            row["結束日期"] = form.GetEndDate().ToShortDateString();

                            if (ServiceLearningByDateDict.ContainsKey(studentID))
                            {
                                // 處理學生學習服務時數
                                int idx = 1;
                                foreach (KeyValuePair<string, decimal> data in ServiceLearningByDateDict[studentID])
                                {
                                    if (idx < 9)
                                    {
                                        row["學習服務區間時數" + idx] = data.Key + " 時數：" + data.Value;
                                    }
                                    idx++;
                                }
                            }
                            // 處理獎懲
                            if (DisciplineCountDict.ContainsKey(studentID))
                            {
                                foreach (KeyValuePair<string, int> data in DisciplineCountDict[studentID])
                                {
                                    switch (data.Key)
                                    {
                                        case "大功": row["大功區間統計"] = data.Value; break;
                                        case "小功": row["小功區間統計"] = data.Value; break;
                                        case "嘉獎": row["嘉獎區間統計"] = data.Value; break;
                                        case "大過": row["大過區間統計"] = data.Value; break;
                                        case "小過": row["小過區間統計"] = data.Value; break;
                                        case "警告": row["警告區間統計"] = data.Value; break;

                                        case "留校":
                                            if (data.Value > 0)
                                                row["留校察看區間"] = "是";
                                            else
                                                row["留校察看區間"] = "";
                                            break;
                                    }
                                }
                            }

                            // 處理缺曠
                            if (AttendanceCountDict.ContainsKey(studentID))
                            {
                                foreach (KeyValuePair<string, int> data in AttendanceCountDict[studentID])
                                {
                                    if (table.Columns.Contains(data.Key))
                                        row[data.Key] = data.Value;
                                }
                            }

                            // 處理缺曠區間明細
                            if (AttendanceDetailDict.ContainsKey(studentID))
                            {
                                int idx = 1;
                                foreach (K12.Data.AttendanceRecord rec in AttendanceDetailDict[studentID])
                                {

                                    foreach (K12.Data.AttendancePeriod per in rec.PeriodDetail)
                                    {
                                        if (PeriodMappingDict.ContainsKey(per.Period))
                                        {
                                            if (idx <= conf.AttendanceDetailLimit)
                                            {

                                                row["缺曠區間明細日期" + idx] = rec.OccurDate.ToShortDateString();
                                                row["缺曠區間明細內容" + idx] = PeriodMappingDict[per.Period] + ":" + per.AbsenceType + " (節次：" + per.Period + ")";
                                                //row["缺曠區間明細C" + idx] = rec.                                        
                                                idx++;
                                            }
                                        }
                                    }

                                }


                            }

                            // 處理獎懲區間明細
                            if (DisciplinedetailDict.ContainsKey(studentID))
                            {
                                int idx = 1;

                                foreach (K12.Data.DisciplineRecord data in DisciplinedetailDict[studentID])
                                {
                                    if (idx <= conf.DisciplineDetailLimit)
                                    {
                                        // 獎懲區間明細 A:日期,B:類別支數,C:事由
                                        row["獎懲區間明細日期" + idx] = data.OccurDate.ToShortDateString();

                                        List<string> strTypeList = new List<string>();
                                        if (data.MeritFlag == "1")
                                        {
                                            if (data.MeritA.HasValue)
                                                strTypeList.Add("大功：" + data.MeritA.Value);

                                            if (data.MeritB.HasValue)
                                                strTypeList.Add("小功：" + data.MeritB.Value);

                                            if (data.MeritC.HasValue)
                                                strTypeList.Add("嘉獎：" + data.MeritC.Value);

                                        }
                                        if (data.MeritFlag == "0")
                                        {
                                            if (data.Cleared != "是")
                                            {
                                                if (data.DemeritA.HasValue)
                                                    strTypeList.Add("大過：" + data.DemeritA.Value);

                                                if (data.DemeritB.HasValue)
                                                    strTypeList.Add("小過：" + data.DemeritB.Value);

                                                if (data.DemeritC.HasValue)
                                                    strTypeList.Add("警告：" + data.DemeritC.Value);
                                            }
                                        }
                                        if (data.MeritFlag == "2")
                                            strTypeList.Add("留校察看：是");

                                        row["獎懲區間明細類別支數" + idx] = string.Join(",", strTypeList.ToArray());
                                        row["獎懲區間明細事由" + idx] = data.Reason;
                                        idx++;
                                    }
                                }

                            }

                            // 處理學習服務
                            if (ServiceLearningDetailByDateDict.ContainsKey(studentID))
                            {
                                int idx = 1;
                                foreach (DataRow dr in ServiceLearningDetailByDateDict[studentID])
                                {
                                    if (idx <= conf.ServiceLearningDetailLimit)
                                    {
                                        //  學習服務區間明細 A:日期,B:內容,C:時數
                                        DateTime dt;
                                        decimal hr;
                                        string cont = dr["reason"].ToString();

                                        if (DateTime.TryParse(dr["occur_date"].ToString(), out dt))
                                            row["學習服務區間明細日期" + idx] = dt.ToShortDateString();
                                        else
                                            row["學習服務區間明細日期" + idx] = "";
                                        row["學習服務區間明細內容" + idx] = cont;

                                        if (decimal.TryParse(dr["hours"].ToString(), out hr))
                                            row["學習服務區間明細時數" + idx] = hr;
                                        else
                                            row["學習服務區間明細時數" + idx] = "";
                                        idx++;
                                    }
                                }
                            }

                            // 新增服務學習欄位
                            if (ServiceLearningByDateDictSum.ContainsKey(studentID))
                            {
                                row["學習服務區間加總"] = ServiceLearningByDateDictSum[studentID];
                            }

                            if (ServiceLearningBySemesterDictSum.ContainsKey(studentID))
                            {
                                row["學習服務當學期加總"] = ServiceLearningBySemesterDictSum[studentID];
                            }

                            #region 基本資料
                            row["學校名稱"] = SmartSchool.Customization.Data.SystemInformation.SchoolChineseName;
                            row["學校地址"] = SmartSchool.Customization.Data.SystemInformation.Address;
                            row["學校電話"] = SmartSchool.Customization.Data.SystemInformation.Telephone;
                            row["收件人地址"] = stuRec.ContactInfo.MailingAddress.FullAddress != "" ?
                                                    stuRec.ContactInfo.MailingAddress.FullAddress : stuRec.ContactInfo.PermanentAddress.FullAddress;
                            row["收件人"] = stuRec.ParentInfo.CustodianName != "" ? stuRec.ParentInfo.CustodianName :
                                                    (stuRec.ParentInfo.FatherName != "" ? stuRec.ParentInfo.FatherName :
                                                        (stuRec.ParentInfo.FatherName != "" ? stuRec.ParentInfo.MotherName : stuRec.StudentName));
                            //«通訊地址»«通訊地址郵遞區號»«通訊地址內容»
                            //«戶籍地址»«戶籍地址郵遞區號»«戶籍地址內容»
                            //«監護人»«父親»«母親»«科別名稱»
                            row["通訊地址"] = stuRec.ContactInfo.MailingAddress.FullAddress;
                            row["通訊地址郵遞區號"] = stuRec.ContactInfo.MailingAddress.ZipCode;
                            row["通訊地址內容"] = stuRec.ContactInfo.MailingAddress.County + stuRec.ContactInfo.MailingAddress.Town + stuRec.ContactInfo.MailingAddress.DetailAddress;
                            row["戶籍地址"] = stuRec.ContactInfo.PermanentAddress.FullAddress;
                            row["戶籍地址郵遞區號"] = stuRec.ContactInfo.PermanentAddress.ZipCode;
                            row["戶籍地址內容"] = stuRec.ContactInfo.PermanentAddress.County + stuRec.ContactInfo.PermanentAddress.Town + stuRec.ContactInfo.PermanentAddress.DetailAddress;
                            row["監護人"] = stuRec.ParentInfo.CustodianName;
                            row["父親"] = stuRec.ParentInfo.FatherName;
                            row["母親"] = stuRec.ParentInfo.MotherName;
                            row["科別名稱"] = stuRec.Department;
                            row["試別"] = conf.ExamRecord.Name;
                            row["家長代碼"] = parentCode.ContainsKey(studentID) ? parentCode[studentID] : "";

                            row["學年度"] = conf.SchoolYear;
                            row["學期"] = conf.Semester;
                            row["班級科別名稱"] = stuRec.RefClass == null ? "" : stuRec.RefClass.Department;
                            row["班級"] = stuRec.RefClass == null ? "" : stuRec.RefClass.ClassName;
                            row["班導師"] = (stuRec.RefClass == null || stuRec.RefClass.RefTeacher == null) ? "" : stuRec.RefClass.RefTeacher.TeacherName;
                            row["座號"] = stuRec.SeatNo;
                            row["學號"] = stuRec.StudentNumber;
                            row["姓名"] = stuRec.StudentName;
                            row["系統編號"] = "系統編號{" + stuRec.StudentID + "}";
                            row["定期評量"] = conf.ExamRecord.Name;

                            if (conf.ScoreCurDate != null)
                            {
                                row["成績校正年"] = conf.ScoreCurDate.Year;
                                row["成績校正月"] = conf.ScoreCurDate.Month;
                                row["成績校正日"] = conf.ScoreCurDate.Day;
                                row["成績校正民國年"] = (conf.ScoreCurDate.Year - 1911);
                                row["成績校正日期"] = conf.ScoreCurDate.ToShortDateString();
                            }

                            #endregion
                            #region 成績資料
                            #region 各科成績資料
                            #region 整理科目順序
                            List<string> subjectNameList = new List<string>();
                            if (studentExamSores.ContainsKey(stuRec.StudentID))
                            {
                                foreach (var subjectName in studentExamSores[studentID].Keys)
                                {
                                    foreach (var courseID in studentExamSores[studentID][subjectName].Keys)
                                    {
                                        if (conf.PrintSubjectList.Contains(subjectName))
                                        {
                                            subjectNameList.Add(subjectName);
                                        }
                                    }
                                }
                            }
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
                                if (subjectIndex <= conf.SubjectLimit)
                                {
                                    decimal? subjectNumber = null;
                                    // 檢查畫面上定期評量列印科目
                                    if (conf.PrintSubjectList.Contains(subjectName))
                                    {
                                        if (studentExamSores.ContainsKey(studentID))
                                        {
                                            if (studentExamSores[studentID].ContainsKey(subjectName))
                                            {
                                                foreach (var courseID in studentExamSores[studentID][subjectName].Keys)
                                                {
                                                    #region 評量成績
                                                    var sceTakeRecord = studentExamSores[studentID][subjectName][courseID];
                                                    if (sceTakeRecord != null)
                                                    {//有輸入
                                                        decimal level;
                                                        subjectNumber = decimal.TryParse(sceTakeRecord.SubjectLevel, out level) ? (decimal?)level : null;
                                                        row["科目名稱" + subjectIndex] = sceTakeRecord.Subject + GetNumber(subjectNumber);
                                                        row["學分數" + subjectIndex] = sceTakeRecord.CreditDec();
                                                        row["科目成績" + subjectIndex] = sceTakeRecord.SpecialCase == "" ? ("" + sceTakeRecord.ExamScore) : sceTakeRecord.SpecialCase;
                                                        #region 班排名及落點分析
                                                        string k1 = "";
                                                        if (RankMatrixDataDict.ContainsKey(studentID))
                                                        {
                                                            k1 = "定期評量/科目成績_" + sceTakeRecord.Subject + "_班排名";
                                                            if (RankMatrixDataDict[studentID].ContainsKey(k1))
                                                            {
                                                                if (RankMatrixDataDict[studentID][k1]["rank"] != null)
                                                                    row["班排名" + subjectIndex] = RankMatrixDataDict[studentID][k1]["rank"].ToString();

                                                                if (RankMatrixDataDict[studentID][k1]["matrix_count"] != null)
                                                                    row["班排名母數" + subjectIndex] = RankMatrixDataDict[studentID][k1]["matrix_count"].ToString();

                                                                // 五標PR填值
                                                                foreach (string rItem in r2List)
                                                                {
                                                                    if (RankMatrixDataDict[studentID][k1][rItem] != null)
                                                                        if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                                            row["班排名" + subjectIndex + "_" + rItem] = ParseScore(RankMatrixDataDict[studentID][k1][rItem].ToString());
                                                                        else
                                                                            row["班排名" + subjectIndex + "_" + rItem] = RankMatrixDataDict[studentID][k1][rItem].ToString();
                                                                }
                                                            }

                                                        }

                                                        #endregion
                                                        #region 科排名及落點分析
                                                        if (RankMatrixDataDict.ContainsKey(studentID))
                                                        {
                                                            k1 = "定期評量/科目成績_" + sceTakeRecord.Subject + "_科排名";
                                                            if (RankMatrixDataDict[studentID].ContainsKey(k1))
                                                            {
                                                                if (RankMatrixDataDict[studentID][k1]["rank"] != null)
                                                                    row["科排名" + subjectIndex] = RankMatrixDataDict[studentID][k1]["rank"].ToString();

                                                                if (RankMatrixDataDict[studentID][k1]["matrix_count"] != null)
                                                                    row["科排名母數" + subjectIndex] = RankMatrixDataDict[studentID][k1]["matrix_count"].ToString();

                                                                // 五標PR填值
                                                                foreach (string rItem in r2List)
                                                                {
                                                                    if (RankMatrixDataDict[studentID][k1][rItem] != null)
                                                                        if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                                            row["科排名" + subjectIndex + "_" + rItem] = ParseScore(RankMatrixDataDict[studentID][k1][rItem].ToString());
                                                                        else
                                                                            row["科排名" + subjectIndex + "_" + rItem] = RankMatrixDataDict[studentID][k1][rItem].ToString();
                                                                }
                                                            }
                                                        }

                                                        #endregion

                                                        #region 全校排名及落點分析
                                                        if (RankMatrixDataDict.ContainsKey(studentID))
                                                        {
                                                            k1 = "定期評量/科目成績_" + sceTakeRecord.Subject + "_年排名";
                                                            if (RankMatrixDataDict[studentID].ContainsKey(k1))
                                                            {
                                                                if (RankMatrixDataDict[studentID][k1]["rank"] != null)
                                                                    row["全校排名" + subjectIndex] = RankMatrixDataDict[studentID][k1]["rank"].ToString();

                                                                if (RankMatrixDataDict[studentID][k1]["matrix_count"] != null)
                                                                    row["全校排名母數" + subjectIndex] = RankMatrixDataDict[studentID][k1]["matrix_count"].ToString();

                                                                // 五標PR填值
                                                                foreach (string rItem in r2List)
                                                                {
                                                                    if (RankMatrixDataDict[studentID][k1][rItem] != null)
                                                                        if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                                            row["全校排名" + subjectIndex + "_" + rItem] = ParseScore(RankMatrixDataDict[studentID][k1][rItem].ToString());
                                                                        else
                                                                            row["全校排名" + subjectIndex + "_" + rItem] = RankMatrixDataDict[studentID][k1][rItem].ToString();
                                                                }
                                                            }
                                                        }

                                                        #endregion
                                                        #region 類別1排名及落點分析
                                                        if (RankMatrixDataDict.ContainsKey(studentID))
                                                        {
                                                            k1 = "定期評量/科目成績_" + sceTakeRecord.Subject + "_類別1排名";
                                                            if (RankMatrixDataDict[studentID].ContainsKey(k1))
                                                            {
                                                                if (RankMatrixDataDict[studentID][k1]["rank"] != null)
                                                                    row["類別1排名" + subjectIndex] = RankMatrixDataDict[studentID][k1]["rank"].ToString();

                                                                if (RankMatrixDataDict[studentID][k1]["matrix_count"] != null)
                                                                    row["類別1排名母數" + subjectIndex] = RankMatrixDataDict[studentID][k1]["matrix_count"].ToString();

                                                                // 五標PR填值
                                                                foreach (string rItem in r2List)
                                                                {
                                                                    if (RankMatrixDataDict[studentID][k1][rItem] != null)
                                                                        if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                                            row["類別1排名" + subjectIndex + "_" + rItem] = ParseScore(RankMatrixDataDict[studentID][k1][rItem].ToString());
                                                                        else
                                                                            row["類別1排名" + subjectIndex + "_" + rItem] = RankMatrixDataDict[studentID][k1][rItem].ToString();
                                                                }
                                                            }
                                                        }
                                                        #endregion
                                                        #region 類別2排名及落點分析
                                                        if (RankMatrixDataDict.ContainsKey(studentID))
                                                        {
                                                            k1 = "定期評量/科目成績_" + sceTakeRecord.Subject + "_類別2排名";
                                                            if (RankMatrixDataDict[studentID].ContainsKey(k1))
                                                            {
                                                                if (RankMatrixDataDict[studentID][k1]["rank"] != null)
                                                                    row["類別2排名" + subjectIndex] = RankMatrixDataDict[studentID][k1]["rank"].ToString();

                                                                if (RankMatrixDataDict[studentID][k1]["matrix_count"] != null)
                                                                    row["類別2排名母數" + subjectIndex] = RankMatrixDataDict[studentID][k1]["matrix_count"].ToString();

                                                                // 五標PR填值
                                                                foreach (string rItem in r2List)
                                                                {
                                                                    if (RankMatrixDataDict[studentID][k1][rItem] != null)
                                                                        if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                                            row["類別2排名" + subjectIndex + "_" + rItem] = ParseScore(RankMatrixDataDict[studentID][k1][rItem].ToString());
                                                                        else
                                                                            row["類別2排名" + subjectIndex + "_" + rItem] = RankMatrixDataDict[studentID][k1][rItem].ToString();
                                                                }
                                                            }
                                                        }
                                                        #endregion
                                                    }
                                                    else
                                                    {//修課有該考試但沒有成績資料
                                                        var courseRecs = accessHelper.CourseHelper.GetCourse(courseID);
                                                        if (courseRecs.Count > 0)
                                                        {
                                                            var courseRec = courseRecs[0];
                                                            decimal level;
                                                            subjectNumber = decimal.TryParse(courseRec.SubjectLevel, out level) ? (decimal?)level : null;
                                                            row["科目名稱" + subjectIndex] = courseRec.Subject + GetNumber(subjectNumber);
                                                            row["學分數" + subjectIndex] = courseRec.CreditDec();
                                                            row["科目成績" + subjectIndex] = "未輸入";
                                                        }
                                                    }
                                                    #endregion
                                                    #region 參考成績
                                                    if (studentRefExamSores.ContainsKey(studentID) && studentRefExamSores[studentID].ContainsKey(courseID))
                                                    {
                                                        row["前次成績" + subjectIndex] =
                                                                studentRefExamSores[studentID][courseID].SpecialCase == ""
                                                                ? ("" + studentRefExamSores[studentID][courseID].ExamScore)
                                                                : studentRefExamSores[studentID][courseID].SpecialCase;
                                                    }
                                                    #endregion
                                                    studentExamSores[studentID][subjectName].Remove(courseID);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    subjectIndex++;
                                }
                                else
                                {
                                    //重要!!發現資料在樣板中印不下時一定要記錄起來，否則使用者自己不會去發現的
                                    if (!overflowRecords.Contains(stuRec))
                                        overflowRecords.Add(stuRec);
                                }
                            }
                            #endregion


                            #region 總分
                            if (studentPrintSubjectSum.ContainsKey(studentID))
                            {
                                row["總分"] = studentPrintSubjectSum[studentID];

                                if (RankMatrixDataDict.ContainsKey(studentID))
                                {
                                    string skey = "定期評量/總計成績_總分_班排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["總分班排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["總分班排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["總分班排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["總分班排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }
                                    }

                                    skey = "定期評量/總計成績_總分_科排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["總分科排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["總分科排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["總分科排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["總分科排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }
                                    }

                                    skey = "定期評量/總計成績_總分_年排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["總分全校排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["總分全校排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["總分全校排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["總分全校排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }
                                    }
                                }
                            }
                            #endregion
                            #region 平均
                            if (studentPrintSubjectAvg.ContainsKey(studentID))
                            {
                                row["平均"] = studentPrintSubjectAvg[studentID];
                                if (RankMatrixDataDict.ContainsKey(studentID))
                                {
                                    string skey = "定期評量/總計成績_平均_班排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["平均班排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["平均班排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["平均班排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["平均班排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }

                                    }

                                    skey = "定期評量/總計成績_平均_科排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["平均科排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["平均科排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["平均科排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["平均科排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }
                                    }

                                    skey = "定期評量/總計成績_平均_年排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["平均全校排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["平均全校排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["平均全校排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["平均全校排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }

                                    }
                                }

                            }
                            #endregion
                            #region 加權總分
                            if (studentPrintSubjectSumW.ContainsKey(studentID))
                            {
                                row["加權總分"] = studentPrintSubjectSumW[studentID];
                                if (RankMatrixDataDict.ContainsKey(studentID))
                                {
                                    string skey = "定期評量/總計成績_加權總分_班排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權總分班排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權總分班排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["加權總分班排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["加權總分班排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }
                                    }

                                    skey = "定期評量/總計成績_加權總分_科排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權總分科排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權總分科排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["加權總分科排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["加權總分科排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }
                                    }

                                    skey = "定期評量/總計成績_加權總分_年排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權總分全校排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權總分全校排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["加權總分全校排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["加權總分全校排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }

                                    }
                                }
                            }
                            #endregion
                            #region 加權平均
                            if (studentPrintSubjectAvgW.ContainsKey(studentID))
                            {
                                row["加權平均"] = studentPrintSubjectAvgW[studentID];
                                if (RankMatrixDataDict.ContainsKey(studentID))
                                {

                                    string skey = "定期評量/總計成績_加權平均_班排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權平均班排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權平均班排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["加權平均班排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["加權平均班排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }
                                    }

                                    skey = "定期評量/總計成績_加權平均_科排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權平均科排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權平均科排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["加權平均科排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["加權平均科排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }

                                    }

                                    skey = "定期評量/總計成績_加權平均_年排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權平均全校排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權平均全校排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();
                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["加權平均全校排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["加權平均全校排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }
                                    }

                                }
                            }
                            #endregion
                            #region 類別1綜合成績


                            if (RankMatrixDataDict.ContainsKey(studentID))
                            {
                                string skey = "定期評量/總計成績_總分_類別1排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別1總分排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別1總分排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["rank_name"] != null)
                                        row["類別1名稱"] = RankMatrixDataDict[studentID][skey]["rank_name"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                row["類別1總分排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                            else
                                                row["類別1總分排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                    }

                                }

                                skey = "定期評量/總計成績_平均_類別1排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別1平均排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別1平均排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                row["類別1平均排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                            else
                                                row["類別1平均排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                    }
                                }

                                skey = "定期評量/總計成績_加權總分_類別1排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別1加權總分排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別1加權總分排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                row["類別1加權總分排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                            else
                                                row["類別1加權總分排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                    }
                                }

                                skey = "定期評量/總計成績_加權平均_類別1排名";
                                if (RankMatrixDataDict.ContainsKey(studentID))
                                {
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["類別1加權平均排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["類別1加權平均排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                                if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                    row["類別1加權平均排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                                else
                                                    row["類別1加權平均排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                        }

                                    }
                                }
                            }

                            #endregion
                            #region 類別2綜合成績

                            if (RankMatrixDataDict.ContainsKey(studentID))
                            {
                                string skey = "定期評量/總計成績_總分_類別2排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別2總分排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別2總分排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();
                                    if (RankMatrixDataDict[studentID][skey]["rank_name"] != null)
                                        row["類別2名稱"] = RankMatrixDataDict[studentID][skey]["rank_name"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                row["類別2總分排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                            else
                                                row["類別2總分排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                    }
                                }

                                skey = "定期評量/總計成績_平均_類別2排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別2平均排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別2平均排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                row["類別2平均排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                            else
                                                row["類別2平均排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                    }
                                }

                                skey = "定期評量/總計成績_加權總分_類別2排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別2加權總分排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別2加權總分排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                row["類別2加權總分排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                            else
                                                row["類別2加權總分排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                    }
                                }

                                skey = "定期評量/總計成績_加權平均_類別2排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別2加權平均排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別2加權平均排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            if (rItem.Contains("avg") || rItem.Contains("pr_") || rItem.Contains("std_dev_pop"))
                                                row["類別2加權平均排名_" + rItem] = ParseScore(RankMatrixDataDict[studentID][skey][rItem].ToString());
                                            else
                                                row["類別2加權平均排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                    }
                                }
                            }

                            #endregion


                            #region 學務資料
                            #region 獎懲統計
                            int 大功 = 0;
                            int 小功 = 0;
                            int 嘉獎 = 0;
                            int 大過 = 0;
                            int 小過 = 0;
                            int 警告 = 0;
                            bool 留校察看 = false;
                            foreach (RewardInfo info in stuRec.RewardList)
                            {
                                if (("" + info.Semester) == conf.Semester && ("" + info.SchoolYear) == conf.SchoolYear)
                                {
                                    大功 += info.AwardA;
                                    小功 += info.AwardB;
                                    嘉獎 += info.AwardC;
                                    if (!info.Cleared)
                                    {
                                        大過 += info.FaultA;
                                        小過 += info.FaultB;
                                        警告 += info.FaultC;
                                    }
                                    if (info.UltimateAdmonition)
                                        留校察看 = true;
                                }
                            }
                            row["大功統計"] = 大功 == 0 ? "" : ("" + 大功);
                            row["小功統計"] = 小功 == 0 ? "" : ("" + 小功);
                            row["嘉獎統計"] = 嘉獎 == 0 ? "" : ("" + 嘉獎);
                            row["大過統計"] = 大過 == 0 ? "" : ("" + 大過);
                            row["小過統計"] = 小過 == 0 ? "" : ("" + 小過);
                            row["警告統計"] = 警告 == 0 ? "" : ("" + 警告);
                            row["留校察看"] = 留校察看 ? "是" : "";
                            #endregion
                            #region 缺曠統計
                            Dictionary<string, int> 缺曠項目統計 = new Dictionary<string, int>();
                            foreach (AttendanceInfo info in stuRec.AttendanceList)
                            {
                                if (("" + info.Semester) == conf.Semester && ("" + info.SchoolYear) == conf.SchoolYear)
                                {
                                    string infoType = "";
                                    if (dicPeriodMappingType.ContainsKey(info.Period))
                                        infoType = dicPeriodMappingType[info.Period];
                                    else
                                        infoType = "";
                                    string attendanceKey = "" + infoType + "_" + info.Absence;
                                    if (!缺曠項目統計.ContainsKey(attendanceKey))
                                        缺曠項目統計.Add(attendanceKey, 0);
                                    缺曠項目統計[attendanceKey]++;
                                }
                            }
                            foreach (string attendanceKey in 缺曠項目統計.Keys)
                            {
                                row[attendanceKey] = 缺曠項目統計[attendanceKey] == 0 ? "" : ("" + 缺曠項目統計[attendanceKey]);
                            }
                            #endregion
                            #endregion
                            table.Rows.Add(row);
                            progressCount++;
                            bkw.ReportProgress(70 + progressCount * 20 / selectedStudents.Count);
                        }
                        #endregion

                        #endregion
                        bkw.ReportProgress(90);

                        // 收集學生清單資料並產生學生清單
                        _wbStudentList = new Aspose.Cells.Workbook();
                        _wbStudentList.Open(new MemoryStream(Properties.Resources.個人學期成績單_學生清單_));

                        int rowIdx = 1;
                        foreach (DataRow dr in table.Rows)
                        {
                            _wbStudentList.Worksheets[0].Cells[rowIdx, 0].PutValue(dr["班級"].ToString());
                            _wbStudentList.Worksheets[0].Cells[rowIdx, 1].PutValue(dr["座號"].ToString());
                            _wbStudentList.Worksheets[0].Cells[rowIdx, 2].PutValue(dr["學號"].ToString());
                            _wbStudentList.Worksheets[0].Cells[rowIdx, 3].PutValue(dr["姓名"].ToString());
                            _wbStudentList.Worksheets[0].Cells[rowIdx, 4].PutValue(dr["收件人"].ToString());
                            _wbStudentList.Worksheets[0].Cells[rowIdx, 5].PutValue(dr["收件人地址"].ToString());
                            rowIdx++;
                        }

                        document = conf.Template;
                        document.MailMerge.Execute(table);

                        //// debug
                        //table.TableName = "dt";
                        //table.WriteXml(Application.StartupPath + "\\dTxml.xml");
                    }
                    catch (Exception exception)
                    {
                        exc = exception;
                    }
                };
                bkw.RunWorkerAsync();
            }
        }

    }
}
