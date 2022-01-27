using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.Customization.Data;
using System.IO;

namespace SHStaticRank.Data
{
    public class CalcSemeSubjectRank
    {
        public static void Setup()
        {
            var button = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["計算固定排名"]["計算學期成績固定排名"];
            // 2018/3/31 穎驊整理模組註解，本功能的權限為 教務作業/功能按鈕/計算成績，先暫時維持原樣
            button.Enable = FISCA.Permission.UserAcl.Current["Button0670"].Executable;//MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["學期成績處理"].Enable = CurrentUser.Acl["Button0670"].Executable;
            button.Click += delegate { DoRank(); }; // 2018/5/31 穎驊為了要把 學期成績處理的 計算排名功能直接接到這邊， 將原本delegate 的內容，移到DoRanK 中，以利外部呼叫。
        }

        // 2018/5/31 穎驊註解，此為多載的寫法，決定使用者是否可以調整學年度學期，如果是教務作業 的批次作業而來 則不給動
        public static void DoRank()
        {
            DoRank(false);
        }

        public static void DoRank(bool fixSchoolyearSemester)
        {
            var conf = new CalcSemeSubjectStatucRank(fixSchoolyearSemester);
            if (conf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
                wb.Worksheets.Clear();
                AccessHelper accessHelper = new AccessHelper();
                UpdateHelper updateHelper = new UpdateHelper();
                var setting = conf.Setting;
                Dictionary<string, List<StudentRecord>> gradeyearStudents = new Dictionary<string, List<StudentRecord>>();

                System.ComponentModel.BackgroundWorker bkw = new System.ComponentModel.BackgroundWorker();
                bkw.WorkerReportsProgress = true;
                bkw.ProgressChanged += delegate (object sender, System.ComponentModel.ProgressChangedEventArgs e)
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("學期科目成績固定排名計算中...", e.ProgressPercentage);
                };
                Exception exc = null;
                bkw.RunWorkerCompleted += delegate
                {
                    #region 儲存檔案
                    string inputReportName = "學期科目成績固定排名";
                    string reportName = inputReportName;

                    string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                    if (!Directory.Exists(path))
                        Directory.CreateDirectory(path);
                    path = Path.Combine(path, reportName + ".xls");

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
                        wb.Save(path, Aspose.Cells.FileFormatType.Excel2003);
                        System.Diagnostics.Process.Start(path);
                    }
                    catch
                    {
                        System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                        sd.Title = "另存新檔";
                        sd.FileName = reportName + ".xls";
                        sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                        if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        {
                            try
                            {
                                wb.Save(path, Aspose.Cells.FileFormatType.Excel2003);

                            }
                            catch
                            {
                                FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                return;
                            }
                        }
                    }
                    #endregion
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("學期科目成績固定排名計算完成。", 100);
                    if (exc != null)
                    {
                        //2022-01-26 Cynthia 因學校不小心將一個學生加了兩個同個分組的不同類別標籤，導致產出來會直接炸掉，改為跳出錯誤訊息。
                        //throw new Exception("產生期末成績單發生錯誤", exc);
                        if (exc.Message.Contains("相同索引鍵"))
                            FISCA.Presentation.Controls.MsgBox.Show("計算排名發生錯誤：" + exc.Message+"\n\r若有設定類別排名，請修正學生類別標籤重複。","錯誤");
                        else
                            FISCA.Presentation.Controls.MsgBox.Show("計算排名發生錯誤：" + exc.Message, "錯誤");
                    }
                };
                bkw.DoWork += delegate
                {
                    try
                    {
                        int yearCount = 0;
                        bkw.ReportProgress(1);
                        #region 整各年級學生
                        foreach (var studentRec in accessHelper.StudentHelper.GetAllStudent())
                        {
                            if (studentRec.Status != "一般")
                                continue;
                            if (studentRec.RefClass != null)
                            {
                                string grade = "";
                                grade = "" + studentRec.RefClass.GradeYear;
                                if ((setting.CalcGradeYear1 && grade == "1") || (setting.CalcGradeYear2 && grade == "2") || (setting.CalcGradeYear3 && grade == "3"))
                                {
                                    if (!gradeyearStudents.ContainsKey(grade))
                                        gradeyearStudents.Add(grade, new List<StudentRecord>());
                                    gradeyearStudents[grade].Add(studentRec);
                                }
                            }
                        }
                        #endregion
                        foreach (var gradeyear in gradeyearStudents.Keys)
                        {
                            yearCount++;
                            var studentSheet = wb.Worksheets[wb.Worksheets.Add()];
                            Aspose.Cells.Worksheet studentSheet2 = null;
                            studentSheet.Name = "" + gradeyear + "年級";
                            int index = 0, index2 = 0;
                            studentSheet.Cells[0, index++].PutValue("學生系統編號");
                            studentSheet.Cells[0, index++].PutValue("學號");
                            studentSheet.Cells[0, index++].PutValue("班級");
                            studentSheet.Cells[0, index++].PutValue("座號");
                            studentSheet.Cells[0, index++].PutValue("科別");
                            studentSheet.Cells[0, index++].PutValue("姓名");
                            studentSheet.Cells[0, index++].PutValue("學年度");
                            studentSheet.Cells[0, index++].PutValue("學期");
                            studentSheet.Cells[0, index++].PutValue("科目");
                            studentSheet.Cells[0, index++].PutValue("科目級別");
                            studentSheet.Cells[0, index++].PutValue("原始成績");
                            studentSheet.Cells[0, index++].PutValue("補考成績");
                            studentSheet.Cells[0, index++].PutValue("重修成績");
                            studentSheet.Cells[0, index++].PutValue("手動調整成績");
                            studentSheet.Cells[0, index++].PutValue("學年調整成績");
                            studentSheet.Cells[0, index++].PutValue("排名採計成績");
                            studentSheet.Cells[0, index++].PutValue("班排名");
                            studentSheet.Cells[0, index++].PutValue("科排名");
                            studentSheet.Cells[0, index++].PutValue("校排名");
                            if (setting.計算學業成績排名)
                            {
                                studentSheet2 = wb.Worksheets[wb.Worksheets.Add()];
                                studentSheet2.Name = "" + gradeyear + "年級學業";
                                studentSheet2.Cells[0, index2++].PutValue("學生系統編號");
                                studentSheet2.Cells[0, index2++].PutValue("學號");
                                studentSheet2.Cells[0, index2++].PutValue("班級");
                                studentSheet2.Cells[0, index2++].PutValue("座號");
                                studentSheet2.Cells[0, index2++].PutValue("科別");
                                studentSheet2.Cells[0, index2++].PutValue("姓名");
                                studentSheet2.Cells[0, index2++].PutValue("學年度");
                                studentSheet2.Cells[0, index2++].PutValue("學期");
                                studentSheet2.Cells[0, index2++].PutValue("學業成績");
                                studentSheet2.Cells[0, index2++].PutValue("班排名");
                                studentSheet2.Cells[0, index2++].PutValue("科排名");
                                studentSheet2.Cells[0, index2++].PutValue("校排名");
                            }
                            //每個類別排名的index
                            Dictionary<string, int> catIndex = new Dictionary<string, int>();
                            Dictionary<string, int> catIndex2 = new Dictionary<string, int>();
                            int rowIndex = 1;
                            int rowIndex2 = 1;
                            var studentList = gradeyearStudents[gradeyear];
                            #region 分析學生所屬類別
                            List<string> cat1List = new List<string>(), cat2List = new List<string>();
                            //不排名的學生移出list
                            List<StudentRecord> notRankList = new List<StudentRecord>();
                            foreach (var studentRec in studentList)
                            {
                                bool noRankStu = false;
                                foreach (var tag in studentRec.StudentCategorys)
                                {
                                    if (tag.SubCategory == "")
                                    {
                                        if (setting.NotRankTag != "" && setting.NotRankTag == tag.Name)
                                        {
                                            notRankList.Add(studentRec);
                                            noRankStu = true;
                                            break;
                                        }
                                    }
                                    else
                                    {
                                        if (setting.NotRankTag != "" && setting.NotRankTag == "[" + tag.Name + "]")
                                        {
                                            notRankList.Add(studentRec);
                                            noRankStu = true;
                                            break;
                                        }
                                    }
                                }
                                if (!noRankStu)
                                {
                                    foreach (var tag in studentRec.StudentCategorys)
                                    {
                                        if (tag.SubCategory == "")
                                        {
                                            if (setting.Rank1Tag != "" && setting.Rank1Tag == tag.Name)
                                            {
                                                studentRec.Fields.Add("tag1", tag.Name);
                                                if (!cat1List.Contains(tag.Name)) cat1List.Add(tag.Name);
                                            }
                                            if (setting.Rank2Tag != "" && setting.Rank2Tag == tag.Name)
                                            {
                                                studentRec.Fields.Add("tag2", tag.Name);
                                                if (!cat2List.Contains(tag.Name)) cat2List.Add(tag.Name);
                                            }
                                        }
                                        else
                                        {
                                            if (setting.Rank1Tag != "" && setting.Rank1Tag == "[" + tag.Name + "]")
                                            {
                                                studentRec.Fields.Add("tag1", tag.SubCategory);
                                                if (!cat1List.Contains(tag.SubCategory)) cat1List.Add(tag.SubCategory);
                                            }
                                            if (setting.Rank2Tag != "" && setting.Rank2Tag == "[" + tag.Name + "]")
                                            {
                                                studentRec.Fields.Add("tag2", tag.SubCategory);
                                                if (!cat2List.Contains(tag.SubCategory)) cat2List.Add(tag.SubCategory);
                                            }
                                        }
                                    }
                                }
                            }
                            //不排名的學生直接移出list
                            foreach (var r in notRankList)
                            {
                                studentList.Remove(r);
                            }
                            foreach (var item in cat1List)
                            {
                                string value = item + "排名(類1)";
                                catIndex.Add(value, index);
                                studentSheet.Cells[0, index++].PutValue(value);
                                if (setting.計算學業成績排名)
                                {
                                    catIndex2.Add(value, index2);
                                    studentSheet2.Cells[0, index2++].PutValue(value);
                                }
                            }
                            foreach (var item in cat2List)
                            {
                                string value = item + "排名(類2)";
                                catIndex.Add(value, index);
                                studentSheet.Cells[0, index++].PutValue(value);
                                if (setting.計算學業成績排名)
                                {
                                    catIndex2.Add(value, index2);
                                    studentSheet2.Cells[0, index2++].PutValue(value);
                                }
                            }
                            #endregion
                            //取得學生學期科目成績
                            accessHelper.StudentHelper.FillSemesterSubjectScore(false, studentList);
                            if (setting.計算學業成績排名)
                            {
                                accessHelper.StudentHelper.FillSemesterEntryScore(false, studentList);
                            }
                            bkw.ReportProgress(1 + 99 * yearCount / gradeyearStudents.Count / 2);
                            Dictionary<string, List<decimal>> ranks = new Dictionary<string, List<decimal>>();
                            Dictionary<string, List<string>> rankStudents = new Dictionary<string, List<string>>();
                            Dictionary<SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo, decimal?> selectScore = new Dictionary<SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo, decimal?>();
                            foreach (var studentRec in studentList)
                            {
                                string studentID = studentRec.StudentID;
                                foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                {
                                    if (("" + subjectScore.SchoolYear) == setting.SchoolYear && ("" + subjectScore.Semester) == setting.Semester)
                                    {
                                        decimal score = decimal.MinValue, tryParseScore;
                                        bool match = false;
                                        #region 取最高分
                                        if (setting.use手動調整成績 && decimal.TryParse(subjectScore.Detail.GetAttribute("擇優採計成績"), out tryParseScore))
                                        {
                                            match = true;
                                            if (score < tryParseScore)
                                                score = tryParseScore;
                                        }
                                        if (setting.use重修成績 && decimal.TryParse(subjectScore.Detail.GetAttribute("重修成績"), out tryParseScore))
                                        {
                                            match = true;
                                            if (score < tryParseScore)
                                                score = tryParseScore;
                                        }
                                        if (setting.use原始成績 && decimal.TryParse(subjectScore.Detail.GetAttribute("原始成績"), out tryParseScore))
                                        {
                                            match = true;
                                            if (score < tryParseScore)
                                                score = tryParseScore;
                                        }
                                        if (setting.use補考成績 && decimal.TryParse(subjectScore.Detail.GetAttribute("補考成績"), out tryParseScore))
                                        {
                                            match = true;
                                            if (score < tryParseScore)
                                                score = tryParseScore;
                                        }
                                        if (setting.use學年調整成績 && decimal.TryParse(subjectScore.Detail.GetAttribute("學年調整成績"), out tryParseScore))
                                        {
                                            match = true;
                                            if (score < tryParseScore)
                                                score = tryParseScore;
                                        }
                                        #endregion
                                        if (!match)
                                        {
                                            selectScore.Add(subjectScore, null);
                                        }
                                        else
                                        {
                                            selectScore.Add(subjectScore, score);
                                            #region 班排名
                                            {
                                                string key = "班排名" + studentRec.RefClass.ClassID + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(score);
                                                rankStudents[key].Add(studentID);
                                            }
                                            #endregion
                                            #region 科排名
                                            {
                                                if (studentRec.Department != "")
                                                {
                                                    //各科目科排名
                                                    string key = "科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                                    if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                    ranks[key].Add(score);
                                                    rankStudents[key].Add(studentID);
                                                }
                                            }
                                            #endregion
                                            #region 全校排名
                                            {
                                                string key = "全校排名" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(score);
                                                rankStudents[key].Add(studentID);
                                            }
                                            #endregion
                                            #region 類別1排名
                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                string key = "類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(score);
                                                rankStudents[key].Add(studentID);
                                            }
                                            #endregion
                                            #region 類別2排名
                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                string key = "類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(score);
                                                rankStudents[key].Add(studentID);
                                            }
                                            #endregion
                                        }
                                    }
                                }
                                if (setting.計算學業成績排名)
                                {
                                    foreach (var semesterEntry in studentRec.SemesterEntryScoreList)
                                    {
                                        if (semesterEntry.Entry == "學業"
                                            && ("" + semesterEntry.SchoolYear) == setting.SchoolYear
                                            && ("" + semesterEntry.Semester) == setting.Semester)
                                        {
                                            #region 班排名
                                            {
                                                string key = "學業成績班排名" + studentRec.RefClass.ClassID + "^^^";
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(semesterEntry.Score);
                                                rankStudents[key].Add(studentID);
                                            }
                                            #endregion
                                            #region 科排名
                                            {
                                                if (studentRec.Department != "")
                                                {
                                                    //各科目科排名
                                                    string key = "學業成績科排名" + studentRec.Department + "^^^" + gradeyear;
                                                    if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                    ranks[key].Add(semesterEntry.Score);
                                                    rankStudents[key].Add(studentID);
                                                }
                                            }
                                            #endregion
                                            #region 全校排名
                                            {
                                                string key = "學業成績全校排名" + gradeyear + "^^^";
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(semesterEntry.Score);
                                                rankStudents[key].Add(studentID);
                                            }
                                            #endregion
                                            #region 類別1排名
                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                string key = "學業成績類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(semesterEntry.Score);
                                                rankStudents[key].Add(studentID);
                                            }
                                            #endregion
                                            #region 類別2排名
                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                string key = "學業成績類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                ranks[key].Add(semesterEntry.Score);
                                                rankStudents[key].Add(studentID);
                                            }
                                            #endregion
                                        }
                                    }
                                }
                            }
                            foreach (var k in ranks.Keys)
                            {
                                var rankscores = ranks[k];
                                //排序
                                rankscores.Sort();
                                rankscores.Reverse();
                            }
                            foreach (var studentRec in studentList)
                            {
                                string studentID = studentRec.StudentID;
                                #region 列表
                                foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                {
                                    if (("" + subjectScore.SchoolYear) == setting.SchoolYear && ("" + subjectScore.Semester) == setting.Semester)
                                    {
                                        index = 0;
                                        studentSheet.Cells[rowIndex, index++].PutValue(studentRec.StudentID);
                                        studentSheet.Cells[rowIndex, index++].PutValue(studentRec.StudentNumber);
                                        studentSheet.Cells[rowIndex, index++].PutValue(studentRec.RefClass.ClassName);
                                        studentSheet.Cells[rowIndex, index++].PutValue(studentRec.SeatNo);
                                        studentSheet.Cells[rowIndex, index++].PutValue(studentRec.Department);
                                        studentSheet.Cells[rowIndex, index++].PutValue(studentRec.StudentName);
                                        studentSheet.Cells[rowIndex, index++].PutValue(subjectScore.SchoolYear);
                                        studentSheet.Cells[rowIndex, index++].PutValue(subjectScore.Semester);
                                        studentSheet.Cells[rowIndex, index++].PutValue(subjectScore.Subject);
                                        studentSheet.Cells[rowIndex, index++].PutValue(subjectScore.Level);
                                        studentSheet.Cells[rowIndex, index++].PutValue(subjectScore.Detail.GetAttribute("原始成績"));
                                        studentSheet.Cells[rowIndex, index++].PutValue(subjectScore.Detail.GetAttribute("補考成績"));
                                        studentSheet.Cells[rowIndex, index++].PutValue(subjectScore.Detail.GetAttribute("重修成績"));
                                        studentSheet.Cells[rowIndex, index++].PutValue(subjectScore.Detail.GetAttribute("擇優採計成績"));
                                        studentSheet.Cells[rowIndex, index++].PutValue(subjectScore.Detail.GetAttribute("學年調整成績"));
                                        studentSheet.Cells[rowIndex, index++].PutValue(selectScore[subjectScore]);
                                        #region studentSheet.Cells[rowIndex, index++].PutValue("班排名");
                                        string key = "班排名" + studentRec.RefClass.ClassID + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                        if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                            studentSheet.Cells[rowIndex, index++].PutValue(ranks[key].IndexOf(selectScore[subjectScore].Value) + 1);
                                        else
                                            index++;
                                        #endregion
                                        #region studentSheet.Cells[rowIndex, index++].PutValue("科排名");
                                        key = "科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                        if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                            studentSheet.Cells[rowIndex, index++].PutValue(ranks[key].IndexOf(selectScore[subjectScore].Value) + 1);
                                        else
                                            index++;
                                        #endregion
                                        #region studentSheet.Cells[rowIndex, index++].PutValue("校排名");
                                        key = "全校排名" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                        if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                            studentSheet.Cells[rowIndex, index++].PutValue(ranks[key].IndexOf(selectScore[subjectScore].Value) + 1);
                                        else
                                            index++;
                                        #endregion
                                        #region 類別排名1
                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            key = "類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                studentSheet.Cells[rowIndex, catIndex["" + studentRec.Fields["tag1"] + "排名(類1)"]].PutValue(ranks[key].IndexOf(selectScore[subjectScore].Value) + 1);
                                        }
                                        #endregion
                                        #region 類別排名2
                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            key = "類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                studentSheet.Cells[rowIndex, catIndex["" + studentRec.Fields["tag2"] + "排名(類2)"]].PutValue(ranks[key].IndexOf(selectScore[subjectScore].Value) + 1);
                                        }
                                        #endregion
                                        rowIndex++;
                                    }
                                }
                                #endregion
                                if (setting.計算學業成績排名)
                                {
                                    index2 = 0;
                                    foreach (var semesterEntryScore in studentRec.SemesterEntryScoreList)
                                    {
                                        if (semesterEntryScore.Entry == "學業"
                                            && ("" + semesterEntryScore.SchoolYear) == setting.SchoolYear
                                            && ("" + semesterEntryScore.Semester) == setting.Semester)
                                        {
                                            studentSheet2.Cells[rowIndex2, index2++].PutValue(studentRec.StudentID);
                                            studentSheet2.Cells[rowIndex2, index2++].PutValue(studentRec.StudentNumber);
                                            studentSheet2.Cells[rowIndex2, index2++].PutValue(studentRec.RefClass.ClassName);
                                            studentSheet2.Cells[rowIndex2, index2++].PutValue(studentRec.SeatNo);
                                            studentSheet2.Cells[rowIndex2, index2++].PutValue(studentRec.Department);
                                            studentSheet2.Cells[rowIndex2, index2++].PutValue(studentRec.StudentName);
                                            studentSheet2.Cells[rowIndex2, index2++].PutValue(semesterEntryScore.SchoolYear);
                                            studentSheet2.Cells[rowIndex2, index2++].PutValue(semesterEntryScore.Semester);
                                            studentSheet2.Cells[rowIndex2, index2++].PutValue(semesterEntryScore.Score);
                                            string key = "學業成績班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                studentSheet2.Cells[rowIndex2, index2++].PutValue(ranks[key].IndexOf(semesterEntryScore.Score) + 1);
                                            else
                                                index2++;
                                            key = "學業成績科排名" + studentRec.Department + "^^^" + gradeyear;
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                studentSheet2.Cells[rowIndex2, index2++].PutValue(ranks[key].IndexOf(semesterEntryScore.Score) + 1);
                                            else
                                                index2++;
                                            key = "學業成績全校排名" + gradeyear + "^^^";
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                studentSheet2.Cells[rowIndex2, index2++].PutValue(ranks[key].IndexOf(semesterEntryScore.Score) + 1);
                                            else
                                                index2++;
                                            #region 類別1排名
                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                key = "學業成績類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                    studentSheet2.Cells[rowIndex2, catIndex2["" + studentRec.Fields["tag1"] + "排名(類1)"]].PutValue(ranks[key].IndexOf(semesterEntryScore.Score) + 1);
                                            }
                                            #endregion
                                            #region 類別2排名
                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                key = "學業成績類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                    studentSheet2.Cells[rowIndex2, catIndex2["" + studentRec.Fields["tag2"] + "排名(類2)"]].PutValue(ranks[key].IndexOf(semesterEntryScore.Score) + 1);
                                            }
                                            #endregion
                                            rowIndex2++;
                                        }
                                    }
                                }
                            }
                            studentSheet.AutoFitColumns();
                            if (setting.計算學業成績排名)
                                studentSheet2.AutoFitColumns();
                            #region 儲存
                            if (!setting.DoNotSaveIt)
                            {
                                System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                                List<string> sql = new List<string>();
                                int count = 0;
                                #region 科目成績排名
                                #region 有排名學生
                                foreach (var studentRec in studentList)
                                {
                                    count++;
                                    bkw.ReportProgress(1 + (99 * yearCount / gradeyearStudents.Count / 2) + (99 * yearCount * count / studentList.Count / gradeyearStudents.Count / 2));

                                    string studentID = studentRec.StudentID;
                                    bool hasclassRating = false;
                                    bool hasdeptRating = false;
                                    bool hasschoolRating = false;
                                    bool hastag1Rating = false;
                                    bool hastag2Rating = false;
                                    var classRating = doc.CreateElement("Rating"); classRating.SetAttribute("範圍人數", "0");
                                    var deptRating = doc.CreateElement("Rating"); deptRating.SetAttribute("範圍人數", "0");
                                    var schoolRating = doc.CreateElement("Rating"); schoolRating.SetAttribute("範圍人數", "0");
                                    var tag1Rating = doc.CreateElement("Rating"); tag1Rating.SetAttribute("範圍人數", "0");
                                    var tag2Rating = doc.CreateElement("Rating"); tag2Rating.SetAttribute("範圍人數", "0");

                                    foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                    {
                                        if (("" + subjectScore.SchoolYear) == setting.SchoolYear && ("" + subjectScore.Semester) == setting.Semester)
                                        {
                                            #region 班排名
                                            string key = "班排名" + studentRec.RefClass.ClassID + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                            {
                                                hasclassRating = true;
                                                var item = doc.CreateElement("Item");
                                                item.SetAttribute("成績", "" + selectScore[subjectScore].Value);
                                                item.SetAttribute("成績人數", "" + ranks[key].Count);
                                                item.SetAttribute("排名", "" + (ranks[key].IndexOf(selectScore[subjectScore].Value) + 1));
                                                item.SetAttribute("科目", "" + subjectScore.Subject);
                                                item.SetAttribute("科目級別", "" + subjectScore.Level);
                                                classRating.AppendChild(item);
                                            }
                                            #endregion
                                            #region 科排名
                                            key = "科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                            {
                                                hasdeptRating = true;
                                                var item = doc.CreateElement("Item");
                                                item.SetAttribute("成績", "" + selectScore[subjectScore].Value);
                                                item.SetAttribute("成績人數", "" + ranks[key].Count);
                                                item.SetAttribute("排名", "" + (ranks[key].IndexOf(selectScore[subjectScore].Value) + 1));
                                                item.SetAttribute("科目", "" + subjectScore.Subject);
                                                item.SetAttribute("科目級別", "" + subjectScore.Level);
                                                deptRating.AppendChild(item);
                                            }
                                            #endregion
                                            #region 全校排名
                                            key = "全校排名" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                            {
                                                hasschoolRating = true;
                                                var item = doc.CreateElement("Item");
                                                item.SetAttribute("成績", "" + selectScore[subjectScore].Value);
                                                item.SetAttribute("成績人數", "" + ranks[key].Count);
                                                item.SetAttribute("排名", "" + (ranks[key].IndexOf(selectScore[subjectScore].Value) + 1));
                                                item.SetAttribute("科目", "" + subjectScore.Subject);
                                                item.SetAttribute("科目級別", "" + subjectScore.Level);
                                                schoolRating.AppendChild(item);
                                            }
                                            #endregion
                                            #region 類別排名1
                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                tag1Rating.SetAttribute("類別", "" + studentRec.Fields["tag1"]);
                                                key = "類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;

                                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                {
                                                    hastag1Rating = true;
                                                    var item = doc.CreateElement("Item");
                                                    item.SetAttribute("成績", "" + selectScore[subjectScore].Value);
                                                    item.SetAttribute("成績人數", "" + ranks[key].Count);
                                                    item.SetAttribute("排名", "" + (ranks[key].IndexOf(selectScore[subjectScore].Value) + 1));
                                                    item.SetAttribute("科目", "" + subjectScore.Subject);
                                                    item.SetAttribute("科目級別", "" + subjectScore.Level);
                                                    tag1Rating.AppendChild(item);
                                                }
                                            }
                                            #endregion
                                            #region 類別排名2
                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                tag2Rating.SetAttribute("類別", "" + studentRec.Fields["tag2"]);
                                                key = "類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjectScore.Subject + "^^^" + subjectScore.Level;
                                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                {
                                                    hastag2Rating = true;
                                                    var item = doc.CreateElement("Item");
                                                    item.SetAttribute("成績", "" + selectScore[subjectScore].Value);
                                                    item.SetAttribute("成績人數", "" + ranks[key].Count);
                                                    item.SetAttribute("排名", "" + (ranks[key].IndexOf(selectScore[subjectScore].Value) + 1));
                                                    item.SetAttribute("科目", "" + subjectScore.Subject);
                                                    item.SetAttribute("科目級別", "" + subjectScore.Level);
                                                    tag2Rating.AppendChild(item);
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                    sql.Add(string.Format(
@"update sems_subj_score set 
    class_rating='{0}'
    , dept_rating='{1}'
    , year_rating='{2}'
    , group_rating='{3}' 
WHERE ref_student_id ={4} and school_year={5} and semester={6};
"
                                        , hasclassRating ? classRating.OuterXml : ""
                                        , hasdeptRating ? deptRating.OuterXml : ""
                                        , hasschoolRating ? schoolRating.OuterXml : ""
                                        , (hastag1Rating && hastag2Rating) ?
                                            ("<Ratings>" + tag1Rating.OuterXml + tag2Rating.OuterXml + "</Ratings>") :
                                            (hastag1Rating ?
                                                ("<Ratings>" + tag1Rating.OuterXml + "</Ratings>") :
                                                (hastag2Rating ?
                                                    ("<Ratings>" + tag2Rating.OuterXml + "</Ratings>") :
                                                    ""))
                                        , studentID
                                        , setting.SchoolYear
                                        , setting.Semester
                                    ));
                                    if (sql.Count == 200)
                                    {
                                        updateHelper.Execute(sql);
                                        sql.Clear();
                                    }
                                }
                                #endregion
                                if (setting.清除不排名學生排名資料)
                                {
                                    #region 不排名學生
                                    foreach (var studentRec in notRankList)
                                    {
                                        count++;
                                        bkw.ReportProgress(1 + (99 * yearCount / gradeyearStudents.Count / 2) + (99 * yearCount * count / studentList.Count / gradeyearStudents.Count / 2));

                                        string studentID = studentRec.StudentID;
                                        bool hasclassRating = false;
                                        bool hasdeptRating = false;
                                        bool hasschoolRating = false;
                                        bool hastag1Rating = false;
                                        bool hastag2Rating = false;
                                        var classRating = doc.CreateElement("Rating"); classRating.SetAttribute("範圍人數", "0");
                                        var deptRating = doc.CreateElement("Rating"); deptRating.SetAttribute("範圍人數", "0");
                                        var schoolRating = doc.CreateElement("Rating"); schoolRating.SetAttribute("範圍人數", "0");
                                        var tag1Rating = doc.CreateElement("Rating"); tag1Rating.SetAttribute("範圍人數", "0");
                                        var tag2Rating = doc.CreateElement("Rating"); tag2Rating.SetAttribute("範圍人數", "0");
                                        sql.Add(string.Format(
@"update sems_subj_score set 
    class_rating='{0}'
    , dept_rating='{1}'
    , year_rating='{2}'
    , group_rating='{3}' 
WHERE ref_student_id ={4} and school_year={5} and semester={6};
"
                                            , hasclassRating ? classRating.OuterXml : ""
                                            , hasdeptRating ? deptRating.OuterXml : ""
                                            , hasschoolRating ? schoolRating.OuterXml : ""
                                            , (hastag1Rating && hastag2Rating) ?
                                                ("<Ratings>" + tag1Rating.OuterXml + tag2Rating.OuterXml + "</Ratings>") :
                                                (hastag1Rating ?
                                                    ("<Ratings>" + tag1Rating.OuterXml + "</Ratings>") :
                                                    (hastag2Rating ?
                                                        ("<Ratings>" + tag2Rating.OuterXml + "</Ratings>") :
                                                        ""))
                                            , studentID
                                            , setting.SchoolYear
                                            , setting.Semester
                                        ));
                                        if (sql.Count == 200)
                                        {
                                            updateHelper.Execute(sql);
                                            sql.Clear();
                                        }
                                    }
                                    #endregion
                                }
                                if (sql.Count > 0)
                                {
                                    updateHelper.Execute(sql);
                                    sql.Clear();
                                }
                                #endregion
                                #region 學業成績排名
                                #region 有排名學生
                                foreach (var studentRec in studentList)
                                {
                                    string studentID = studentRec.StudentID;
                                    bool hasclassRating = false;
                                    bool hasdeptRating = false;
                                    bool hasschoolRating = false;
                                    bool hastag1Rating = false;
                                    bool hastag2Rating = false;
                                    var classRating = doc.CreateElement("Rating"); classRating.SetAttribute("範圍人數", "0");
                                    var deptRating = doc.CreateElement("Rating"); deptRating.SetAttribute("範圍人數", "0");
                                    var schoolRating = doc.CreateElement("Rating"); schoolRating.SetAttribute("範圍人數", "0");
                                    var tag1Rating = doc.CreateElement("Rating"); tag1Rating.SetAttribute("範圍人數", "0");
                                    var tag2Rating = doc.CreateElement("Rating"); tag2Rating.SetAttribute("範圍人數", "0");

                                    foreach (var entryScore in studentRec.SemesterEntryScoreList)
                                    {

                                        if (entryScore.Entry == "學業"
                                            && ("" + entryScore.SchoolYear) == setting.SchoolYear
                                            && ("" + entryScore.Semester) == setting.Semester)
                                        {
                                            #region 班排名
                                            string key = "學業成績班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                            {
                                                hasclassRating = true;
                                                var item = doc.CreateElement("Item");
                                                item.SetAttribute("成績", "" + entryScore.Score);
                                                item.SetAttribute("成績人數", "" + ranks[key].Count);
                                                item.SetAttribute("排名", "" + (ranks[key].IndexOf(entryScore.Score) + 1));
                                                item.SetAttribute("分項", "學業");
                                                classRating.AppendChild(item);
                                            }
                                            #endregion
                                            #region 科排名
                                            key = "學業成績科排名" + studentRec.Department + "^^^" + gradeyear;
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                            {
                                                hasdeptRating = true;
                                                var item = doc.CreateElement("Item");
                                                item.SetAttribute("成績", "" + entryScore.Score);
                                                item.SetAttribute("成績人數", "" + ranks[key].Count);
                                                item.SetAttribute("排名", "" + (ranks[key].IndexOf(entryScore.Score) + 1));
                                                item.SetAttribute("分項", "學業");
                                                deptRating.AppendChild(item);
                                            }
                                            #endregion
                                            #region 全校排名
                                            key = "學業成績全校排名" + gradeyear + "^^^";
                                            if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                            {
                                                hasschoolRating = true;
                                                var item = doc.CreateElement("Item");
                                                item.SetAttribute("成績", "" + entryScore.Score);
                                                item.SetAttribute("成績人數", "" + ranks[key].Count);
                                                item.SetAttribute("排名", "" + (ranks[key].IndexOf(entryScore.Score) + 1));
                                                item.SetAttribute("分項", "學業");
                                                schoolRating.AppendChild(item);
                                            }
                                            #endregion
                                            #region 類別排名1
                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                tag1Rating.SetAttribute("類別", "" + studentRec.Fields["tag1"]);
                                                key = "學業成績類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;

                                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                {
                                                    hastag1Rating = true;
                                                    var item = doc.CreateElement("Item");
                                                    item.SetAttribute("成績", "" + entryScore.Score);
                                                    item.SetAttribute("成績人數", "" + ranks[key].Count);
                                                    item.SetAttribute("排名", "" + (ranks[key].IndexOf(entryScore.Score) + 1));
                                                    item.SetAttribute("分項", "學業");
                                                    tag1Rating.AppendChild(item);
                                                }
                                            }
                                            #endregion
                                            #region 類別排名2
                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                tag2Rating.SetAttribute("類別", "" + studentRec.Fields["tag2"]);
                                                key = key = "學業成績類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                {
                                                    hastag2Rating = true;
                                                    var item = doc.CreateElement("Item");
                                                    item.SetAttribute("成績", "" + entryScore.Score);
                                                    item.SetAttribute("成績人數", "" + ranks[key].Count);
                                                    item.SetAttribute("排名", "" + (ranks[key].IndexOf(entryScore.Score) + 1));
                                                    item.SetAttribute("分項", "學業");
                                                    tag2Rating.AppendChild(item);
                                                }
                                            }
                                            #endregion
                                        }
                                    }
                                    sql.Add(string.Format(
@"update sems_entry_score set 
    class_rating='{0}'
    , dept_rating='{1}'
    , year_rating='{2}'
    , group_rating='{3}' 
WHERE ref_student_id ={4} and school_year={5} and semester={6} and entry_group=1;
"
                                        , hasclassRating ? classRating.OuterXml : ""
                                        , hasdeptRating ? deptRating.OuterXml : ""
                                        , hasschoolRating ? schoolRating.OuterXml : ""
                                        , (hastag1Rating && hastag2Rating) ?
                                            ("<Ratings>" + tag1Rating.OuterXml + tag2Rating.OuterXml + "</Ratings>") :
                                            (hastag1Rating ?
                                                ("<Ratings>" + tag1Rating.OuterXml + "</Ratings>") :
                                                (hastag2Rating ?
                                                    ("<Ratings>" + tag2Rating.OuterXml + "</Ratings>") :
                                                    ""))
                                        , studentID
                                        , setting.SchoolYear
                                        , setting.Semester
                                    ));
                                    if (sql.Count == 200)
                                    {
                                        updateHelper.Execute(sql);
                                        sql.Clear();
                                    }
                                }
                                #endregion
                                if (setting.清除不排名學生排名資料)
                                {
                                    #region 不排名學生
                                    foreach (var studentRec in notRankList)
                                    {
                                        string studentID = studentRec.StudentID;
                                        bool hasclassRating = false;
                                        bool hasdeptRating = false;
                                        bool hasschoolRating = false;
                                        bool hastag1Rating = false;
                                        bool hastag2Rating = false;
                                        var classRating = doc.CreateElement("Rating"); classRating.SetAttribute("範圍人數", "0");
                                        var deptRating = doc.CreateElement("Rating"); deptRating.SetAttribute("範圍人數", "0");
                                        var schoolRating = doc.CreateElement("Rating"); schoolRating.SetAttribute("範圍人數", "0");
                                        var tag1Rating = doc.CreateElement("Rating"); tag1Rating.SetAttribute("範圍人數", "0");
                                        var tag2Rating = doc.CreateElement("Rating"); tag2Rating.SetAttribute("範圍人數", "0");

                                        sql.Add(string.Format(
@"update sems_entry_score set 
    class_rating='{0}'
    , dept_rating='{1}'
    , year_rating='{2}'
    , group_rating='{3}' 
WHERE ref_student_id ={4} and school_year={5} and semester={6} and entry_group=1;
"
                                            , hasclassRating ? classRating.OuterXml : ""
                                            , hasdeptRating ? deptRating.OuterXml : ""
                                            , hasschoolRating ? schoolRating.OuterXml : ""
                                            , (hastag1Rating && hastag2Rating) ?
                                                ("<Ratings>" + tag1Rating.OuterXml + tag2Rating.OuterXml + "</Ratings>") :
                                                (hastag1Rating ?
                                                    ("<Ratings>" + tag1Rating.OuterXml + "</Ratings>") :
                                                    (hastag2Rating ?
                                                        ("<Ratings>" + tag2Rating.OuterXml + "</Ratings>") :
                                                        ""))
                                            , studentID
                                            , setting.SchoolYear
                                            , setting.Semester
                                        ));
                                        if (sql.Count == 200)
                                        {
                                            updateHelper.Execute(sql);
                                            sql.Clear();
                                        }
                                    }
                                    #endregion
                                }
                                if (sql.Count > 0)
                                {
                                    updateHelper.Execute(sql);
                                    sql.Clear();
                                }
                                #endregion
                            }
                            #endregion
                        }
                    }
                    catch (Exception ex)
                    {
                        exc = ex;
                    }
                };
                bkw.RunWorkerAsync();
            }


        }
    }
}
