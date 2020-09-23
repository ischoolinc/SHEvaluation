using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.Customization.Data;
using System.IO;
using FISCA.DSAUtil;
using System.Data;
using Aspose.Words;
using SmartSchool;
using FISCA.Presentation;
using System.Xml.Linq;

namespace SHStaticRank2.Data
{
    public class CalcMutilSemeSubjectRank
    {
        public delegate void OneClassCompleteDelegate();
        public static event OneClassCompleteDelegate OneClassCompleted;
        public static DataTable _table = new DataTable();
        public static string FolderName = "";
        static Dictionary<string, List<string>> SubjMappingDict = new Dictionary<string, List<string>>();


        public static void Setup(SHStaticRank2.Data.Configure setting)
        {

            // 計算學生科目數
            Dictionary<string, List<string>> _studSubjCountDict = new Dictionary<string, List<string>>();
            Dictionary<string, string> _studNameDict = new Dictionary<string, string>();
            Dictionary<string, Aspose.Words.Document> _WordDocDict = new Dictionary<string, Aspose.Words.Document>();
            Dictionary<string, Aspose.Cells.Workbook> _ExcelCellDict = new Dictionary<string, Aspose.Cells.Workbook>();
            List<memoText> _memoText = new List<memoText>();
            List<string> _PPSubjNameList = new List<string>();

            AccessHelper accessHelper = new AccessHelper();
            UpdateHelper updateHelper = new UpdateHelper();
            Dictionary<string, List<StudentRecord>> gradeyearStudents = new Dictionary<string, List<StudentRecord>>();
            System.ComponentModel.BackgroundWorker bkw = new System.ComponentModel.BackgroundWorker();
            bkw.WorkerReportsProgress = true;
            bkw.ProgressChanged += delegate (object sender, System.ComponentModel.ProgressChangedEventArgs e)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("多學期科目成績固定排名計算中...", e.ProgressPercentage);
            };
            Exception exc = null;
            bkw.RunWorkerCompleted += delegate
            {

                // 檢查學生科目數是否超過樣板可容納數
                List<string> ErrorStudIDList = new List<string>();

                foreach (KeyValuePair<string, List<string>> data in _studSubjCountDict)
                {
                    if (data.Value.Count > setting.SubjectLimit)
                        ErrorStudIDList.Add(data.Key);
                }

                if (ErrorStudIDList.Count > 0)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("學生科目數超出範本科目數，清單放在工作表:[學生科目數超出範本科目數清單]");


                    int rowIdx = 0;
                    foreach (string sid in ErrorStudIDList)
                    {
                        if (_studNameDict.ContainsKey(sid))
                        {
                            //wb.Worksheets["學生科目數超出範本科目數清單"].Cells[rowIdx, 0].PutValue(_studNameDict[sid]);
                            rowIdx++;
                        }
                    }
                }
                else
                {
                    //wb.Worksheets.RemoveAt("學生科目數超出範本科目數清單");
                }

                FISCA.Presentation.MotherForm.SetStatusBarMessage("多學期科目成績固定排名計算完成。", 100);
                // 清除暫存
                _ExcelCellDict.Clear();
                _WordDocDict.Clear();
                _memoText.Clear();
                GC.Collect();
                if (exc != null)
                {
                    throw new Exception("產生期末成績單發生錯誤", exc);
                }

                System.Diagnostics.Process.Start(Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports", FolderName));

            };
            bkw.DoWork += delegate
            {
                try
                {

                    FolderName = "多學期科目成績固定排名" + DateTime.Now.Year + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day) + string.Format("{0:00}", DateTime.Now.Hour) + string.Format("{0:00}", DateTime.Now.Minute);

                    int yearCount = 0;
                    bkw.ReportProgress(1);

                    if (setting.Name == "班級歷年成績單" || setting.Name == "教務作業_報表_班級歷年成績單")
                    {
                        setting.CheckExportStudent = false;
                    }
                    else
                    {
                        setting.CheckExportStudent = true;
                    }




                    #region 整各年級學生
                    foreach (var studentRec in accessHelper.StudentHelper.GetAllStudent())
                    {
                        if (studentRec.Status != "一般")
                            continue;
                        if (studentRec.RefClass != null)
                        {
                            string grade = "";
                            grade = "" + studentRec.RefClass.GradeYear;
                            if ((setting.CalcGradeYear1 && grade == "1") || (setting.CalcGradeYear2 && grade == "2") || (setting.CalcGradeYear3 && grade == "3") || (setting.CalcGradeYear4 && grade == "4"))
                            {
                                if (!gradeyearStudents.ContainsKey(grade))
                                    gradeyearStudents.Add(grade, new List<StudentRecord>());
                                gradeyearStudents[grade].Add(studentRec);
                            }
                        }
                    }
                    _ExcelCellDict.Clear();
                    _memoText.Clear();
                    _PPSubjNameList.Clear();


                    #region 取得學生成績資料
                    foreach (var gradeyear in gradeyearStudents.Keys)
                    {
                        var studentList = gradeyearStudents[gradeyear];
                        //取得學生學期科目成績
                        accessHelper.StudentHelper.FillSemesterSubjectScore(true, studentList);
                        if (setting.計算學業成績排名)
                        {
                            accessHelper.StudentHelper.FillSemesterEntryScore(true, studentList);
                        }
                    }
                    #endregion


                    // 學生類別一、類別二名稱
                    Dictionary<string, string> cat1Dict = new Dictionary<string, string>();
                    Dictionary<string, string> cat2Dict = new Dictionary<string, string>();
                    List<string> noRankList = new List<string>();
                    List<string> cat1List = new List<string>();
                    List<string> cat2List = new List<string>();
                    #region 取得學生類別對照
                    foreach (var gradeyear in gradeyearStudents.Keys)
                    {
                        var studentList = gradeyearStudents[gradeyear];

                        #region 分析學生所屬類別
                        foreach (var studentRec in studentList)
                        {
                            if (!cat1Dict.ContainsKey(studentRec.StudentID))
                                cat1Dict.Add(studentRec.StudentID, "");

                            if (!cat2Dict.ContainsKey(studentRec.StudentID))
                                cat2Dict.Add(studentRec.StudentID, "");

                            foreach (var tag in studentRec.StudentCategorys)
                            {
                                if (tag.SubCategory == "")
                                {
                                    if (setting.NotRankTag != "" && setting.NotRankTag == tag.Name)
                                    {
                                        if (!noRankList.Contains(studentRec.StudentID)) noRankList.Add(studentRec.StudentID);
                                    }
                                    if (setting.Rank1Tag != "" && setting.Rank1Tag == tag.Name)
                                    {
                                        if (!studentRec.Fields.ContainsKey("tag1"))
                                            studentRec.Fields.Add("tag1", tag.Name);
                                        if (!cat1List.Contains(tag.Name)) cat1List.Add(tag.Name);
                                    }
                                    if (setting.Rank2Tag != "" && setting.Rank2Tag == tag.Name)
                                    {
                                        if (!studentRec.Fields.ContainsKey("tag2"))
                                            studentRec.Fields.Add("tag2", tag.Name);
                                        if (!cat2List.Contains(tag.Name)) cat2List.Add(tag.Name);
                                    }
                                }
                                else
                                {
                                    if (setting.NotRankTag != "" && setting.NotRankTag == "[" + tag.Name + "]")
                                    {
                                        if (!noRankList.Contains(studentRec.StudentID)) noRankList.Add(studentRec.StudentID);
                                    }
                                    if (setting.Rank1Tag != "" && setting.Rank1Tag == "[" + tag.Name + "]")
                                    {
                                        if (!studentRec.Fields.ContainsKey("tag1"))
                                            studentRec.Fields.Add("tag1", tag.SubCategory);
                                        if (!cat1List.Contains(tag.SubCategory)) cat1List.Add(tag.SubCategory);
                                    }
                                    if (setting.Rank2Tag != "" && setting.Rank2Tag == "[" + tag.Name + "]")
                                    {
                                        if (!studentRec.Fields.ContainsKey("tag2"))
                                            studentRec.Fields.Add("tag2", tag.SubCategory);
                                        if (!cat2List.Contains(tag.SubCategory)) cat2List.Add(tag.SubCategory);
                                    }
                                }
                            }

                            // 比對學生類別與類別一、二
                            string tag1 = "", tag2 = "";
                            if (studentRec.Fields.ContainsKey("tag1"))
                                tag1 = studentRec.Fields["tag1"].ToString();

                            if (studentRec.Fields.ContainsKey("tag2"))
                                tag2 = studentRec.Fields["tag2"].ToString();

                            if (cat1List.Contains(tag1))
                                cat1Dict[studentRec.StudentID] = tag1;

                            if (cat2List.Contains(tag2))
                                cat2Dict[studentRec.StudentID] = tag2;
                        }

                        #endregion
                    }
                    #endregion

                    List<string> ssList = new List<string>();
                    int scc = 1;

                    var replace部訂必修專業及實習科目 = false;
                    var replaceTag1部訂必修專業及實習科目 = false;
                    var replaceTag2部訂必修專業及實習科目 = false;

                    #region 部訂必修專業及實習科目
                    // 檢查勾選科目
                    if (setting.useSubjectPrintList.Contains("部訂必修專業及實習科目"))
                    {
                        setting.useSubjectPrintList.Clear();
                        replace部訂必修專業及實習科目 = true;
                    }
                    // 檢查類別一勾選科目
                    if (!string.IsNullOrEmpty(setting.Rank1Tag) && setting.useSubjecOrder1List.Contains("部訂必修專業及實習科目"))
                    {
                        setting.useSubjecOrder1List.Clear();
                        replaceTag1部訂必修專業及實習科目 = true;
                    }

                    // 檢查類別二勾選科目
                    if (!string.IsNullOrEmpty(setting.Rank2Tag) && setting.useSubjecOrder2List.Contains("部訂必修專業及實習科目"))
                    {
                        setting.useSubjecOrder2List.Clear();
                        replaceTag2部訂必修專業及實習科目 = true;
                    }
                    #endregion

                    #region 更動勾選的科目
                    if (replace部訂必修專業及實習科目 || replaceTag1部訂必修專業及實習科目 || replaceTag2部訂必修專業及實習科目)
                    {
                        foreach (var gradeyear in gradeyearStudents.Keys)
                        {
                            foreach (StudentRecord student in gradeyearStudents[gradeyear])
                            {
                                foreach (var subjectScore in student.SemesterSubjectScoreList)
                                {
                                    // 當成績不在勾選年級學期跳過
                                    string gs = subjectScore.GradeYear + "" + subjectScore.Semester;
                                    if (!setting.useGradeSemesterList.Contains(gs))
                                        continue;
                                    if (subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                        && subjectScore.Require == true
                                        && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                    {
                                        if (replace部訂必修專業及實習科目)
                                            if (!setting.useSubjectPrintList.Contains(subjectScore.Subject))
                                                setting.useSubjectPrintList.Add(subjectScore.Subject);

                                        if (replaceTag1部訂必修專業及實習科目)
                                            if (!setting.useSubjecOrder1List.Contains(subjectScore.Subject))
                                                setting.useSubjecOrder1List.Add(subjectScore.Subject);

                                        if (replaceTag2部訂必修專業及實習科目)
                                            if (!setting.useSubjecOrder2List.Contains(subjectScore.Subject))
                                                setting.useSubjecOrder2List.Add(subjectScore.Subject);
                                    }
                                }
                            }
                        }
                    }
                    #endregion 更動勾選的科目

                    foreach (string subjName in setting.useSubjectPrintList)
                    {
                        if (scc < 7)
                        {
                            ssList.Add(subjName);

                        }
                        else
                        {
                            _PPSubjNameList.Add("科目：" + string.Join(",", ssList.ToArray()));
                            scc = 1;
                            ssList.Clear();
                            ssList.Add(subjName);
                        }
                        scc++;
                    }

                    if (ssList.Count > 0)
                        _PPSubjNameList.Add("科目：" + string.Join(",", ssList.ToArray()));

                    int ss1 = 1;
                    foreach (string str in _PPSubjNameList)
                    {
                        memoText mt = new memoText();
                        mt.Memo = str;
                        mt.WorkSheetName = "科目排名" + ss1;
                        _memoText.Add(mt);
                        ss1++;
                    }


                    #endregion

                    bkw.ReportProgress(20);

                    Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
                    foreach (var gradeyear in gradeyearStudents.Keys)
                    {
                        // 學期學業
                        if (setting.計算學業成績排名)
                        {
                            // 學業
                            #region Excel標題_學業
                            string shtNamne = "" + gradeyear + "年級學業";
                            if (!_ExcelCellDict.ContainsKey(shtNamne))
                            {
                                _ExcelCellDict.Add(shtNamne, new Aspose.Cells.Workbook());
                                memoText mt = new memoText();
                                mt.WorkSheetName = shtNamne;
                                mt.Memo = shtNamne;
                                _memoText.Add(mt);
                            }
                            _ExcelCellDict[shtNamne].Worksheets[0].Name = shtNamne;
                            var studentSheet2 = _ExcelCellDict[shtNamne].Worksheets[0];

                            //var studentSheet2 = wb.Worksheets[wb.Worksheets.Add()];
                            //studentSheet2.Name = "" + gradeyear + "年級學業";
                            int index2 = 0;

                            studentSheet2.Cells[0, index2++].PutValue("學號");
                            studentSheet2.Cells[0, index2++].PutValue("班級");
                            studentSheet2.Cells[0, index2++].PutValue("座號");
                            studentSheet2.Cells[0, index2++].PutValue("姓名");
                            studentSheet2.Cells[0, index2++].PutValue("類別一分類");
                            studentSheet2.Cells[0, index2++].PutValue("類別二分類");
                            studentSheet2.Cells[0, index2++].PutValue("一上");
                            studentSheet2.Cells[0, index2++].PutValue("一下");
                            studentSheet2.Cells[0, index2++].PutValue("二上");
                            studentSheet2.Cells[0, index2++].PutValue("二下");
                            studentSheet2.Cells[0, index2++].PutValue("三上");
                            studentSheet2.Cells[0, index2++].PutValue("三下");
                            studentSheet2.Cells[0, index2++].PutValue("四上");
                            studentSheet2.Cells[0, index2++].PutValue("四下");
                            studentSheet2.Cells[0, index2++].PutValue("總分");
                            studentSheet2.Cells[0, index2++].PutValue("總分班排名");
                            studentSheet2.Cells[0, index2++].PutValue("總分科排名");
                            studentSheet2.Cells[0, index2++].PutValue("總分校排名");
                            studentSheet2.Cells[0, index2++].PutValue("總分類別一排名");
                            studentSheet2.Cells[0, index2++].PutValue("總分類別二排名");
                            studentSheet2.Cells[0, index2++].PutValue("平均");
                            studentSheet2.Cells[0, index2++].PutValue("平均班排名");
                            studentSheet2.Cells[0, index2++].PutValue("平均科排名");
                            studentSheet2.Cells[0, index2++].PutValue("平均校排名");
                            studentSheet2.Cells[0, index2++].PutValue("平均類別一排名");
                            studentSheet2.Cells[0, index2++].PutValue("平均類別二排名");
                            studentSheet2.Cells[0, index2++].PutValue("加權總分");
                            studentSheet2.Cells[0, index2++].PutValue("加權總分班排名");
                            studentSheet2.Cells[0, index2++].PutValue("加權總分科排名");
                            studentSheet2.Cells[0, index2++].PutValue("加權總分校排名");
                            studentSheet2.Cells[0, index2++].PutValue("加權總分類別一排名");
                            studentSheet2.Cells[0, index2++].PutValue("加權總分類別二排名");
                            studentSheet2.Cells[0, index2++].PutValue("加權平均");
                            studentSheet2.Cells[0, index2++].PutValue("加權平均班排名");
                            studentSheet2.Cells[0, index2++].PutValue("加權平均科排名");
                            studentSheet2.Cells[0, index2++].PutValue("加權平均校排名");
                            studentSheet2.Cells[0, index2++].PutValue("加權平均類別一排名");
                            studentSheet2.Cells[0, index2++].PutValue("加權平均類別二排名");
                            #endregion Excel標題_學業

                            // 小郭, 2014/01/02
                            #region Excel標題_學業(原始)
                            string shtName_5 = "" + gradeyear + "年級學業(原始)";
                            if (!_ExcelCellDict.ContainsKey(shtName_5))
                            {
                                _ExcelCellDict.Add(shtName_5, new Aspose.Cells.Workbook());
                                memoText mt = new memoText();
                                mt.WorkSheetName = shtName_5;
                                mt.Memo = shtName_5;
                                _memoText.Add(mt);
                            }
                            _ExcelCellDict[shtName_5].Worksheets[0].Name = shtName_5;
                            var studentSheet2_5 = _ExcelCellDict[shtName_5].Worksheets[0];

                            int index2_5 = 0;

                            studentSheet2_5.Cells[0, index2_5++].PutValue("學號");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("班級");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("座號");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("姓名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("類別一分類");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("類別二分類");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("一上");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("一下");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("二上");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("二下");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("三上");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("三下");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("四上");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("四下");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("總分");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("總分班排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("總分科排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("總分校排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("總分類別一排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("總分類別二排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("平均");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("平均班排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("平均科排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("平均校排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("平均類別一排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("平均類別二排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權總分");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權總分班排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權總分科排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權總分校排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權總分類別一排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權總分類別二排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權平均");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權平均班排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權平均科排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權平均校排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權平均類別一排名");
                            studentSheet2_5.Cells[0, index2_5++].PutValue("加權平均類別二排名");
                            #endregion Excel標題_學業(原始)

                            // 體育
                            #region Excel標題_體育
                            string shtName_1 = "" + gradeyear + "年級學業體育";
                            if (!_ExcelCellDict.ContainsKey(shtName_1))
                            {
                                _ExcelCellDict.Add(shtName_1, new Aspose.Cells.Workbook());
                                memoText mt = new memoText();
                                mt.WorkSheetName = shtName_1;
                                mt.Memo = shtName_1;
                                _memoText.Add(mt);
                            }
                            _ExcelCellDict[shtName_1].Worksheets[0].Name = shtName_1;
                            var studentSheet2_1 = _ExcelCellDict[shtName_1].Worksheets[0];
                            //var studentSheet2_1 = wb.Worksheets[wb.Worksheets.Add()];
                            //studentSheet2_1.Name = "" + gradeyear + "年級學業體育";
                            int index2_1 = 0;

                            studentSheet2_1.Cells[0, index2_1++].PutValue("學號");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("班級");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("座號");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("姓名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("類別一分類");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("類別二分類");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("一上");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("一下");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("二上");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("二下");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("三上");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("三下");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("四上");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("四下");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("總分");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("總分班排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("總分科排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("總分校排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("總分類別一排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("總分類別二排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("平均");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("平均班排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("平均科排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("平均校排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("平均類別一排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("平均類別二排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權總分");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權總分班排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權總分科排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權總分校排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權總分類別一排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權總分類別二排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權平均");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權平均班排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權平均科排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權平均校排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權平均類別一排名");
                            studentSheet2_1.Cells[0, index2_1++].PutValue("加權平均類別二排名");
                            #endregion Excel標題_體育

                            // 健康與護理
                            #region Excel標題_健康與護理
                            string shtName_2 = "" + gradeyear + "年級學業健康與護理";
                            if (!_ExcelCellDict.ContainsKey(shtName_2))
                            {
                                _ExcelCellDict.Add(shtName_2, new Aspose.Cells.Workbook());
                                memoText mt = new memoText();
                                mt.WorkSheetName = shtName_2;
                                mt.Memo = shtName_2;
                                _memoText.Add(mt);
                            }
                            _ExcelCellDict[shtName_2].Worksheets[0].Name = shtName_2;
                            var studentSheet2_2 = _ExcelCellDict[shtName_2].Worksheets[0];
                            //var studentSheet2_2 = wb.Worksheets[wb.Worksheets.Add()];
                            //studentSheet2_2.Name = "" + gradeyear + "年級學業健康與護理";
                            int index2_2 = 0;

                            studentSheet2_2.Cells[0, index2_2++].PutValue("學號");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("班級");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("座號");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("姓名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("類別一分類");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("類別二分類");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("一上");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("一下");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("二上");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("二下");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("三上");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("三下");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("四上");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("四下");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("總分");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("總分班排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("總分科排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("總分校排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("總分類別一排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("總分類別二排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("平均");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("平均班排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("平均科排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("平均校排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("平均類別一排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("平均類別二排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權總分");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權總分班排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權總分科排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權總分校排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權總分類別一排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權總分類別二排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權平均");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權平均班排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權平均科排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權平均校排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權平均類別一排名");
                            studentSheet2_2.Cells[0, index2_2++].PutValue("加權平均類別二排名");
                            #endregion Excel標題_健康與護理

                            // 國防通識
                            #region Excel標題_國防通識
                            string shtName_3 = "" + gradeyear + "年級學業國防通識";
                            if (!_ExcelCellDict.ContainsKey(shtName_3))
                            {
                                _ExcelCellDict.Add(shtName_3, new Aspose.Cells.Workbook());
                                memoText mt = new memoText();
                                mt.WorkSheetName = shtName_3;
                                mt.Memo = shtName_3;
                                _memoText.Add(mt);
                            }
                            _ExcelCellDict[shtName_3].Worksheets[0].Name = shtName_3;
                            var studentSheet2_3 = _ExcelCellDict[shtName_3].Worksheets[0];

                            //var studentSheet2_3 = wb.Worksheets[wb.Worksheets.Add()];
                            //studentSheet2_3.Name = "" + gradeyear + "年級學業國防通識";
                            int index2_3 = 0;

                            studentSheet2_3.Cells[0, index2_3++].PutValue("學號");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("班級");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("座號");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("姓名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("類別一分類");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("類別二分類");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("一上");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("一下");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("二上");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("二下");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("三上");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("三下");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("四上");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("四下");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("總分");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("總分班排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("總分科排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("總分校排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("總分類別一排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("總分類別二排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("平均");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("平均班排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("平均科排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("平均校排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("平均類別一排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("平均類別二排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權總分");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權總分班排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權總分科排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權總分校排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權總分類別一排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權總分類別二排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權平均");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權平均班排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權平均科排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權平均校排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權平均類別一排名");
                            studentSheet2_3.Cells[0, index2_3++].PutValue("加權平均類別二排名");
                            #endregion Excel標題_國防通識

                            // 實習科目
                            #region Excel標題_實習科目
                            string shtName_4 = "" + gradeyear + "年級學業實習科目";
                            if (!_ExcelCellDict.ContainsKey(shtName_4))
                            {
                                _ExcelCellDict.Add(shtName_4, new Aspose.Cells.Workbook());
                                memoText mt = new memoText();
                                mt.WorkSheetName = shtName_4;
                                mt.Memo = shtName_4;
                                _memoText.Add(mt);
                            }
                            _ExcelCellDict[shtName_4].Worksheets[0].Name = shtName_4;
                            var studentSheet2_4 = _ExcelCellDict[shtName_4].Worksheets[0];

                            //var studentSheet2_4 = wb.Worksheets[wb.Worksheets.Add()];
                            //studentSheet2_4.Name = "" + gradeyear + "年級學業實習科目";
                            int index2_4 = 0;

                            studentSheet2_4.Cells[0, index2_4++].PutValue("學號");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("班級");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("座號");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("姓名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("類別一分類");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("類別二分類");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("一上");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("一下");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("二上");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("二下");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("三上");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("三下");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("四上");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("四下");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("總分");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("總分班排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("總分科排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("總分校排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("總分類別一排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("總分類別二排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("平均");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("平均班排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("平均科排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("平均校排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("平均類別一排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("平均類別二排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權總分");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權總分班排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權總分科排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權總分校排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權總分類別一排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權總分類別二排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權平均");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權平均班排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權平均科排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權平均校排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權平均類別一排名");
                            studentSheet2_4.Cells[0, index2_4++].PutValue("加權平均類別二排名");
                            #endregion Excel標題_實習科目
                        }


                        // 總計成績
                        #region 總計成績
                        string shtNameSS = "" + gradeyear + "年級總計成績";
                        if (!_ExcelCellDict.ContainsKey(shtNameSS))
                        {
                            _ExcelCellDict.Add(shtNameSS, new Aspose.Cells.Workbook());
                            memoText mt = new memoText();
                            mt.WorkSheetName = shtNameSS;
                            mt.Memo = shtNameSS;
                            _memoText.Add(mt);
                        }
                        _ExcelCellDict[shtNameSS].Worksheets[0].Name = shtNameSS;
                        var studentSheet3 = _ExcelCellDict[shtNameSS].Worksheets[0];

                        //var studentSheet3 = wb.Worksheets[wb.Worksheets.Add()];
                        //studentSheet3.Name = "" + gradeyear + "年級總計成績";
                        int index3 = 0;
                        studentSheet3.Cells[0, index3++].PutValue("學號");
                        studentSheet3.Cells[0, index3++].PutValue("班級");
                        studentSheet3.Cells[0, index3++].PutValue("座號");
                        studentSheet3.Cells[0, index3++].PutValue("姓名");
                        studentSheet3.Cells[0, index3++].PutValue("類別一分類");
                        studentSheet3.Cells[0, index3++].PutValue("類別二分類");
                        studentSheet3.Cells[0, index3++].PutValue("總分(類別一)");
                        studentSheet3.Cells[0, index3++].PutValue("總分類別一排名");
                        studentSheet3.Cells[0, index3++].PutValue("平均(類別一)");
                        studentSheet3.Cells[0, index3++].PutValue("平均類別一排名");
                        studentSheet3.Cells[0, index3++].PutValue("加權總分(類別一)");
                        studentSheet3.Cells[0, index3++].PutValue("加權總分類別一排名");
                        studentSheet3.Cells[0, index3++].PutValue("加權平均(類別一)");
                        studentSheet3.Cells[0, index3++].PutValue("加權平均類別一排名");
                        studentSheet3.Cells[0, index3++].PutValue("總分(類別二)");
                        studentSheet3.Cells[0, index3++].PutValue("總分類別二排名");
                        studentSheet3.Cells[0, index3++].PutValue("平均(類別二)");
                        studentSheet3.Cells[0, index3++].PutValue("平均類別二排名");
                        studentSheet3.Cells[0, index3++].PutValue("加權總分(類別二)");
                        studentSheet3.Cells[0, index3++].PutValue("加權總分類別二排名");
                        studentSheet3.Cells[0, index3++].PutValue("加權平均(類別二)");
                        studentSheet3.Cells[0, index3++].PutValue("加權平均類別二排名");
                        studentSheet3.Cells[0, index3++].PutValue("總分");
                        studentSheet3.Cells[0, index3++].PutValue("總分班排名");
                        studentSheet3.Cells[0, index3++].PutValue("總分科排名");
                        studentSheet3.Cells[0, index3++].PutValue("總分校排名");
                        studentSheet3.Cells[0, index3++].PutValue("平均");
                        studentSheet3.Cells[0, index3++].PutValue("平均班排名");
                        studentSheet3.Cells[0, index3++].PutValue("平均科排名");
                        studentSheet3.Cells[0, index3++].PutValue("平均校排名");
                        studentSheet3.Cells[0, index3++].PutValue("加權總分");
                        studentSheet3.Cells[0, index3++].PutValue("加權總分班排名");
                        studentSheet3.Cells[0, index3++].PutValue("加權總分科排名");
                        studentSheet3.Cells[0, index3++].PutValue("加權總分校排名");
                        studentSheet3.Cells[0, index3++].PutValue("加權平均");
                        studentSheet3.Cells[0, index3++].PutValue("加權平均班排名");
                        studentSheet3.Cells[0, index3++].PutValue("加權平均科排名");
                        studentSheet3.Cells[0, index3++].PutValue("加權平均校排名");
                        #endregion 總計成績

                        // 科目排名
                        #region 科目排名
                        string shtName1 = "科目排名";
                        int subjCount = 1;
                        int subjnameIdx = 1;
                        int BeginColumn = 0;
                        string shtNameSubj = shtName1 + subjnameIdx;
                        if (!_ExcelCellDict.ContainsKey(shtNameSubj))
                        {
                            _ExcelCellDict.Add(shtNameSubj, new Aspose.Cells.Workbook());
                            memoText mt = new memoText();
                            mt.WorkSheetName = shtNameSubj;
                            _memoText.Add(mt);

                        }
                        _ExcelCellDict[shtNameSubj].Worksheets[0].Name = shtNameSubj;
                        var studentSheet = _ExcelCellDict[shtNameSubj].Worksheets[0];

                        //var studentSheet = wb.Worksheets[wb.Worksheets.Add()];
                        //studentSheet.Name = shtName1 + subjnameIdx;
                        foreach (string subjName in setting.useSubjectPrintList)
                        {
                            if (subjCount > 6)
                            {
                                subjCount = 1;
                                subjnameIdx++;
                                BeginColumn = 0;

                                string shtNameSubj1 = shtName1 + subjnameIdx;
                                if (!_ExcelCellDict.ContainsKey(shtNameSubj1))
                                {
                                    _ExcelCellDict.Add(shtNameSubj1, new Aspose.Cells.Workbook());
                                    memoText mt = new memoText();
                                    mt.WorkSheetName = shtNameSubj1;
                                    _memoText.Add(mt);
                                }
                                studentSheet = _ExcelCellDict[shtNameSubj1].Worksheets[0];
                                studentSheet.Name = shtNameSubj1;
                                //studentSheet = wb.Worksheets[wb.Worksheets.Add()];
                                //studentSheet.Name = shtName1 + subjnameIdx;
                            }

                            yearCount++;
                            int index = BeginColumn;
                            studentSheet.Cells[0, index++].PutValue("年級科目名稱");
                            studentSheet.Cells[0, index++].PutValue("學號");
                            studentSheet.Cells[0, index++].PutValue("班級");
                            studentSheet.Cells[0, index++].PutValue("座號");
                            studentSheet.Cells[0, index++].PutValue("姓名");
                            studentSheet.Cells[0, index++].PutValue("類別一分類");
                            studentSheet.Cells[0, index++].PutValue("類別二分類");
                            studentSheet.Cells[0, index++].PutValue("一上");
                            studentSheet.Cells[0, index++].PutValue("一下");
                            studentSheet.Cells[0, index++].PutValue("二上");
                            studentSheet.Cells[0, index++].PutValue("二下");
                            studentSheet.Cells[0, index++].PutValue("三上");
                            studentSheet.Cells[0, index++].PutValue("三下");
                            studentSheet.Cells[0, index++].PutValue("四上");
                            studentSheet.Cells[0, index++].PutValue("四下");
                            studentSheet.Cells[0, index++].PutValue("總分");
                            studentSheet.Cells[0, index++].PutValue("總分班排名");
                            studentSheet.Cells[0, index++].PutValue("總分科排名");
                            studentSheet.Cells[0, index++].PutValue("總分校排名");
                            studentSheet.Cells[0, index++].PutValue("總分類別一排名");
                            studentSheet.Cells[0, index++].PutValue("總分類別二排名");
                            studentSheet.Cells[0, index++].PutValue("平均");
                            studentSheet.Cells[0, index++].PutValue("平均班排名");
                            studentSheet.Cells[0, index++].PutValue("平均科排名");
                            studentSheet.Cells[0, index++].PutValue("平均校排名");
                            studentSheet.Cells[0, index++].PutValue("平均類別一排名");
                            studentSheet.Cells[0, index++].PutValue("平均類別二排名");
                            studentSheet.Cells[0, index++].PutValue("加權總分");
                            studentSheet.Cells[0, index++].PutValue("加權總分班排名");
                            studentSheet.Cells[0, index++].PutValue("加權總分科排名");
                            studentSheet.Cells[0, index++].PutValue("加權總分校排名");
                            studentSheet.Cells[0, index++].PutValue("加權總分類別一排名");
                            studentSheet.Cells[0, index++].PutValue("加權總分類別二排名");
                            studentSheet.Cells[0, index++].PutValue("加權平均");
                            studentSheet.Cells[0, index++].PutValue("加權平均班排名");
                            studentSheet.Cells[0, index++].PutValue("加權平均科排名");
                            studentSheet.Cells[0, index++].PutValue("加權平均校排名");
                            studentSheet.Cells[0, index++].PutValue("加權平均類別一排名");
                            studentSheet.Cells[0, index++].PutValue("加權平均類別二排名");

                            subjCount++;
                            BeginColumn += 40;
                        }
                        #endregion 科目排名

                        var studentList = gradeyearStudents[gradeyear];

                        //bkw.ReportProgress(1 + 99 * yearCount / gradeyearStudents.Count / 2);
                        Dictionary<string, List<decimal>> ranks = new Dictionary<string, List<decimal>>();
                        Dictionary<string, List<string>> rankStudents = new Dictionary<string, List<string>>();
                        //Dictionary<SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo, decimal?> selectScore = new Dictionary<SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo, decimal?>();


                        // 儲存成績用
                        Dictionary<string, studScore> selectScore = new Dictionary<string, studScore>();


                        #region 解析回歸科目對照
                        SubjMappingDict.Clear();

                        try
                        {
                            // 解析科目對照
                            XElement elmRoot = XElement.Parse(setting.SubjectMapping);
                            foreach (XElement elm in elmRoot.Elements("Item"))
                            {
                                string Subj = elm.Attribute("Subject").Value;
                                string SysSubj = elm.Attribute("SysSubject").Value;
                                if (!SubjMappingDict.ContainsKey(Subj))
                                    SubjMappingDict.Add(Subj, new List<string>());

                                // 解析字串+
                                List<string> strList = SysSubj.Split('+').ToList();

                                foreach (string str in strList)
                                    SubjMappingDict[Subj].Add(str);
                            }

                            List<string> HasSubjNameList = new List<string>();
                            List<string> newMappingList = new List<string>();

                            foreach (string key in SubjMappingDict.Keys)
                            {
                                if (!HasSubjNameList.Contains(key))
                                    HasSubjNameList.Add(key);

                                foreach (string value in SubjMappingDict[key])
                                {
                                    if (!HasSubjNameList.Contains(value))
                                        HasSubjNameList.Add(value);
                                }
                            }

                            // 加入過濾後科目
                            foreach (var studentRec in studentList)
                            {
                                foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                {
                                    if (!HasSubjNameList.Contains(subjectScore.Subject))
                                    {
                                        if (!newMappingList.Contains(subjectScore.Subject))
                                            newMappingList.Add(subjectScore.Subject);
                                    }
                                }
                            }

                            newMappingList.Sort();

                            // 顯示全部科目`,非回歸科目放後
                            if (setting.CheckExportSubjectMapping == false)
                            {

                                foreach (string str in newMappingList)
                                {
                                    if (!SubjMappingDict.ContainsKey(str))
                                    {
                                        List<string> li = new List<string>();
                                        li.Add(str);
                                        SubjMappingDict.Add(str, li);
                                    }

                                }
                            }

                        }
                        catch (Exception ex)
                        { }

                        #endregion



                        foreach (var studentRec in studentList)
                        {
                            string studentID = studentRec.StudentID;
                            string name = "學號：" + studentRec.StudentNumber + ",班級:" + studentRec.RefClass.ClassName + ",座號：" + studentRec.SeatNo + ",姓名：" + studentRec.StudentName;
                            if (!_studNameDict.ContainsKey(studentID))
                                _studNameDict.Add(studentID, name);

                            // 處理列印科目成績
                            #region 處理列印科目成績

                            foreach (string SubjName in setting.useSubjectPrintList)
                            {
                                // 總分,加權總分,平均,加權平均
                                string subjKey = studentID + "^^^" + SubjName;
                                bool chkHasSubjectName = false;
                                bool chkHasTag1 = false, chkHasTag2 = false;

                                #region 處理學期科目成績
                                foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                {

                                    // 當成績不在勾選年級學期，跳過
                                    string gs = subjectScore.GradeYear + "" + subjectScore.Semester;
                                    if (!setting.useGradeSemesterList.Contains(gs))
                                        continue;
                                    // 判斷此科目是否為需要產出的
                                    if (subjectScore.Subject == SubjName
                                        && (
                                            replace部訂必修專業及實習科目 == false ||
                                            (
                                                subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                && subjectScore.Require == true
                                                && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                            )
                                        )
                                    {
                                        // 計算學生科目數
                                        if (!_studSubjCountDict.ContainsKey(studentID))
                                            _studSubjCountDict.Add(studentID, new List<string>());
                                        if (!_studSubjCountDict[studentID].Contains(subjectScore.Subject))
                                            _studSubjCountDict[studentID].Add(subjectScore.Subject);

                                        chkHasSubjectName = true;
                                        if (!selectScore.ContainsKey(subjKey))
                                            selectScore.Add(subjKey, new studScore());

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

                                        // 沒有任何成績跳過
                                        if (match == false)
                                            continue;

                                        // 記錄科目學期成績
                                        if (subjectScore.GradeYear == 1 && subjectScore.Semester == 1)
                                        {
                                            selectScore[subjKey].gsScore11 = score;
                                            selectScore[subjKey].gsSchoolYear11 = subjectScore.SchoolYear;
                                            selectScore[subjKey].gsCredit11 = subjectScore.CreditDec();
                                        }
                                        if (subjectScore.GradeYear == 1 && subjectScore.Semester == 2)
                                        {
                                            selectScore[subjKey].gsScore12 = score;
                                            selectScore[subjKey].gsSchoolYear12 = subjectScore.SchoolYear;
                                            selectScore[subjKey].gsCredit12 = subjectScore.CreditDec();
                                        }
                                        if (subjectScore.GradeYear == 2 && subjectScore.Semester == 1)
                                        {
                                            selectScore[subjKey].gsScore21 = score;
                                            selectScore[subjKey].gsSchoolYear21 = subjectScore.SchoolYear;
                                            selectScore[subjKey].gsCredit21 = subjectScore.CreditDec();
                                        }

                                        if (subjectScore.GradeYear == 2 && subjectScore.Semester == 2)
                                        {
                                            selectScore[subjKey].gsScore22 = score;
                                            selectScore[subjKey].gsSchoolYear22 = subjectScore.SchoolYear;
                                            selectScore[subjKey].gsCredit22 = subjectScore.CreditDec();
                                        }
                                        if (subjectScore.GradeYear == 3 && subjectScore.Semester == 1)
                                        {
                                            selectScore[subjKey].gsScore31 = score;
                                            selectScore[subjKey].gsSchoolYear31 = subjectScore.SchoolYear;
                                            selectScore[subjKey].gsCredit31 = subjectScore.CreditDec();
                                        }
                                        if (subjectScore.GradeYear == 3 && subjectScore.Semester == 2)
                                        {
                                            selectScore[subjKey].gsScore32 = score;
                                            selectScore[subjKey].gsSchoolYear32 = subjectScore.SchoolYear;
                                            selectScore[subjKey].gsCredit32 = subjectScore.CreditDec();
                                        }

                                        if (subjectScore.GradeYear == 4 && subjectScore.Semester == 1)
                                        {
                                            selectScore[subjKey].gsScore41 = score;
                                            selectScore[subjKey].gsSchoolYear41 = subjectScore.SchoolYear;
                                            selectScore[subjKey].gsCredit41 = subjectScore.CreditDec();
                                        }
                                        if (subjectScore.GradeYear == 4 && subjectScore.Semester == 2)
                                        {
                                            selectScore[subjKey].gsScore42 = score;
                                            selectScore[subjKey].gsSchoolYear42 = subjectScore.SchoolYear;
                                            selectScore[subjKey].gsCredit42 = subjectScore.CreditDec();
                                        }
                                        // 總分
                                        selectScore[subjKey].sumScore += score;
                                        // 總分加權
                                        selectScore[subjKey].sumScoreA += (score * subjectScore.CreditDec());
                                        // 筆數
                                        selectScore[subjKey].subjCount++;
                                        // 學分加總
                                        selectScore[subjKey].sumCredit += subjectScore.CreditDec();

                                        // 類別一處理, 判斷此科目是否為類別1需要的
                                        if (setting.useSubjecOrder1List.Contains(SubjName)
                                            && (
                                                replaceTag1部訂必修專業及實習科目 == false ||
                                                (
                                                    subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                    && subjectScore.Require == true
                                                    && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                )
                                            )
                                        {
                                            chkHasTag1 = true;
                                            // 總分
                                            selectScore[subjKey].sumScoreC1 += score;
                                            // 總分加權
                                            selectScore[subjKey].sumScoreAC1 += (score * subjectScore.CreditDec());
                                            // 筆數
                                            selectScore[subjKey].subjCountC1++;
                                            // 學分加總
                                            selectScore[subjKey].sumCreditC1 += subjectScore.CreditDec();
                                        }

                                        // 類別二處理, 判斷此科目是否為類別1需要的
                                        if (setting.useSubjecOrder2List.Contains(SubjName)
                                            && (
                                                replaceTag2部訂必修專業及實習科目 == false ||
                                                (
                                                    subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                    && subjectScore.Require == true
                                                    && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                )
                                            )
                                        {
                                            chkHasTag2 = true;
                                            // 總分
                                            selectScore[subjKey].sumScoreC2 += score;
                                            // 總分加權
                                            selectScore[subjKey].sumScoreAC2 += (score * subjectScore.CreditDec());
                                            // 筆數
                                            selectScore[subjKey].subjCountC2++;
                                            // 學分加總
                                            selectScore[subjKey].sumCreditC2 += subjectScore.CreditDec();
                                        }

                                    }
                                }
                                #endregion 處理學期科目成績
                                #region 處理單一科目多學期總計成績
                                if (chkHasSubjectName)
                                {
                                    // 平均
                                    if (selectScore[subjKey].subjCount > 0)
                                        selectScore[subjKey].avgScore = (decimal)(selectScore[subjKey].sumScore / selectScore[subjKey].subjCount);
                                    // 加權平均
                                    if (selectScore[subjKey].sumCredit > 0)
                                        selectScore[subjKey].avgScoreA = (decimal)(selectScore[subjKey].sumScoreA / selectScore[subjKey].sumCredit);

                                    // 平均(類別一)
                                    if (selectScore[subjKey].subjCountC1 > 0)
                                        selectScore[subjKey].avgScoreC1 = (decimal)(selectScore[subjKey].sumScoreC1 / selectScore[subjKey].subjCountC1);
                                    // 加權平均(類別一)
                                    if (selectScore[subjKey].sumCreditC1 > 0)
                                        selectScore[subjKey].avgScoreAC1 = (decimal)(selectScore[subjKey].sumScoreAC1 / selectScore[subjKey].sumCreditC1);

                                    // 平均(類別二)
                                    if (selectScore[subjKey].subjCountC2 > 0)
                                        selectScore[subjKey].avgScoreC2 = (decimal)(selectScore[subjKey].sumScoreC2 / selectScore[subjKey].subjCountC2);
                                    // 加權平均(類別二)
                                    if (selectScore[subjKey].sumCreditC2 > 0)
                                        selectScore[subjKey].avgScoreAC2 = (decimal)(selectScore[subjKey].sumScoreAC2 / selectScore[subjKey].sumCreditC2);

                                    if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                    {
                                        #region 班排名
                                        {
                                            string key1 = "總分班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[subjKey].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "加權總分班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[subjKey].sumScoreA);
                                            rankStudents[key2].Add(studentID);

                                            string key3 = "平均班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                            if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                            ranks[key3].Add(selectScore[subjKey].avgScore);
                                            rankStudents[key3].Add(studentID);


                                            string key4 = "加權平均班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                            if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                            ranks[key4].Add(selectScore[subjKey].avgScoreA);
                                            rankStudents[key4].Add(studentID);

                                        }
                                        #endregion
                                        #region 科排名
                                        {
                                            if (studentRec.Department != "")
                                            {
                                                //各科目科排名
                                                string key1 = "總分科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[subjKey].sumScore);
                                                rankStudents[key1].Add(studentID);

                                                //各科目科排名
                                                string key2 = "加權總分科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[subjKey].sumScoreA);
                                                rankStudents[key2].Add(studentID);

                                                //各科目科排名
                                                string key3 = "平均科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                                ranks[key3].Add(selectScore[subjKey].avgScore);
                                                rankStudents[key3].Add(studentID);

                                                //各科目科排名
                                                string key4 = "加權平均科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                                ranks[key4].Add(selectScore[subjKey].avgScoreA);
                                                rankStudents[key4].Add(studentID);

                                            }
                                        }
                                        #endregion
                                        #region 全校排名
                                        {
                                            string key1 = "總分全校排名" + gradeyear + "^^^" + SubjName;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[subjKey].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "加權總分全校排名" + gradeyear + "^^^" + SubjName;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[subjKey].sumScoreA);
                                            rankStudents[key2].Add(studentID);

                                            string key3 = "平均全校排名" + gradeyear + "^^^" + SubjName;
                                            if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                            ranks[key3].Add(selectScore[subjKey].avgScore);
                                            rankStudents[key3].Add(studentID);

                                            string key4 = "加權平均全校排名" + gradeyear + "^^^" + SubjName;
                                            if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                            ranks[key4].Add(selectScore[subjKey].avgScoreA);
                                            rankStudents[key4].Add(studentID);

                                        }
                                        #endregion
                                        #region 類別1排名
                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            // 假如上面有處理過類別1的科目
                                            if (chkHasTag1 == true)
                                            {
                                                string key1 = "總分類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[subjKey].sumScoreC1);
                                                rankStudents[key1].Add(studentID);

                                                string key2 = "加權總分類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[subjKey].sumScoreAC1);
                                                rankStudents[key2].Add(studentID);

                                                string key3 = "平均類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                                ranks[key3].Add(selectScore[subjKey].avgScoreC1);
                                                rankStudents[key3].Add(studentID);

                                                string key4 = "加權平均類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                                ranks[key4].Add(selectScore[subjKey].avgScoreAC1);
                                                rankStudents[key4].Add(studentID);
                                            }
                                        }
                                        #endregion
                                        #region 類別2排名
                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            // 假如上面有處理過類別2的科目
                                            if (chkHasTag2 == true)
                                            {
                                                string key1 = "總分類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[subjKey].sumScoreC2);
                                                rankStudents[key1].Add(studentID);

                                                string key2 = "加權總分類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[subjKey].sumScoreAC2);
                                                rankStudents[key2].Add(studentID);

                                                string key3 = "平均類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                                ranks[key3].Add(selectScore[subjKey].avgScoreC2);
                                                rankStudents[key3].Add(studentID);

                                                string key4 = "加權平均類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                                ranks[key4].Add(selectScore[subjKey].avgScoreAC2);
                                                rankStudents[key4].Add(studentID);
                                            }
                                        }
                                        #endregion
                                    }
                                }
                                #endregion
                            }
                            #endregion 處理列印科目成績


                            // 2015/10/22 因實中新增
                            #region 處理回歸科目成績

                            // 建立學生成績索引                      
                            try
                            {
                                #region 解析回歸科目成績並填入
                                foreach (string SubjName in SubjMappingDict.Keys)
                                {
                                    // 處理回歸總分,加權總分,平均,加權平均
                                    string subjKeyR1 = studentID + "^^^回歸" + SubjName;
                                    bool chkHasSubjectName = false;
                                    bool chkHasTag1 = false, chkHasTag2 = false;
                                    decimal sumScore11 = 0, sumCredit11 = 0, sumScore12 = 0, sumCredit12 = 0, sumScore21 = 0, sumCredit21 = 0, sumScore22 = 0, sumCredit22 = 0, sumScore31 = 0, sumCredit31 = 0, sumScore32 = 0, sumCredit32 = 0, sumScore41 = 0, sumCredit41 = 0, sumScore42 = 0, sumCredit42 = 0;

                                    #region 填入科目成績
                                    foreach (string MappingSubj in SubjMappingDict[SubjName])
                                    {
                                        decimal score = decimal.MinValue, tryParseScore;
                                        bool match = false;


                                        // 讀取學生學期成績資料
                                        foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                        {
                                            // 判斷此科目是否為需要產出的
                                            if (subjectScore.Subject == MappingSubj
                                                && (
                                                    replace部訂必修專業及實習科目 == false ||
                                                    (
                                                        subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                        && subjectScore.Require == true
                                                        && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                    )
                                                )
                                            {
                                                if (!selectScore.ContainsKey(subjKeyR1))
                                                    selectScore.Add(subjKeyR1, new studScore());

                                                score = 0;
                                                chkHasSubjectName = true;


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


                                                // 沒有任何成績跳過
                                                if (match == false)
                                                    continue;


                                                // 有科目再處理
                                                if (chkHasSubjectName)
                                                {
                                                    // 記錄科目學期成績
                                                    if (subjectScore.GradeYear == 1 && subjectScore.Semester == 1 && subjectScore.Subject == MappingSubj)
                                                    {
                                                        selectScore[subjKeyR1].gsSchoolYear11 = subjectScore.SchoolYear;
                                                        sumScore11 += score * subjectScore.CreditDec();
                                                        sumCredit11 += subjectScore.CreditDec();
                                                    }
                                                    if (subjectScore.GradeYear == 1 && subjectScore.Semester == 2 && subjectScore.Subject == MappingSubj)
                                                    {
                                                        selectScore[subjKeyR1].gsSchoolYear12 = subjectScore.SchoolYear;
                                                        sumScore12 += score * subjectScore.CreditDec();
                                                        sumCredit12 += subjectScore.CreditDec();

                                                    }
                                                    if (subjectScore.GradeYear == 2 && subjectScore.Semester == 1 && subjectScore.Subject == MappingSubj)
                                                    {
                                                        selectScore[subjKeyR1].gsSchoolYear21 = subjectScore.SchoolYear;
                                                        sumScore21 += score * subjectScore.CreditDec();
                                                        sumCredit21 += subjectScore.CreditDec();
                                                    }

                                                    if (subjectScore.GradeYear == 2 && subjectScore.Semester == 2 && subjectScore.Subject == MappingSubj)
                                                    {
                                                        selectScore[subjKeyR1].gsSchoolYear22 = subjectScore.SchoolYear;
                                                        sumScore22 += score * subjectScore.CreditDec();
                                                        sumCredit22 += subjectScore.CreditDec();
                                                    }
                                                    if (subjectScore.GradeYear == 3 && subjectScore.Semester == 1 && subjectScore.Subject == MappingSubj)
                                                    {
                                                        selectScore[subjKeyR1].gsSchoolYear31 = subjectScore.SchoolYear;
                                                        sumScore31 += score * subjectScore.CreditDec();
                                                        sumCredit31 += subjectScore.CreditDec();


                                                    }
                                                    if (subjectScore.GradeYear == 3 && subjectScore.Semester == 2 && subjectScore.Subject == MappingSubj)
                                                    {
                                                        selectScore[subjKeyR1].gsSchoolYear32 = subjectScore.SchoolYear;
                                                        sumScore32 += score * subjectScore.CreditDec();
                                                        sumCredit32 += subjectScore.CreditDec();

                                                    }

                                                    if (subjectScore.GradeYear == 4 && subjectScore.Semester == 1 && subjectScore.Subject == MappingSubj)
                                                    {
                                                        selectScore[subjKeyR1].gsSchoolYear41 = subjectScore.SchoolYear;
                                                        sumScore41 += score * subjectScore.CreditDec();
                                                        sumCredit41 += subjectScore.CreditDec();


                                                    }
                                                    if (subjectScore.GradeYear == 4 && subjectScore.Semester == 2 && subjectScore.Subject == MappingSubj)
                                                    {
                                                        selectScore[subjKeyR1].gsSchoolYear42 = subjectScore.SchoolYear;
                                                        sumScore42 += score * subjectScore.CreditDec();
                                                        sumCredit42 += subjectScore.CreditDec();

                                                    }
                                                }
                                            }
                                        }
                                    }


                                    if (selectScore.ContainsKey(subjKeyR1))
                                    {

                                        // 處理分數
                                        if (sumCredit11 > 0)
                                        {
                                            selectScore[subjKeyR1].gsScore11 = Math.Round(sumScore11 / sumCredit11, 0, MidpointRounding.AwayFromZero);
                                            selectScore[subjKeyR1].gsCredit11 = sumCredit11;
                                            selectScore[subjKeyR1].subjCount++;
                                            selectScore[subjKeyR1].sumCredit += selectScore[subjKeyR1].gsCredit11.Value;
                                            selectScore[subjKeyR1].sumScore += selectScore[subjKeyR1].gsScore11.Value;
                                            selectScore[subjKeyR1].sumScoreA += selectScore[subjKeyR1].gsScore11.Value * selectScore[subjKeyR1].gsCredit11.Value;

                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                chkHasTag1 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC1 += selectScore[subjKeyR1].gsScore11.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC1 += selectScore[subjKeyR1].gsScore11.Value * selectScore[subjKeyR1].gsCredit11.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC1++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC1 += selectScore[subjKeyR1].gsCredit11.Value;
                                            }

                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                chkHasTag2 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC2 += selectScore[subjKeyR1].gsScore11.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC2 += selectScore[subjKeyR1].gsScore11.Value * selectScore[subjKeyR1].gsCredit11.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC2++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC2 += selectScore[subjKeyR1].gsCredit11.Value;
                                            }
                                        }
                                        if (sumCredit12 > 0)
                                        {
                                            selectScore[subjKeyR1].gsScore12 = Math.Round(sumScore12 / sumCredit12, 0, MidpointRounding.AwayFromZero);
                                            selectScore[subjKeyR1].gsCredit12 = sumCredit12;
                                            selectScore[subjKeyR1].subjCount++;
                                            selectScore[subjKeyR1].sumCredit += selectScore[subjKeyR1].gsCredit12.Value;
                                            selectScore[subjKeyR1].sumScore += selectScore[subjKeyR1].gsScore12.Value;
                                            selectScore[subjKeyR1].sumScoreA += selectScore[subjKeyR1].gsScore12.Value * selectScore[subjKeyR1].gsCredit12.Value;

                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                chkHasTag1 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC1 += selectScore[subjKeyR1].gsScore12.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC1 += selectScore[subjKeyR1].gsScore12.Value * selectScore[subjKeyR1].gsCredit12.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC1++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC1 += selectScore[subjKeyR1].gsCredit12.Value;
                                            }

                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                chkHasTag2 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC2 += selectScore[subjKeyR1].gsScore12.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC2 += selectScore[subjKeyR1].gsScore12.Value * selectScore[subjKeyR1].gsCredit12.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC2++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC2 += selectScore[subjKeyR1].gsCredit12.Value;
                                            }

                                        }
                                        if (sumCredit21 > 0)
                                        {
                                            selectScore[subjKeyR1].gsScore21 = Math.Round(sumScore21 / sumCredit21, 0, MidpointRounding.AwayFromZero);
                                            selectScore[subjKeyR1].gsCredit21 = sumCredit21;
                                            selectScore[subjKeyR1].subjCount++;
                                            selectScore[subjKeyR1].sumCredit += selectScore[subjKeyR1].gsCredit21.Value;
                                            selectScore[subjKeyR1].sumScore += selectScore[subjKeyR1].gsScore21.Value;
                                            selectScore[subjKeyR1].sumScoreA += selectScore[subjKeyR1].gsScore21.Value * selectScore[subjKeyR1].gsCredit21.Value;

                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                chkHasTag1 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC1 += selectScore[subjKeyR1].gsScore21.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC1 += selectScore[subjKeyR1].gsScore21.Value * selectScore[subjKeyR1].gsCredit21.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC1++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC1 += selectScore[subjKeyR1].gsCredit21.Value;
                                            }

                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                chkHasTag2 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC2 += selectScore[subjKeyR1].gsScore21.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC2 += selectScore[subjKeyR1].gsScore21.Value * selectScore[subjKeyR1].gsCredit21.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC2++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC2 += selectScore[subjKeyR1].gsCredit21.Value;
                                            }
                                        }
                                        if (sumCredit22 > 0)
                                        {
                                            selectScore[subjKeyR1].gsScore22 = Math.Round(sumScore22 / sumCredit22, 0, MidpointRounding.AwayFromZero);
                                            selectScore[subjKeyR1].gsCredit22 = sumCredit22;
                                            selectScore[subjKeyR1].subjCount++;
                                            selectScore[subjKeyR1].sumCredit += selectScore[subjKeyR1].gsCredit22.Value;
                                            selectScore[subjKeyR1].sumScore += selectScore[subjKeyR1].gsScore22.Value;
                                            selectScore[subjKeyR1].sumScoreA += selectScore[subjKeyR1].gsScore22.Value * selectScore[subjKeyR1].gsCredit22.Value;

                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                chkHasTag1 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC1 += selectScore[subjKeyR1].gsScore22.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC1 += selectScore[subjKeyR1].gsScore22.Value * selectScore[subjKeyR1].gsCredit22.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC1++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC1 += selectScore[subjKeyR1].gsCredit22.Value;
                                            }

                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                chkHasTag2 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC2 += selectScore[subjKeyR1].gsScore22.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC2 += selectScore[subjKeyR1].gsScore22.Value * selectScore[subjKeyR1].gsCredit22.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC2++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC2 += selectScore[subjKeyR1].gsCredit22.Value;
                                            }

                                        }
                                        if (sumCredit31 > 0)
                                        {
                                            selectScore[subjKeyR1].gsScore31 = Math.Round(sumScore31 / sumCredit31, 0, MidpointRounding.AwayFromZero);
                                            selectScore[subjKeyR1].gsCredit31 = sumCredit31;
                                            selectScore[subjKeyR1].subjCount++;
                                            selectScore[subjKeyR1].sumCredit += selectScore[subjKeyR1].gsCredit31.Value;
                                            selectScore[subjKeyR1].sumScore += selectScore[subjKeyR1].gsScore31.Value;
                                            selectScore[subjKeyR1].sumScoreA += selectScore[subjKeyR1].gsScore31.Value * selectScore[subjKeyR1].gsCredit31.Value;

                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                chkHasTag1 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC1 += selectScore[subjKeyR1].gsScore31.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC1 += selectScore[subjKeyR1].gsScore31.Value * selectScore[subjKeyR1].gsCredit31.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC1++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC1 += selectScore[subjKeyR1].gsCredit31.Value;
                                            }

                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                chkHasTag2 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC2 += selectScore[subjKeyR1].gsScore31.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC2 += selectScore[subjKeyR1].gsScore31.Value * selectScore[subjKeyR1].gsCredit31.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC2++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC2 += selectScore[subjKeyR1].gsCredit31.Value;
                                            }

                                        }
                                        if (sumCredit32 > 0)
                                        {
                                            selectScore[subjKeyR1].gsScore32 = Math.Round(sumScore32 / sumCredit32, 0, MidpointRounding.AwayFromZero);
                                            selectScore[subjKeyR1].gsCredit32 = sumCredit32;
                                            selectScore[subjKeyR1].subjCount++;
                                            selectScore[subjKeyR1].sumCredit += selectScore[subjKeyR1].gsCredit32.Value;
                                            selectScore[subjKeyR1].sumScore += selectScore[subjKeyR1].gsScore32.Value;
                                            selectScore[subjKeyR1].sumScoreA += selectScore[subjKeyR1].gsScore32.Value * selectScore[subjKeyR1].gsCredit32.Value;

                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                chkHasTag1 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC1 += selectScore[subjKeyR1].gsScore32.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC1 += selectScore[subjKeyR1].gsScore32.Value * selectScore[subjKeyR1].gsCredit32.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC1++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC1 += selectScore[subjKeyR1].gsCredit32.Value;
                                            }

                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                chkHasTag2 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC2 += selectScore[subjKeyR1].gsScore32.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC2 += selectScore[subjKeyR1].gsScore32.Value * selectScore[subjKeyR1].gsCredit32.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC2++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC2 += selectScore[subjKeyR1].gsCredit32.Value;
                                            }

                                        }
                                        if (sumCredit41 > 0)
                                        {
                                            selectScore[subjKeyR1].gsScore41 = Math.Round(sumScore41 / sumCredit41, 0, MidpointRounding.AwayFromZero);
                                            selectScore[subjKeyR1].gsCredit41 = sumCredit41;
                                            selectScore[subjKeyR1].subjCount++;
                                            selectScore[subjKeyR1].sumCredit += selectScore[subjKeyR1].gsCredit41.Value;
                                            selectScore[subjKeyR1].sumScore += selectScore[subjKeyR1].gsScore41.Value;
                                            selectScore[subjKeyR1].sumScoreA += selectScore[subjKeyR1].gsScore41.Value * selectScore[subjKeyR1].gsCredit41.Value;

                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                chkHasTag1 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC1 += selectScore[subjKeyR1].gsScore41.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC1 += selectScore[subjKeyR1].gsScore41.Value * selectScore[subjKeyR1].gsCredit41.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC1++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC1 += selectScore[subjKeyR1].gsCredit41.Value;
                                            }

                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                chkHasTag2 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC2 += selectScore[subjKeyR1].gsScore41.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC2 += selectScore[subjKeyR1].gsScore41.Value * selectScore[subjKeyR1].gsCredit41.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC2++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC2 += selectScore[subjKeyR1].gsCredit41.Value;
                                            }

                                        }
                                        if (sumCredit42 > 0)
                                        {
                                            selectScore[subjKeyR1].gsScore42 = Math.Round(sumScore42 / sumCredit42, 0, MidpointRounding.AwayFromZero);
                                            selectScore[subjKeyR1].gsCredit42 = sumCredit42;
                                            selectScore[subjKeyR1].subjCount++;
                                            selectScore[subjKeyR1].sumCredit += selectScore[subjKeyR1].gsCredit42.Value;
                                            selectScore[subjKeyR1].sumScore += selectScore[subjKeyR1].gsScore42.Value;
                                            selectScore[subjKeyR1].sumScoreA += selectScore[subjKeyR1].gsScore42.Value * selectScore[subjKeyR1].gsCredit42.Value;

                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                chkHasTag1 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC1 += selectScore[subjKeyR1].gsScore42.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC1 += selectScore[subjKeyR1].gsScore42.Value * selectScore[subjKeyR1].gsCredit42.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC1++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC1 += selectScore[subjKeyR1].gsCredit42.Value;
                                            }

                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                chkHasTag2 = true;
                                                // 總分
                                                selectScore[subjKeyR1].sumScoreC2 += selectScore[subjKeyR1].gsScore42.Value; ;
                                                // 總分加權
                                                selectScore[subjKeyR1].sumScoreAC2 += selectScore[subjKeyR1].gsScore42.Value * selectScore[subjKeyR1].gsCredit42.Value;
                                                // 筆數
                                                selectScore[subjKeyR1].subjCountC2++;
                                                // 學分加總
                                                selectScore[subjKeyR1].sumCreditC2 += selectScore[subjKeyR1].gsCredit42.Value;
                                            }

                                        }

                                        // 計算平均,加權平均
                                        if (selectScore[subjKeyR1].subjCount > 0)
                                            selectScore[subjKeyR1].avgScore = selectScore[subjKeyR1].sumScore / selectScore[subjKeyR1].subjCount;

                                        if (selectScore[subjKeyR1].subjCountC1 > 0)
                                            selectScore[subjKeyR1].avgScoreC1 = selectScore[subjKeyR1].sumScoreC1 / selectScore[subjKeyR1].subjCountC1;

                                        if (selectScore[subjKeyR1].subjCountC2 > 0)
                                            selectScore[subjKeyR1].avgScoreC2 = selectScore[subjKeyR1].sumScoreC2 / selectScore[subjKeyR1].subjCountC2;

                                        if (selectScore[subjKeyR1].sumCredit > 0)
                                            selectScore[subjKeyR1].avgScoreA = selectScore[subjKeyR1].sumScoreA / selectScore[subjKeyR1].sumCredit;

                                        if (selectScore[subjKeyR1].sumCreditC1 > 0)
                                            selectScore[subjKeyR1].avgScoreAC1 = selectScore[subjKeyR1].sumScoreAC1 / selectScore[subjKeyR1].sumCreditC1;

                                        if (selectScore[subjKeyR1].sumCreditC2 > 0)
                                            selectScore[subjKeyR1].avgScoreAC2 = selectScore[subjKeyR1].sumScoreAC2 / selectScore[subjKeyR1].sumCreditC2;

                                    }
                                    #endregion


                                    #region 總分、平均、排名
                                    if (chkHasSubjectName)
                                    {
                                        // 平均
                                        if (selectScore[subjKeyR1].subjCount > 0)
                                            selectScore[subjKeyR1].avgScore = (decimal)(selectScore[subjKeyR1].sumScore / selectScore[subjKeyR1].subjCount);
                                        // 加權平均
                                        if (selectScore[subjKeyR1].sumCredit > 0)
                                            selectScore[subjKeyR1].avgScoreA = (decimal)(selectScore[subjKeyR1].sumScoreA / selectScore[subjKeyR1].sumCredit);

                                        // 平均(類別一)
                                        if (selectScore[subjKeyR1].subjCountC1 > 0)
                                            selectScore[subjKeyR1].avgScoreC1 = (decimal)(selectScore[subjKeyR1].sumScoreC1 / selectScore[subjKeyR1].subjCountC1);
                                        // 加權平均(類別一)
                                        if (selectScore[subjKeyR1].sumCreditC1 > 0)
                                            selectScore[subjKeyR1].avgScoreAC1 = (decimal)(selectScore[subjKeyR1].sumScoreAC1 / selectScore[subjKeyR1].sumCreditC1);

                                        // 平均(類別二)
                                        if (selectScore[subjKeyR1].subjCountC2 > 0)
                                            selectScore[subjKeyR1].avgScoreC2 = (decimal)(selectScore[subjKeyR1].sumScoreC2 / selectScore[subjKeyR1].subjCountC2);
                                        // 加權平均(類別二)
                                        if (selectScore[subjKeyR1].sumCreditC2 > 0)
                                            selectScore[subjKeyR1].avgScoreAC2 = (decimal)(selectScore[subjKeyR1].sumScoreAC2 / selectScore[subjKeyR1].sumCreditC2);

                                        if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                        {
                                            #region 各學期科目成績班排名
                                            {
                                                string key11 = "回歸一上科目班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key11)) ranks.Add(key11, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key11)) rankStudents.Add(key11, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore11.HasValue)
                                                {
                                                    ranks[key11].Add(selectScore[subjKeyR1].gsScore11.Value);
                                                    rankStudents[key11].Add(studentID);
                                                }

                                                string key12 = "回歸一下科目班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key12)) ranks.Add(key12, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key12)) rankStudents.Add(key12, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore12.HasValue)
                                                {
                                                    ranks[key12].Add(selectScore[subjKeyR1].gsScore12.Value);
                                                    rankStudents[key12].Add(studentID);
                                                }

                                                string key21 = "回歸二上科目班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key21)) ranks.Add(key21, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key21)) rankStudents.Add(key21, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore21.HasValue)
                                                {
                                                    ranks[key21].Add(selectScore[subjKeyR1].gsScore21.Value);
                                                    rankStudents[key21].Add(studentID);
                                                }

                                                string key22 = "回歸二下科目班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key22)) ranks.Add(key22, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key22)) rankStudents.Add(key22, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore22.HasValue)
                                                {
                                                    ranks[key22].Add(selectScore[subjKeyR1].gsScore22.Value);
                                                    rankStudents[key22].Add(studentID);
                                                }

                                                string key31 = "回歸三上科目班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key31)) ranks.Add(key31, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key31)) rankStudents.Add(key31, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore31.HasValue)
                                                {
                                                    ranks[key31].Add(selectScore[subjKeyR1].gsScore31.Value);
                                                    rankStudents[key31].Add(studentID);
                                                }

                                                string key32 = "回歸三下科目班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key32)) ranks.Add(key32, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key32)) rankStudents.Add(key32, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore32.HasValue)
                                                {
                                                    ranks[key32].Add(selectScore[subjKeyR1].gsScore32.Value);
                                                    rankStudents[key32].Add(studentID);
                                                }

                                                string key41 = "回歸四上科目班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key41)) ranks.Add(key41, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key41)) rankStudents.Add(key41, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore41.HasValue)
                                                {
                                                    ranks[key41].Add(selectScore[subjKeyR1].gsScore41.Value);
                                                    rankStudents[key41].Add(studentID);
                                                }

                                                string key42 = "回歸四下科目班排名" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key42)) ranks.Add(key42, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key42)) rankStudents.Add(key42, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore42.HasValue)
                                                {
                                                    ranks[key42].Add(selectScore[subjKeyR1].gsScore42.Value);
                                                    rankStudents[key42].Add(studentID);
                                                }
                                            }

                                            #endregion

                                            #region 各學期科目成績科科排名
                                            {
                                                if (studentRec.Department != "")
                                                {
                                                    string key11 = "回歸一上科目科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key11)) ranks.Add(key11, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key11)) rankStudents.Add(key11, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore11.HasValue)
                                                    {
                                                        ranks[key11].Add(selectScore[subjKeyR1].gsScore11.Value);
                                                        rankStudents[key11].Add(studentID);
                                                    }

                                                    string key12 = "回歸一下科目科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key12)) ranks.Add(key12, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key12)) rankStudents.Add(key12, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore12.HasValue)
                                                    {
                                                        ranks[key12].Add(selectScore[subjKeyR1].gsScore12.Value);
                                                        rankStudents[key12].Add(studentID);
                                                    }

                                                    string key21 = "回歸二上科目科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key21)) ranks.Add(key21, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key21)) rankStudents.Add(key21, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore21.HasValue)
                                                    {
                                                        ranks[key21].Add(selectScore[subjKeyR1].gsScore21.Value);
                                                        rankStudents[key21].Add(studentID);
                                                    }

                                                    string key22 = "回歸二下科目科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key22)) ranks.Add(key22, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key22)) rankStudents.Add(key22, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore22.HasValue)
                                                    {
                                                        ranks[key22].Add(selectScore[subjKeyR1].gsScore22.Value);
                                                        rankStudents[key22].Add(studentID);
                                                    }

                                                    string key31 = "回歸三上科目科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key31)) ranks.Add(key31, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key31)) rankStudents.Add(key31, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore31.HasValue)
                                                    {
                                                        ranks[key31].Add(selectScore[subjKeyR1].gsScore31.Value);
                                                        rankStudents[key31].Add(studentID);
                                                    }

                                                    string key32 = "回歸三下科目科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key32)) ranks.Add(key32, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key32)) rankStudents.Add(key32, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore32.HasValue)
                                                    {
                                                        ranks[key32].Add(selectScore[subjKeyR1].gsScore32.Value);
                                                        rankStudents[key32].Add(studentID);
                                                    }

                                                    string key41 = "回歸四上科目科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key41)) ranks.Add(key41, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key41)) rankStudents.Add(key41, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore41.HasValue)
                                                    {
                                                        ranks[key41].Add(selectScore[subjKeyR1].gsScore41.Value);
                                                        rankStudents[key41].Add(studentID);
                                                    }

                                                    string key42 = "回歸四下科目科排名" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key42)) ranks.Add(key42, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key42)) rankStudents.Add(key42, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore42.HasValue)
                                                    {
                                                        ranks[key42].Add(selectScore[subjKeyR1].gsScore42.Value);
                                                        rankStudents[key42].Add(studentID);
                                                    }

                                                }
                                            }
                                            #endregion

                                            #region 各學期科目成績科校排名
                                            {
                                                string key11 = "回歸一上科目校排名" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key11)) ranks.Add(key11, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key11)) rankStudents.Add(key11, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore11.HasValue)
                                                {
                                                    ranks[key11].Add(selectScore[subjKeyR1].gsScore11.Value);
                                                    rankStudents[key11].Add(studentID);
                                                }

                                                string key12 = "回歸一下科目校排名" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key12)) ranks.Add(key12, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key12)) rankStudents.Add(key12, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore12.HasValue)
                                                {
                                                    ranks[key12].Add(selectScore[subjKeyR1].gsScore12.Value);
                                                    rankStudents[key12].Add(studentID);
                                                }

                                                string key21 = "回歸二上科目校排名" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key21)) ranks.Add(key21, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key21)) rankStudents.Add(key21, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore21.HasValue)
                                                {
                                                    ranks[key21].Add(selectScore[subjKeyR1].gsScore21.Value);
                                                    rankStudents[key21].Add(studentID);
                                                }

                                                string key22 = "回歸二下科目校排名" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key22)) ranks.Add(key22, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key22)) rankStudents.Add(key22, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore22.HasValue)
                                                {
                                                    ranks[key22].Add(selectScore[subjKeyR1].gsScore22.Value);
                                                    rankStudents[key22].Add(studentID);
                                                }

                                                string key31 = "回歸三上科目校排名" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key31)) ranks.Add(key31, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key31)) rankStudents.Add(key31, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore31.HasValue)
                                                {
                                                    ranks[key31].Add(selectScore[subjKeyR1].gsScore31.Value);
                                                    rankStudents[key31].Add(studentID);
                                                }

                                                string key32 = "回歸三下科目校排名" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key32)) ranks.Add(key32, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key32)) rankStudents.Add(key32, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore32.HasValue)
                                                {
                                                    ranks[key32].Add(selectScore[subjKeyR1].gsScore32.Value);
                                                    rankStudents[key32].Add(studentID);
                                                }

                                                string key41 = "回歸四上科目校排名" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key41)) ranks.Add(key41, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key41)) rankStudents.Add(key41, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore41.HasValue)
                                                {
                                                    ranks[key41].Add(selectScore[subjKeyR1].gsScore41.Value);
                                                    rankStudents[key41].Add(studentID);
                                                }

                                                string key42 = "回歸四下科目校排名" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key42)) ranks.Add(key42, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key42)) rankStudents.Add(key42, new List<string>());
                                                if (selectScore[subjKeyR1].gsScore42.HasValue)
                                                {
                                                    ranks[key42].Add(selectScore[subjKeyR1].gsScore42.Value);
                                                    rankStudents[key42].Add(studentID);
                                                }
                                            }
                                            #endregion

                                            #region 各學期科目成績科類1排名
                                            {
                                                if (studentRec.Fields.ContainsKey("tag1"))
                                                {
                                                    string key11 = "回歸一上科目類1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key11)) ranks.Add(key11, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key11)) rankStudents.Add(key11, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore11.HasValue)
                                                    {
                                                        ranks[key11].Add(selectScore[subjKeyR1].gsScore11.Value);
                                                        rankStudents[key11].Add(studentID);
                                                    }

                                                    string key12 = "回歸一下科目類1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key12)) ranks.Add(key12, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key12)) rankStudents.Add(key12, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore12.HasValue)
                                                    {
                                                        ranks[key12].Add(selectScore[subjKeyR1].gsScore12.Value);
                                                        rankStudents[key12].Add(studentID);
                                                    }

                                                    string key21 = "回歸二上科目類1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key21)) ranks.Add(key21, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key21)) rankStudents.Add(key21, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore21.HasValue)
                                                    {
                                                        ranks[key21].Add(selectScore[subjKeyR1].gsScore21.Value);
                                                        rankStudents[key21].Add(studentID);
                                                    }

                                                    string key22 = "回歸二下科目類1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key22)) ranks.Add(key22, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key22)) rankStudents.Add(key22, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore22.HasValue)
                                                    {
                                                        ranks[key22].Add(selectScore[subjKeyR1].gsScore22.Value);
                                                        rankStudents[key22].Add(studentID);
                                                    }

                                                    string key31 = "回歸三上科目類1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key31)) ranks.Add(key31, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key31)) rankStudents.Add(key31, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore31.HasValue)
                                                    {
                                                        ranks[key31].Add(selectScore[subjKeyR1].gsScore31.Value);
                                                        rankStudents[key31].Add(studentID);
                                                    }

                                                    string key32 = "回歸三下科目類1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key32)) ranks.Add(key32, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key32)) rankStudents.Add(key32, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore32.HasValue)
                                                    {
                                                        ranks[key32].Add(selectScore[subjKeyR1].gsScore32.Value);
                                                        rankStudents[key32].Add(studentID);
                                                    }

                                                    string key41 = "回歸四上科目類1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key41)) ranks.Add(key41, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key41)) rankStudents.Add(key41, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore41.HasValue)
                                                    {
                                                        ranks[key41].Add(selectScore[subjKeyR1].gsScore41.Value);
                                                        rankStudents[key41].Add(studentID);
                                                    }

                                                    string key42 = "回歸四下科目類1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key42)) ranks.Add(key42, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key42)) rankStudents.Add(key42, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore42.HasValue)
                                                    {
                                                        ranks[key42].Add(selectScore[subjKeyR1].gsScore42.Value);
                                                        rankStudents[key42].Add(studentID);
                                                    }
                                                }
                                            }
                                            #endregion

                                            #region 各學期科目成績科類2排名
                                            {
                                                if (studentRec.Fields.ContainsKey("tag2"))
                                                {
                                                    string key11 = "回歸一上科目類2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key11)) ranks.Add(key11, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key11)) rankStudents.Add(key11, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore11.HasValue)
                                                    {
                                                        ranks[key11].Add(selectScore[subjKeyR1].gsScore11.Value);
                                                        rankStudents[key11].Add(studentID);
                                                    }

                                                    string key12 = "回歸一下科目類2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key12)) ranks.Add(key12, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key12)) rankStudents.Add(key12, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore12.HasValue)
                                                    {
                                                        ranks[key12].Add(selectScore[subjKeyR1].gsScore12.Value);
                                                        rankStudents[key12].Add(studentID);
                                                    }

                                                    string key21 = "回歸二上科目類2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key21)) ranks.Add(key21, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key21)) rankStudents.Add(key21, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore21.HasValue)
                                                    {
                                                        ranks[key21].Add(selectScore[subjKeyR1].gsScore21.Value);
                                                        rankStudents[key21].Add(studentID);
                                                    }

                                                    string key22 = "回歸二下科目類2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key22)) ranks.Add(key22, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key22)) rankStudents.Add(key22, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore22.HasValue)
                                                    {
                                                        ranks[key22].Add(selectScore[subjKeyR1].gsScore22.Value);
                                                        rankStudents[key22].Add(studentID);
                                                    }

                                                    string key31 = "回歸三上科目類2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key31)) ranks.Add(key31, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key31)) rankStudents.Add(key31, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore31.HasValue)
                                                    {
                                                        ranks[key31].Add(selectScore[subjKeyR1].gsScore31.Value);
                                                        rankStudents[key31].Add(studentID);
                                                    }

                                                    string key32 = "回歸三下科目類2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key32)) ranks.Add(key32, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key32)) rankStudents.Add(key32, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore32.HasValue)
                                                    {
                                                        ranks[key32].Add(selectScore[subjKeyR1].gsScore32.Value);
                                                        rankStudents[key32].Add(studentID);
                                                    }

                                                    string key41 = "回歸四上科目類2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key41)) ranks.Add(key41, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key41)) rankStudents.Add(key41, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore41.HasValue)
                                                    {
                                                        ranks[key41].Add(selectScore[subjKeyR1].gsScore41.Value);
                                                        rankStudents[key41].Add(studentID);
                                                    }

                                                    string key42 = "回歸四下科目類2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key42)) ranks.Add(key42, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key42)) rankStudents.Add(key42, new List<string>());
                                                    if (selectScore[subjKeyR1].gsScore42.HasValue)
                                                    {
                                                        ranks[key42].Add(selectScore[subjKeyR1].gsScore42.Value);
                                                        rankStudents[key42].Add(studentID);
                                                    }
                                                }
                                            }
                                            #endregion


                                            #region 班排名_回歸_
                                            {
                                                string key1 = "總分班排名_回歸_" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[subjKeyR1].sumScore);
                                                rankStudents[key1].Add(studentID);

                                                string key2 = "加權總分班排名_回歸_" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[subjKeyR1].sumScoreA);
                                                rankStudents[key2].Add(studentID);

                                                string key3 = "平均班排名_回歸_" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                                ranks[key3].Add(selectScore[subjKeyR1].avgScore);
                                                rankStudents[key3].Add(studentID);


                                                string key4 = "加權平均班排名_回歸_" + studentRec.RefClass.ClassID + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                                ranks[key4].Add(selectScore[subjKeyR1].avgScoreA);
                                                rankStudents[key4].Add(studentID);

                                            }
                                            #endregion
                                            #region 科排名_回歸_
                                            {
                                                if (studentRec.Department != "")
                                                {
                                                    //各科目科排名_回歸_
                                                    string key1 = "總分科排名_回歸_" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                    ranks[key1].Add(selectScore[subjKeyR1].sumScore);
                                                    rankStudents[key1].Add(studentID);

                                                    //各科目科排名_回歸_
                                                    string key2 = "加權總分科排名_回歸_" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                    ranks[key2].Add(selectScore[subjKeyR1].sumScoreA);
                                                    rankStudents[key2].Add(studentID);

                                                    //各科目科排名_回歸_
                                                    string key3 = "平均科排名_回歸_" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                                    ranks[key3].Add(selectScore[subjKeyR1].avgScore);
                                                    rankStudents[key3].Add(studentID);

                                                    //各科目科排名_回歸_
                                                    string key4 = "加權平均科排名_回歸_" + studentRec.Department + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                                    ranks[key4].Add(selectScore[subjKeyR1].avgScoreA);
                                                    rankStudents[key4].Add(studentID);

                                                }
                                            }
                                            #endregion
                                            #region 全校排名_回歸_
                                            {
                                                string key1 = "總分全校排名_回歸_" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[subjKeyR1].sumScore);
                                                rankStudents[key1].Add(studentID);

                                                string key2 = "加權總分全校排名_回歸_" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[subjKeyR1].sumScoreA);
                                                rankStudents[key2].Add(studentID);

                                                string key3 = "平均全校排名_回歸_" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                                ranks[key3].Add(selectScore[subjKeyR1].avgScore);
                                                rankStudents[key3].Add(studentID);

                                                string key4 = "加權平均全校排名_回歸_" + gradeyear + "^^^" + SubjName;
                                                if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                                ranks[key4].Add(selectScore[subjKeyR1].avgScoreA);
                                                rankStudents[key4].Add(studentID);

                                            }
                                            #endregion
                                            #region 類別1排名_回歸_
                                            if (studentRec.Fields.ContainsKey("tag1"))
                                            {
                                                // 假如上面有處理過類別1的科目
                                                if (chkHasTag1 == true)
                                                {
                                                    string key1 = "總分類別1排名_回歸_" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                    ranks[key1].Add(selectScore[subjKeyR1].sumScoreC1);
                                                    rankStudents[key1].Add(studentID);

                                                    string key2 = "加權總分類別1排名_回歸_" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                    ranks[key2].Add(selectScore[subjKeyR1].sumScoreAC1);
                                                    rankStudents[key2].Add(studentID);

                                                    string key3 = "平均類別1排名_回歸_" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                                    ranks[key3].Add(selectScore[subjKeyR1].avgScoreC1);
                                                    rankStudents[key3].Add(studentID);

                                                    string key4 = "加權平均類別1排名_回歸_" + studentRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                                    ranks[key4].Add(selectScore[subjKeyR1].avgScoreAC1);
                                                    rankStudents[key4].Add(studentID);
                                                }
                                            }
                                            #endregion
                                            #region 類別2排名_回歸_
                                            if (studentRec.Fields.ContainsKey("tag2"))
                                            {
                                                // 假如上面有處理過類別2的科目
                                                if (chkHasTag2 == true)
                                                {
                                                    string key1 = "總分類別2排名_回歸_" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                    ranks[key1].Add(selectScore[subjKeyR1].sumScoreC2);
                                                    rankStudents[key1].Add(studentID);

                                                    string key2 = "加權總分類別2排名_回歸_" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                    ranks[key2].Add(selectScore[subjKeyR1].sumScoreAC2);
                                                    rankStudents[key2].Add(studentID);

                                                    string key3 = "平均類別2排名_回歸_" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                                    ranks[key3].Add(selectScore[subjKeyR1].avgScoreC2);
                                                    rankStudents[key3].Add(studentID);

                                                    string key4 = "加權平均類別2排名_回歸_" + studentRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + SubjName;
                                                    if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                                    if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                                    ranks[key4].Add(selectScore[subjKeyR1].avgScoreAC2);
                                                    rankStudents[key4].Add(studentID);
                                                }
                                            }
                                            #endregion
                                        }
                                    }

                                    #endregion

                                }
                                #endregion                                
                            }
                            catch (Exception ex)
                            { }
                            #endregion


                            if (setting.計算學業成績排名)
                            {
                                #region 處理學業成績
                                // 只處理年級學期，不處理科目                                     
                                string SemsKey = studentID + "學業";
                                foreach (var semesterEntry in studentRec.SemesterEntryScoreList)
                                {
                                    string gs = semesterEntry.GradeYear + "" + semesterEntry.Semester;
                                    // 不在勾選年級學期內跳過
                                    if (!setting.useGradeSemesterList.Contains(gs))
                                        continue;

                                    if (semesterEntry.Entry == "學業")
                                    {

                                        if (!selectScore.ContainsKey(SemsKey))
                                            selectScore.Add(SemsKey, new studScore());


                                        // 處理年級學期的學期成績
                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey].gsScore11 = semesterEntry.Score;
                                            selectScore[SemsKey].gsSchoolYear11 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey].gsScore12 = semesterEntry.Score;
                                            selectScore[SemsKey].gsSchoolYear12 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey].gsScore21 = semesterEntry.Score;
                                            selectScore[SemsKey].gsSchoolYear21 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey].gsScore22 = semesterEntry.Score;
                                            selectScore[SemsKey].gsSchoolYear22 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey].gsScore31 = semesterEntry.Score;
                                            selectScore[SemsKey].gsSchoolYear31 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey].gsScore32 = semesterEntry.Score;
                                            selectScore[SemsKey].gsSchoolYear32 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey].gsScore41 = semesterEntry.Score;
                                            selectScore[SemsKey].gsSchoolYear41 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey].gsScore42 = semesterEntry.Score;
                                            selectScore[SemsKey].gsSchoolYear42 = semesterEntry.SchoolYear;
                                        }
                                        // 總分
                                        selectScore[SemsKey].sumScore += semesterEntry.Score;
                                        selectScore[SemsKey].subjCount++;

                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            // 總分
                                            selectScore[SemsKey].sumScoreC1 += semesterEntry.Score;
                                            selectScore[SemsKey].subjCountC1++;

                                        }

                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            // 總分
                                            selectScore[SemsKey].sumScoreC2 += semesterEntry.Score;
                                            selectScore[SemsKey].subjCountC2++;

                                        }
                                    }
                                }
                                // 計算學業平均
                                if (selectScore.ContainsKey(SemsKey))
                                {
                                    if (selectScore[SemsKey].subjCount > 0)
                                        selectScore[SemsKey].avgScore = (decimal)(selectScore[SemsKey].sumScore / selectScore[SemsKey].subjCount);

                                    if (selectScore[SemsKey].subjCountC1 > 0)
                                        selectScore[SemsKey].avgScoreC1 = (decimal)(selectScore[SemsKey].sumScoreC1 / selectScore[SemsKey].subjCountC1);

                                    if (selectScore[SemsKey].subjCountC2 > 0)
                                        selectScore[SemsKey].avgScoreC2 = (decimal)(selectScore[SemsKey].sumScoreC2 / selectScore[SemsKey].subjCountC2);


                                    if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                    {
                                        #region 班排名
                                        {
                                            string key1 = "學業成績總分班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業成績平均班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 科排名
                                        {
                                            if (studentRec.Department != "")
                                            {
                                                string key1 = "學業成績總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[SemsKey].sumScore);
                                                rankStudents[key1].Add(studentID);

                                                string key2 = "學業成績平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[SemsKey].avgScore);
                                                rankStudents[key2].Add(studentID);
                                            }
                                        }
                                        #endregion
                                        #region 全校排名
                                        {
                                            string key1 = "學業成績總分全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業成績平均全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 類別1排名
                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            string key1 = "學業成績總分類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey].sumScoreC1);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業成績平均類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey].avgScoreC1);
                                            rankStudents[key2].Add(studentID);
                                        }
                                        #endregion
                                        #region 類別2排名
                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            string key1 = "學業成績總分類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey].sumScoreC2);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業成績平均類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey].avgScoreC2);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                    }
                                }
                                #endregion 處理學業成績
                                // 小郭, 2014/01/02
                                #region 處理學業原始成績
                                // 只處理年級學期，不處理科目                                     
                                string SemsKey_5 = studentID + "學業原始";
                                foreach (var semesterEntry in studentRec.SemesterEntryScoreList)
                                {
                                    string gs = semesterEntry.GradeYear + "" + semesterEntry.Semester;
                                    // 不在勾選年級學期內跳過
                                    if (!setting.useGradeSemesterList.Contains(gs))
                                        continue;

                                    if (semesterEntry.Entry == "學業(原始)")
                                    {

                                        if (!selectScore.ContainsKey(SemsKey_5))
                                            selectScore.Add(SemsKey_5, new studScore());


                                        // 處理年級學期的學期成績
                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_5].gsScore11 = semesterEntry.Score;
                                            selectScore[SemsKey_5].gsSchoolYear11 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_5].gsScore12 = semesterEntry.Score;
                                            selectScore[SemsKey_5].gsSchoolYear12 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_5].gsScore21 = semesterEntry.Score;
                                            selectScore[SemsKey_5].gsSchoolYear21 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_5].gsScore22 = semesterEntry.Score;
                                            selectScore[SemsKey_5].gsSchoolYear22 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_5].gsScore31 = semesterEntry.Score;
                                            selectScore[SemsKey_5].gsSchoolYear31 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_5].gsScore32 = semesterEntry.Score;
                                            selectScore[SemsKey_5].gsSchoolYear32 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_5].gsScore41 = semesterEntry.Score;
                                            selectScore[SemsKey_5].gsSchoolYear41 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_5].gsScore42 = semesterEntry.Score;
                                            selectScore[SemsKey_5].gsSchoolYear42 = semesterEntry.SchoolYear;
                                        }
                                        // 總分
                                        selectScore[SemsKey_5].sumScore += semesterEntry.Score;
                                        selectScore[SemsKey_5].subjCount++;

                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            // 總分
                                            selectScore[SemsKey_5].sumScoreC1 += semesterEntry.Score;
                                            selectScore[SemsKey_5].subjCountC1++;

                                        }

                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            // 總分
                                            selectScore[SemsKey_5].sumScoreC2 += semesterEntry.Score;
                                            selectScore[SemsKey_5].subjCountC2++;

                                        }
                                    }
                                }
                                // 計算學業平均
                                if (selectScore.ContainsKey(SemsKey_5))
                                {
                                    if (selectScore[SemsKey_5].subjCount > 0)
                                        selectScore[SemsKey_5].avgScore = (decimal)(selectScore[SemsKey_5].sumScore / selectScore[SemsKey_5].subjCount);

                                    if (selectScore[SemsKey_5].subjCountC1 > 0)
                                        selectScore[SemsKey_5].avgScoreC1 = (decimal)(selectScore[SemsKey_5].sumScoreC1 / selectScore[SemsKey_5].subjCountC1);

                                    if (selectScore[SemsKey_5].subjCountC2 > 0)
                                        selectScore[SemsKey_5].avgScoreC2 = (decimal)(selectScore[SemsKey_5].sumScoreC2 / selectScore[SemsKey_5].subjCountC2);


                                    if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                    {
                                        #region 班排名
                                        {
                                            string key1 = "學業原始成績總分班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_5].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業原始成績平均班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_5].avgScore);
                                            rankStudents[key2].Add(studentID);
                                        }
                                        #endregion
                                        #region 科排名
                                        {
                                            if (studentRec.Department != "")
                                            {
                                                string key1 = "學業原始成績總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[SemsKey_5].sumScore);
                                                rankStudents[key1].Add(studentID);

                                                string key2 = "學業原始成績平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[SemsKey_5].avgScore);
                                                rankStudents[key2].Add(studentID);
                                            }
                                        }
                                        #endregion
                                        #region 全校排名
                                        {
                                            string key1 = "學業原始成績總分全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_5].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業原始成績平均全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_5].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 類別1排名
                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            string key1 = "學業原始成績總分類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_5].sumScoreC1);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業原始成績平均類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_5].avgScoreC1);
                                            rankStudents[key2].Add(studentID);
                                        }
                                        #endregion
                                        #region 類別2排名
                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            string key1 = "學業原始成績總分類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_5].sumScoreC2);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業原始成績平均類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_5].avgScoreC2);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                    }
                                }
                                #endregion 處理學業原始成績

                                #region 處理體育
                                // 只處理年級學期，不處理科目                                     
                                string SemsKey_1 = studentID + "學業體育";
                                foreach (var semesterEntry in studentRec.SemesterEntryScoreList)
                                {
                                    string gs = semesterEntry.GradeYear + "" + semesterEntry.Semester;
                                    // 不在勾選年級學期內跳過
                                    if (!setting.useGradeSemesterList.Contains(gs))
                                        continue;

                                    if (semesterEntry.Entry == "體育")
                                    {

                                        if (!selectScore.ContainsKey(SemsKey_1))
                                            selectScore.Add(SemsKey_1, new studScore());


                                        // 處理年級學期的學期成績
                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_1].gsScore11 = semesterEntry.Score;
                                            selectScore[SemsKey_1].gsSchoolYear11 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_1].gsScore12 = semesterEntry.Score;
                                            selectScore[SemsKey_1].gsSchoolYear12 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_1].gsScore21 = semesterEntry.Score;
                                            selectScore[SemsKey_1].gsSchoolYear21 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_1].gsScore22 = semesterEntry.Score;
                                            selectScore[SemsKey_1].gsSchoolYear22 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_1].gsScore31 = semesterEntry.Score;
                                            selectScore[SemsKey_1].gsSchoolYear31 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_1].gsScore32 = semesterEntry.Score;
                                            selectScore[SemsKey_1].gsSchoolYear32 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_1].gsScore41 = semesterEntry.Score;
                                            selectScore[SemsKey_1].gsSchoolYear41 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_1].gsScore42 = semesterEntry.Score;
                                            selectScore[SemsKey_1].gsSchoolYear42 = semesterEntry.SchoolYear;
                                        }
                                        // 總分
                                        selectScore[SemsKey_1].sumScore += semesterEntry.Score;
                                        selectScore[SemsKey_1].subjCount++;

                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            // 總分
                                            selectScore[SemsKey_1].sumScoreC1 += semesterEntry.Score;
                                            selectScore[SemsKey_1].subjCountC1++;

                                        }

                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            // 總分
                                            selectScore[SemsKey_1].sumScoreC2 += semesterEntry.Score;
                                            selectScore[SemsKey_1].subjCountC2++;

                                        }
                                    }
                                }
                                // 計算學業體育平均
                                if (selectScore.ContainsKey(SemsKey_1))
                                {
                                    if (selectScore[SemsKey_1].subjCount > 0)
                                        selectScore[SemsKey_1].avgScore = (decimal)(selectScore[SemsKey_1].sumScore / selectScore[SemsKey_1].subjCount);

                                    if (selectScore[SemsKey_1].subjCountC1 > 0)
                                        selectScore[SemsKey_1].avgScoreC1 = (decimal)(selectScore[SemsKey_1].sumScoreC1 / selectScore[SemsKey_1].subjCountC1);

                                    if (selectScore[SemsKey_1].subjCountC2 > 0)
                                        selectScore[SemsKey_1].avgScoreC2 = (decimal)(selectScore[SemsKey_1].sumScoreC2 / selectScore[SemsKey_1].subjCountC2);


                                    if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                    {
                                        #region 班排名
                                        {
                                            string key1 = "學業體育成績總分班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_1].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業體育成績平均班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_1].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 科排名
                                        {
                                            if (studentRec.Department != "")
                                            {
                                                string key1 = "學業體育成績總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[SemsKey_1].sumScore);
                                                rankStudents[key1].Add(studentID);

                                                string key2 = "學業體育成績平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[SemsKey_1].avgScore);
                                                rankStudents[key2].Add(studentID);
                                            }
                                        }
                                        #endregion
                                        #region 全校排名
                                        {
                                            string key1 = "學業體育成績總分全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_1].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業體育成績平均全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_1].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 類別1排名
                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            string key1 = "學業體育成績總分類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_1].sumScoreC1);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業體育成績平均類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_1].avgScoreC1);
                                            rankStudents[key2].Add(studentID);
                                        }
                                        #endregion
                                        #region 類別2排名
                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            string key1 = "學業體育成績總分類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_1].sumScoreC2);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業體育成績平均類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_1].avgScoreC2);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                    }
                                }
                                #endregion 處理體育

                                #region 處理健康與護理
                                // 只處理年級學期，不處理科目                                     
                                string SemsKey_2 = studentID + "學業健康與護理";
                                foreach (var semesterEntry in studentRec.SemesterEntryScoreList)
                                {
                                    string gs = semesterEntry.GradeYear + "" + semesterEntry.Semester;
                                    // 不在勾選年級學期內跳過
                                    if (!setting.useGradeSemesterList.Contains(gs))
                                        continue;

                                    if (semesterEntry.Entry == "健康與護理")
                                    {

                                        if (!selectScore.ContainsKey(SemsKey_2))
                                            selectScore.Add(SemsKey_2, new studScore());


                                        // 處理年級學期的學期成績
                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_2].gsScore11 = semesterEntry.Score;
                                            selectScore[SemsKey_2].gsSchoolYear11 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_2].gsScore12 = semesterEntry.Score;
                                            selectScore[SemsKey_2].gsSchoolYear12 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_2].gsScore21 = semesterEntry.Score;
                                            selectScore[SemsKey_2].gsSchoolYear21 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_2].gsScore22 = semesterEntry.Score;
                                            selectScore[SemsKey_2].gsSchoolYear22 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_2].gsScore31 = semesterEntry.Score;
                                            selectScore[SemsKey_2].gsSchoolYear31 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_2].gsScore32 = semesterEntry.Score;
                                            selectScore[SemsKey_2].gsSchoolYear32 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_2].gsScore41 = semesterEntry.Score;
                                            selectScore[SemsKey_2].gsSchoolYear41 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_2].gsScore42 = semesterEntry.Score;
                                            selectScore[SemsKey_2].gsSchoolYear42 = semesterEntry.SchoolYear;
                                        }
                                        // 總分
                                        selectScore[SemsKey_2].sumScore += semesterEntry.Score;
                                        selectScore[SemsKey_2].subjCount++;

                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            // 總分
                                            selectScore[SemsKey_2].sumScoreC1 += semesterEntry.Score;
                                            selectScore[SemsKey_2].subjCountC1++;

                                        }

                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            // 總分
                                            selectScore[SemsKey_2].sumScoreC2 += semesterEntry.Score;
                                            selectScore[SemsKey_2].subjCountC2++;

                                        }
                                    }
                                }
                                // 計算學業健康與護理平均
                                if (selectScore.ContainsKey(SemsKey_2))
                                {
                                    if (selectScore[SemsKey_2].subjCount > 0)
                                        selectScore[SemsKey_2].avgScore = (decimal)(selectScore[SemsKey_2].sumScore / selectScore[SemsKey_2].subjCount);

                                    if (selectScore[SemsKey_2].subjCountC1 > 0)
                                        selectScore[SemsKey_2].avgScoreC1 = (decimal)(selectScore[SemsKey_2].sumScoreC1 / selectScore[SemsKey_2].subjCountC1);

                                    if (selectScore[SemsKey_2].subjCountC2 > 0)
                                        selectScore[SemsKey_2].avgScoreC2 = (decimal)(selectScore[SemsKey_2].sumScoreC2 / selectScore[SemsKey_2].subjCountC2);


                                    if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                    {
                                        #region 班排名
                                        {
                                            string key1 = "學業健康與護理成績總分班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_2].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業健康與護理成績平均班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_2].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 科排名
                                        {
                                            if (studentRec.Department != "")
                                            {
                                                string key1 = "學業健康與護理成績總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[SemsKey_2].sumScore);
                                                rankStudents[key1].Add(studentID);

                                                string key2 = "學業健康與護理成績平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[SemsKey_2].avgScore);
                                                rankStudents[key2].Add(studentID);
                                            }
                                        }
                                        #endregion
                                        #region 全校排名
                                        {
                                            string key1 = "學業健康與護理成績總分全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_2].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業健康與護理成績平均全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_2].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 類別1排名
                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            string key1 = "學業健康與護理成績總分類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_2].sumScoreC1);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業健康與護理成績平均類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_2].avgScoreC1);
                                            rankStudents[key2].Add(studentID);
                                        }
                                        #endregion
                                        #region 類別2排名
                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            string key1 = "學業健康與護理成績總分類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_2].sumScoreC2);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業健康與護理成績平均類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_2].avgScoreC2);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                    }
                                }
                                #endregion 處理健康與護理

                                #region 處理國防通識
                                // 只處理年級學期，不處理科目                                     
                                string SemsKey_3 = studentID + "學業國防通識";
                                foreach (var semesterEntry in studentRec.SemesterEntryScoreList)
                                {
                                    string gs = semesterEntry.GradeYear + "" + semesterEntry.Semester;
                                    // 不在勾選年級學期內跳過
                                    if (!setting.useGradeSemesterList.Contains(gs))
                                        continue;

                                    if (semesterEntry.Entry == "國防通識")
                                    {

                                        if (!selectScore.ContainsKey(SemsKey_3))
                                            selectScore.Add(SemsKey_3, new studScore());


                                        // 處理年級學期的學期成績
                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_3].gsScore11 = semesterEntry.Score;
                                            selectScore[SemsKey_3].gsSchoolYear11 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_3].gsScore12 = semesterEntry.Score;
                                            selectScore[SemsKey_3].gsSchoolYear12 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_3].gsScore21 = semesterEntry.Score;
                                            selectScore[SemsKey_3].gsSchoolYear21 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_3].gsScore22 = semesterEntry.Score;
                                            selectScore[SemsKey_3].gsSchoolYear22 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_3].gsScore31 = semesterEntry.Score;
                                            selectScore[SemsKey_3].gsSchoolYear31 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_3].gsScore32 = semesterEntry.Score;
                                            selectScore[SemsKey_3].gsSchoolYear32 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_3].gsScore41 = semesterEntry.Score;
                                            selectScore[SemsKey_3].gsSchoolYear41 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_3].gsScore42 = semesterEntry.Score;
                                            selectScore[SemsKey_3].gsSchoolYear42 = semesterEntry.SchoolYear;
                                        }
                                        // 總分
                                        selectScore[SemsKey_3].sumScore += semesterEntry.Score;
                                        selectScore[SemsKey_3].subjCount++;

                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            // 總分
                                            selectScore[SemsKey_3].sumScoreC1 += semesterEntry.Score;
                                            selectScore[SemsKey_3].subjCountC1++;

                                        }

                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            // 總分
                                            selectScore[SemsKey_3].sumScoreC2 += semesterEntry.Score;
                                            selectScore[SemsKey_3].subjCountC2++;

                                        }
                                    }
                                }
                                // 計算學業國防通識平均
                                if (selectScore.ContainsKey(SemsKey_3))
                                {
                                    if (selectScore[SemsKey_3].subjCount > 0)
                                        selectScore[SemsKey_3].avgScore = (decimal)(selectScore[SemsKey_3].sumScore / selectScore[SemsKey_3].subjCount);

                                    if (selectScore[SemsKey_3].subjCountC1 > 0)
                                        selectScore[SemsKey_3].avgScoreC1 = (decimal)(selectScore[SemsKey_3].sumScoreC1 / selectScore[SemsKey_3].subjCountC1);

                                    if (selectScore[SemsKey_3].subjCountC2 > 0)
                                        selectScore[SemsKey_3].avgScoreC2 = (decimal)(selectScore[SemsKey_3].sumScoreC2 / selectScore[SemsKey_3].subjCountC2);

                                    if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                    {

                                        #region 班排名
                                        {
                                            string key1 = "學業國防通識成績總分班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_3].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業國防通識成績平均班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_3].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 科排名
                                        {
                                            if (studentRec.Department != "")
                                            {
                                                string key1 = "學業國防通識成績總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[SemsKey_3].sumScore);
                                                rankStudents[key1].Add(studentID);

                                                string key2 = "學業國防通識成績平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[SemsKey_3].avgScore);
                                                rankStudents[key2].Add(studentID);
                                            }
                                        }
                                        #endregion
                                        #region 全校排名
                                        {
                                            string key1 = "學業國防通識成績總分全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_3].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業國防通識成績平均全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_3].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 類別1排名
                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            string key1 = "學業國防通識成績總分類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_3].sumScoreC1);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業國防通識成績平均類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_3].avgScoreC1);
                                            rankStudents[key2].Add(studentID);
                                        }
                                        #endregion
                                        #region 類別2排名
                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            string key1 = "學業國防通識成績總分類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_3].sumScoreC2);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業國防通識成績平均類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_3].avgScoreC2);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                    }
                                }
                                #endregion 處理國防通識

                                #region 處理實習科目
                                // 只處理年級學期，不處理科目                                     
                                string SemsKey_4 = studentID + "學業實習科目";
                                foreach (var semesterEntry in studentRec.SemesterEntryScoreList)
                                {
                                    string gs = semesterEntry.GradeYear + "" + semesterEntry.Semester;
                                    // 不在勾選年級學期內跳過
                                    if (!setting.useGradeSemesterList.Contains(gs))
                                        continue;

                                    if (semesterEntry.Entry == "實習科目")
                                    {

                                        if (!selectScore.ContainsKey(SemsKey_4))
                                            selectScore.Add(SemsKey_4, new studScore());


                                        // 處理年級學期的學期成績
                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_4].gsScore11 = semesterEntry.Score;
                                            selectScore[SemsKey_4].gsSchoolYear11 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 1 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_4].gsScore12 = semesterEntry.Score;
                                            selectScore[SemsKey_4].gsSchoolYear12 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_4].gsScore21 = semesterEntry.Score;
                                            selectScore[SemsKey_4].gsSchoolYear21 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 2 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_4].gsScore22 = semesterEntry.Score;
                                            selectScore[SemsKey_4].gsSchoolYear22 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_4].gsScore31 = semesterEntry.Score;
                                            selectScore[SemsKey_4].gsSchoolYear31 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 3 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_4].gsScore32 = semesterEntry.Score;
                                            selectScore[SemsKey_4].gsSchoolYear32 = semesterEntry.SchoolYear;
                                        }
                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 1)
                                        {
                                            selectScore[SemsKey_4].gsScore41 = semesterEntry.Score;
                                            selectScore[SemsKey_4].gsSchoolYear41 = semesterEntry.SchoolYear;
                                        }

                                        if (semesterEntry.GradeYear == 4 && semesterEntry.Semester == 2)
                                        {
                                            selectScore[SemsKey_4].gsScore42 = semesterEntry.Score;
                                            selectScore[SemsKey_4].gsSchoolYear42 = semesterEntry.SchoolYear;
                                        }
                                        // 總分
                                        selectScore[SemsKey_4].sumScore += semesterEntry.Score;
                                        selectScore[SemsKey_4].subjCount++;

                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            // 總分
                                            selectScore[SemsKey_4].sumScoreC1 += semesterEntry.Score;
                                            selectScore[SemsKey_4].subjCountC1++;

                                        }

                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            // 總分
                                            selectScore[SemsKey_4].sumScoreC2 += semesterEntry.Score;
                                            selectScore[SemsKey_4].subjCountC2++;

                                        }
                                    }
                                }
                                // 計算學業實習科目平均
                                if (selectScore.ContainsKey(SemsKey_4))
                                {
                                    if (selectScore[SemsKey_4].subjCount > 0)
                                        selectScore[SemsKey_4].avgScore = (decimal)(selectScore[SemsKey_4].sumScore / selectScore[SemsKey_4].subjCount);

                                    if (selectScore[SemsKey_4].subjCountC1 > 0)
                                        selectScore[SemsKey_4].avgScoreC1 = (decimal)(selectScore[SemsKey_4].sumScoreC1 / selectScore[SemsKey_4].subjCountC1);

                                    if (selectScore[SemsKey_4].subjCountC2 > 0)
                                        selectScore[SemsKey_4].avgScoreC2 = (decimal)(selectScore[SemsKey_4].sumScoreC2 / selectScore[SemsKey_4].subjCountC2);

                                    if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                    {
                                        #region 班排名
                                        {
                                            string key1 = "學業實習科目成績總分班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_4].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業實習科目成績平均班排名" + studentRec.RefClass.ClassID + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_4].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 科排名
                                        {
                                            if (studentRec.Department != "")
                                            {
                                                string key1 = "學業實習科目成績總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                                ranks[key1].Add(selectScore[SemsKey_4].sumScore);
                                                rankStudents[key1].Add(studentID);

                                                string key2 = "學業實習科目成績平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                                if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                                if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                                ranks[key2].Add(selectScore[SemsKey_4].avgScore);
                                                rankStudents[key2].Add(studentID);
                                            }
                                        }
                                        #endregion
                                        #region 全校排名
                                        {
                                            string key1 = "學業實習科目成績總分全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_4].sumScore);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業實習科目成績平均全校排名" + gradeyear + "^^^";
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_4].avgScore);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                        #region 類別1排名
                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            string key1 = "學業實習科目成績總分類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_4].sumScoreC1);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業實習科目成績平均類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_4].avgScoreC1);
                                            rankStudents[key2].Add(studentID);
                                        }
                                        #endregion
                                        #region 類別2排名
                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            string key1 = "學業實習科目成績總分類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                            ranks[key1].Add(selectScore[SemsKey_4].sumScoreC2);
                                            rankStudents[key1].Add(studentID);

                                            string key2 = "學業實習科目成績平均類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                            ranks[key2].Add(selectScore[SemsKey_4].avgScoreC2);
                                            rankStudents[key2].Add(studentID);

                                        }
                                        #endregion
                                    }
                                }
                                #endregion 處理實習科目

                            }   // 學業

                            bkw.ReportProgress(50);


                            // 總計成績
                            // 處理科目成績全學期總計
                            #region 處理科目成績全學期總計
                            string subjKeyAll = studentID + "總計成績";
                            selectScore.Add(subjKeyAll, new studScore());
                            // 總分,加權總分,平均,加權平均

                            foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                            {
                                string SubjName = subjectScore.Subject;
                                // 當成績不在勾選年級學期，跳過
                                string gs = subjectScore.GradeYear + "" + subjectScore.Semester;
                                if (!setting.useGradeSemesterList.Contains(gs))
                                    continue;

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
                                // 沒有任何成績跳過
                                if (match == false)
                                    continue;
                                // 判斷是否為需要的科目
                                if (setting.useSubjectPrintList.Contains(SubjName)
                                    && (
                                        replace部訂必修專業及實習科目 == false ||
                                        (
                                            subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                            && subjectScore.Require == true
                                            && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                        )
                                    )
                                {
                                    // 總分
                                    selectScore[subjKeyAll].sumScore += score;
                                    // 加總
                                    selectScore[subjKeyAll].sumScoreA += (score * subjectScore.CreditDec());
                                    // 筆數
                                    selectScore[subjKeyAll].subjCount++;
                                    // 學分加總
                                    selectScore[subjKeyAll].sumCredit += subjectScore.CreditDec();
                                }
                                // 類別一處理, 判斷此科目是否為類別1需要的
                                if (setting.useSubjecOrder1List.Contains(SubjName)
                                    && (
                                        replaceTag1部訂必修專業及實習科目 == false ||
                                        (
                                            subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                            && subjectScore.Require == true
                                            && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                        )
                                    )
                                {
                                    // 總分
                                    selectScore[subjKeyAll].sumScoreC1 += score;
                                    // 總分加權
                                    selectScore[subjKeyAll].sumScoreAC1 += (score * subjectScore.CreditDec());
                                    // 筆數
                                    selectScore[subjKeyAll].subjCountC1++;
                                    // 學分加總
                                    selectScore[subjKeyAll].sumCreditC1 += subjectScore.CreditDec();
                                }

                                // 類別二處理, 判斷此科目是否為類別2需要的
                                if (setting.useSubjecOrder2List.Contains(SubjName)
                                    && (
                                        replaceTag2部訂必修專業及實習科目 == false ||
                                        (
                                            subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                            && subjectScore.Require == true
                                            && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                        )
                                    )
                                {
                                    // 總分
                                    selectScore[subjKeyAll].sumScoreC2 += score;
                                    // 總分加權
                                    selectScore[subjKeyAll].sumScoreAC2 += (score * subjectScore.CreditDec());
                                    // 筆數
                                    selectScore[subjKeyAll].subjCountC2++;
                                    // 學分加總
                                    selectScore[subjKeyAll].sumCreditC2 += subjectScore.CreditDec();
                                }
                            }

                            // 平均
                            if (selectScore[subjKeyAll].subjCount > 0)
                                selectScore[subjKeyAll].avgScore = (decimal)(selectScore[subjKeyAll].sumScore / selectScore[subjKeyAll].subjCount);
                            // 加權平均
                            if (selectScore[subjKeyAll].sumCredit > 0)
                                selectScore[subjKeyAll].avgScoreA = (decimal)(selectScore[subjKeyAll].sumScoreA / selectScore[subjKeyAll].sumCredit);

                            // 平均(類別一)
                            if (selectScore[subjKeyAll].subjCountC1 > 0)
                                selectScore[subjKeyAll].avgScoreC1 = (decimal)(selectScore[subjKeyAll].sumScoreC1 / selectScore[subjKeyAll].subjCountC1);
                            // 加權平均(類別一)
                            if (selectScore[subjKeyAll].sumCreditC1 > 0)
                                selectScore[subjKeyAll].avgScoreAC1 = (decimal)(selectScore[subjKeyAll].sumScoreAC1 / selectScore[subjKeyAll].sumCreditC1);

                            // 平均(類別二)
                            if (selectScore[subjKeyAll].subjCountC2 > 0)
                                selectScore[subjKeyAll].avgScoreC2 = (decimal)(selectScore[subjKeyAll].sumScoreC2 / selectScore[subjKeyAll].subjCountC2);
                            // 加權平均(類別二)
                            if (selectScore[subjKeyAll].sumCreditC2 > 0)
                                selectScore[subjKeyAll].avgScoreAC2 = (decimal)(selectScore[subjKeyAll].sumScoreAC2 / selectScore[subjKeyAll].sumCreditC2);


                            if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                            {
                                #region 班排名
                                {
                                    string key1 = "總計總分班排名" + studentRec.RefClass.ClassID;
                                    if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                    ranks[key1].Add(selectScore[subjKeyAll].sumScore);
                                    rankStudents[key1].Add(studentID);

                                    string key2 = "總計加權總分班排名" + studentRec.RefClass.ClassID;
                                    if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                    ranks[key2].Add(selectScore[subjKeyAll].sumScoreA);
                                    rankStudents[key2].Add(studentID);

                                    string key3 = "總計平均班排名" + studentRec.RefClass.ClassID;
                                    if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                    ranks[key3].Add(selectScore[subjKeyAll].avgScore);
                                    rankStudents[key3].Add(studentID);


                                    string key4 = "總計加權平均班排名" + studentRec.RefClass.ClassID;
                                    if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                    ranks[key4].Add(selectScore[subjKeyAll].avgScoreA);
                                    rankStudents[key4].Add(studentID);

                                }
                                #endregion
                                #region 科排名
                                {
                                    if (studentRec.Department != "")
                                    {
                                        //各科目科排名
                                        string key1 = "總計總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                        if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                        ranks[key1].Add(selectScore[subjKeyAll].sumScore);
                                        rankStudents[key1].Add(studentID);

                                        //各科目科排名
                                        string key2 = "總計加權總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                        if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                        ranks[key2].Add(selectScore[subjKeyAll].sumScoreA);
                                        rankStudents[key2].Add(studentID);

                                        //各科目科排名
                                        string key3 = "總計平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                        if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                        ranks[key3].Add(selectScore[subjKeyAll].avgScore);
                                        rankStudents[key3].Add(studentID);

                                        //各科目科排名
                                        string key4 = "總計加權平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                        if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                        ranks[key4].Add(selectScore[subjKeyAll].avgScoreA);
                                        rankStudents[key4].Add(studentID);

                                    }
                                }
                                #endregion
                                #region 全校排名
                                {
                                    string key1 = "總計總分全校排名" + gradeyear;
                                    if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                    ranks[key1].Add(selectScore[subjKeyAll].sumScore);
                                    rankStudents[key1].Add(studentID);

                                    string key2 = "總計加權總分全校排名" + gradeyear;
                                    if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                    ranks[key2].Add(selectScore[subjKeyAll].sumScoreA);
                                    rankStudents[key2].Add(studentID);

                                    string key3 = "總計平均全校排名" + gradeyear;
                                    if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                    ranks[key3].Add(selectScore[subjKeyAll].avgScore);
                                    rankStudents[key3].Add(studentID);

                                    string key4 = "總計加權平均全校排名" + gradeyear;
                                    if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                    ranks[key4].Add(selectScore[subjKeyAll].avgScoreA);
                                    rankStudents[key4].Add(studentID);

                                }
                                #endregion
                                #region 類別1排名
                                if (studentRec.Fields.ContainsKey("tag1"))
                                {
                                    string key1 = "總計總分類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                    ranks[key1].Add(selectScore[subjKeyAll].sumScoreC1);
                                    rankStudents[key1].Add(studentID);

                                    string key2 = "總計加權總分類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                    ranks[key2].Add(selectScore[subjKeyAll].sumScoreAC1);
                                    rankStudents[key2].Add(studentID);

                                    string key3 = "總計平均類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                    ranks[key3].Add(selectScore[subjKeyAll].avgScoreC1);
                                    rankStudents[key3].Add(studentID);

                                    string key4 = "總計加權平均類別1排名" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                    ranks[key4].Add(selectScore[subjKeyAll].avgScoreAC1);
                                    rankStudents[key4].Add(studentID);
                                }
                                #endregion
                                #region 類別2排名
                                if (studentRec.Fields.ContainsKey("tag2"))
                                {
                                    string key1 = "總計總分類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key1)) ranks.Add(key1, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key1)) rankStudents.Add(key1, new List<string>());
                                    ranks[key1].Add(selectScore[subjKeyAll].sumScoreC2);
                                    rankStudents[key1].Add(studentID);

                                    string key2 = "總計加權總分類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key2)) ranks.Add(key2, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key2)) rankStudents.Add(key2, new List<string>());
                                    ranks[key2].Add(selectScore[subjKeyAll].sumScoreAC2);
                                    rankStudents[key2].Add(studentID);

                                    string key3 = "總計平均類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key3)) ranks.Add(key3, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key3)) rankStudents.Add(key3, new List<string>());
                                    ranks[key3].Add(selectScore[subjKeyAll].avgScoreC2);
                                    rankStudents[key3].Add(studentID);

                                    string key4 = "總計加權平均類別2排名" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key4)) ranks.Add(key4, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key4)) rankStudents.Add(key4, new List<string>());
                                    ranks[key4].Add(selectScore[subjKeyAll].avgScoreAC2);
                                    rankStudents[key4].Add(studentID);
                                }
                                #endregion
                            }
                            #endregion 處理科目成績全學期總計

                            // 處理科目原始成績加權平均
                            #region 處理科目原始成績各學期加權平均
                            string key5_name = "科目原始成績加權平均";//學期對就計入，不管勾選科目
                            string key7_name = "篩選科目原始成績加權平均";
                            for (int g = 1; g <= 4; g++)
                            {
                                string key5 = "", key7 = "";
                                string gsS = "";
                                for (int s = 1; s <= 2; s++)
                                {
                                    // 當成績不在勾選年級學期，跳過
                                    string gs = g + "" + s;
                                    if (!setting.useGradeSemesterList.Contains(gs))
                                        continue;

                                    if (g == 1 && s == 1) gsS = "一上";
                                    if (g == 1 && s == 2) gsS = "一下";
                                    if (g == 2 && s == 1) gsS = "二上";
                                    if (g == 2 && s == 2) gsS = "二下";
                                    if (g == 3 && s == 1) gsS = "三上";
                                    if (g == 3 && s == 2) gsS = "三下";
                                    if (g == 4 && s == 1) gsS = "四上";
                                    if (g == 4 && s == 2) gsS = "四下";

                                    key5 = studentID + gsS + key5_name;
                                    key7 = studentID + gsS + key7_name;

                                    selectScore.Add(key5, new studScore());

                                    selectScore.Add(key7, new studScore());

                                    foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                    {
                                        if (g == subjectScore.GradeYear && s == subjectScore.Semester)
                                        {
                                            decimal score = 0, Sscore;
                                            if (decimal.TryParse(subjectScore.Detail.GetAttribute("原始成績"), out Sscore))
                                            {
                                                score = Sscore;
                                                #region 處理科目原始成績
                                                // 假如是使用者勾選的科目
                                                if (setting.useSubjectPrintList.Contains(subjectScore.Subject)
                                                    && (
                                                        replace部訂必修專業及實習科目 == false ||
                                                        (
                                                            subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                            && subjectScore.Require == true
                                                            && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                        )
                                                    )
                                                {
                                                    // 總分
                                                    selectScore[key5].sumScore += score;
                                                    // 加總
                                                    selectScore[key5].sumScoreA += (score * subjectScore.CreditDec());
                                                    // 筆數
                                                    selectScore[key5].subjCount++;
                                                    // 學分加總
                                                    selectScore[key5].sumCredit += subjectScore.CreditDec();
                                                }
                                                // 類別一處理
                                                if (setting.useSubjecOrder1List.Contains(subjectScore.Subject)
                                                    && (
                                                        replaceTag1部訂必修專業及實習科目 == false ||
                                                        (
                                                            subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                            && subjectScore.Require == true
                                                            && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                        )
                                                    )
                                                {
                                                    // 總分
                                                    selectScore[key5].sumScoreC1 += score;
                                                    // 總分加權
                                                    selectScore[key5].sumScoreAC1 += (score * subjectScore.CreditDec());
                                                    // 筆數
                                                    selectScore[key5].subjCountC1++;
                                                    // 學分加總
                                                    selectScore[key5].sumCreditC1 += subjectScore.CreditDec();
                                                }
                                                // 類別二處理
                                                if (setting.useSubjecOrder2List.Contains(subjectScore.Subject)
                                                    && (
                                                        replaceTag2部訂必修專業及實習科目 == false ||
                                                        (
                                                            subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                            && subjectScore.Require == true
                                                            && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                        )
                                                    )
                                                {
                                                    // 總分
                                                    selectScore[key5].sumScoreC2 += score;
                                                    // 總分加權
                                                    selectScore[key5].sumScoreAC2 += (score * subjectScore.CreditDec());
                                                    // 筆數
                                                    selectScore[key5].subjCountC2++;
                                                    // 學分加總
                                                    selectScore[key5].sumCreditC2 += subjectScore.CreditDec();
                                                }
                                                #endregion 處理科目原始成績

                                                #region 處理篩選科目原始成績
                                                // 假如是使用者勾選的科目
                                                if (setting.useSubjectPrintList.Contains(subjectScore.Subject)
                                                    && (
                                                        replace部訂必修專業及實習科目 == false ||
                                                        (
                                                            subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                            && subjectScore.Require == true
                                                            && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                        )
                                                    )
                                                {
                                                    // 總分
                                                    selectScore[key7].sumScore += score;
                                                    // 加總
                                                    selectScore[key7].sumScoreA += (score * subjectScore.CreditDec());
                                                    // 筆數
                                                    selectScore[key7].subjCount++;
                                                    // 學分加總
                                                    selectScore[key7].sumCredit += subjectScore.CreditDec();
                                                }
                                                // 類別一處理
                                                if (setting.useSubjecOrder1List.Contains(subjectScore.Subject)
                                                    && (
                                                        replaceTag1部訂必修專業及實習科目 == false ||
                                                        (
                                                            subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                            && subjectScore.Require == true
                                                            && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                        )
                                                    )
                                                {
                                                    // 總分
                                                    selectScore[key7].sumScoreC1 += score;
                                                    // 總分加權
                                                    selectScore[key7].sumScoreAC1 += (score * subjectScore.CreditDec());
                                                    // 筆數
                                                    selectScore[key7].subjCountC1++;
                                                    // 學分加總
                                                    selectScore[key7].sumCreditC1 += subjectScore.CreditDec();
                                                }
                                                // 類別二處理
                                                if (setting.useSubjecOrder2List.Contains(subjectScore.Subject)
                                                    && (
                                                        replaceTag2部訂必修專業及實習科目 == false ||
                                                        (
                                                            subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                            && subjectScore.Require == true
                                                            && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                        )
                                                    )
                                                {
                                                    // 總分
                                                    selectScore[key7].sumScoreC2 += score;
                                                    // 總分加權
                                                    selectScore[key7].sumScoreAC2 += (score * subjectScore.CreditDec());
                                                    // 筆數
                                                    selectScore[key7].subjCountC2++;
                                                    // 學分加總
                                                    selectScore[key7].sumCreditC2 += subjectScore.CreditDec();
                                                }
                                                #endregion 處理篩選科目原始成績
                                            }
                                        }
                                    }

                                    #region 處理科目原始成績
                                    if (selectScore[key5].subjCount > 0)
                                        selectScore[key5].avgScore = (decimal)(selectScore[key5].sumScore / selectScore[key5].subjCount);
                                    // 加權平均
                                    if (selectScore[key5].sumCredit > 0)
                                        selectScore[key5].avgScoreA = (decimal)(selectScore[key5].sumScoreA / selectScore[key5].sumCredit);

                                    // 平均(類別一)
                                    if (selectScore[key5].subjCountC1 > 0)
                                        selectScore[key5].avgScoreC1 = (decimal)(selectScore[key5].sumScoreC1 / selectScore[key5].subjCountC1);
                                    // 加權平均(類別一)
                                    if (selectScore[key5].sumCreditC1 > 0)
                                        selectScore[key5].avgScoreAC1 = (decimal)(selectScore[key5].sumScoreAC1 / selectScore[key5].sumCreditC1);

                                    // 平均(類別二)
                                    if (selectScore[key5].subjCountC2 > 0)
                                        selectScore[key5].avgScoreC2 = (decimal)(selectScore[key5].sumScoreC2 / selectScore[key5].subjCountC2);
                                    // 加權平均(類別二)
                                    if (selectScore[key5].sumCreditC2 > 0)
                                        selectScore[key5].avgScoreAC2 = (decimal)(selectScore[key5].sumScoreAC2 / selectScore[key5].sumCreditC2);


                                    if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                    {
                                        // 科目原始成績加權平均班排名
                                        string key5_1 = gsS + "科目原始成績加權平均班排名" + studentRec.RefClass.ClassID;
                                        if (!ranks.ContainsKey(key5_1)) ranks.Add(key5_1, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key5_1)) rankStudents.Add(key5_1, new List<string>());
                                        ranks[key5_1].Add(selectScore[key5].avgScoreA);
                                        rankStudents[key5_1].Add(studentID);
                                        //各科目科排名
                                        key5_1 = gsS + "科目原始成績加權平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                        if (!ranks.ContainsKey(key5_1)) ranks.Add(key5_1, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key5_1)) rankStudents.Add(key5_1, new List<string>());
                                        ranks[key5_1].Add(selectScore[key5].avgScoreA);
                                        rankStudents[key5_1].Add(studentID);

                                        key5_1 = gsS + "科目原始成績加權平均校排名" + gradeyear;
                                        if (!ranks.ContainsKey(key5_1)) ranks.Add(key5_1, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key5_1)) rankStudents.Add(key5_1, new List<string>());
                                        ranks[key5_1].Add(selectScore[key5].avgScoreA);
                                        rankStudents[key5_1].Add(studentID);

                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            key5_1 = gsS + "科目原始成績加權平均類別一" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key5_1)) ranks.Add(key5_1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key5_1)) rankStudents.Add(key5_1, new List<string>());
                                            ranks[key5_1].Add(selectScore[key5].avgScoreAC1);
                                            rankStudents[key5_1].Add(studentID);
                                        }
                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            key5_1 = gsS + "科目原始成績加權平均類別二" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key5_1)) ranks.Add(key5_1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key5_1)) rankStudents.Add(key5_1, new List<string>());
                                            ranks[key5_1].Add(selectScore[key5].avgScoreAC2);
                                            rankStudents[key5_1].Add(studentID);
                                        }
                                    }
                                    #endregion 處理科目原始成績

                                    #region 處理篩選科目原始成績
                                    // 平均
                                    if (selectScore[key7].subjCount > 0)
                                        selectScore[key7].avgScore = (decimal)(selectScore[key7].sumScore / selectScore[key7].subjCount);
                                    // 加權平均
                                    if (selectScore[key7].sumCredit > 0)
                                        selectScore[key7].avgScoreA = (decimal)(selectScore[key7].sumScoreA / selectScore[key7].sumCredit);

                                    // 平均(類別一)
                                    if (selectScore[key7].subjCountC1 > 0)
                                        selectScore[key7].avgScoreC1 = (decimal)(selectScore[key7].sumScoreC1 / selectScore[key7].subjCountC1);
                                    // 加權平均(類別一)
                                    if (selectScore[key7].sumCreditC1 > 0)
                                        selectScore[key7].avgScoreAC1 = (decimal)(selectScore[key7].sumScoreAC1 / selectScore[key7].sumCreditC1);

                                    // 平均(類別二)
                                    if (selectScore[key7].subjCountC2 > 0)
                                        selectScore[key7].avgScoreC2 = (decimal)(selectScore[key7].sumScoreC2 / selectScore[key7].subjCountC2);
                                    // 加權平均(類別二)
                                    if (selectScore[key7].sumCreditC2 > 0)
                                        selectScore[key7].avgScoreAC2 = (decimal)(selectScore[key7].sumScoreAC2 / selectScore[key7].sumCreditC2);

                                    if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                    {
                                        // 篩選科目原始成績加權平均班排名
                                        string key7_1 = gsS + "篩選科目原始成績加權平均班排名" + studentRec.RefClass.ClassID;
                                        if (!ranks.ContainsKey(key7_1)) ranks.Add(key7_1, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key7_1)) rankStudents.Add(key7_1, new List<string>());
                                        ranks[key7_1].Add(selectScore[key7].avgScoreA);
                                        rankStudents[key7_1].Add(studentID);

                                        //各科目科排名
                                        key7_1 = gsS + "篩選科目原始成績加權平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                        if (!ranks.ContainsKey(key7_1)) ranks.Add(key7_1, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key7_1)) rankStudents.Add(key7_1, new List<string>());
                                        ranks[key7_1].Add(selectScore[key7].avgScoreA);
                                        rankStudents[key7_1].Add(studentID);

                                        key7_1 = gsS + "篩選科目原始成績加權平均校排名" + gradeyear;
                                        if (!ranks.ContainsKey(key7_1)) ranks.Add(key7_1, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key7_1)) rankStudents.Add(key7_1, new List<string>());
                                        ranks[key7_1].Add(selectScore[key7].avgScoreA);
                                        rankStudents[key7_1].Add(studentID);

                                        if (studentRec.Fields.ContainsKey("tag1"))
                                        {
                                            key7_1 = gsS + "篩選科目原始成績加權平均類別一" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key7_1)) ranks.Add(key7_1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key7_1)) rankStudents.Add(key7_1, new List<string>());
                                            ranks[key7_1].Add(selectScore[key7].avgScoreAC1);
                                            rankStudents[key7_1].Add(studentID);
                                        }
                                        if (studentRec.Fields.ContainsKey("tag2"))
                                        {
                                            key7_1 = gsS + "篩選科目原始成績加權平均類別二" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (!ranks.ContainsKey(key7_1)) ranks.Add(key7_1, new List<decimal>());
                                            if (!rankStudents.ContainsKey(key7_1)) rankStudents.Add(key7_1, new List<string>());
                                            ranks[key7_1].Add(selectScore[key7].avgScoreAC2);
                                            rankStudents[key7_1].Add(studentID);
                                        }
                                    }
                                    #endregion 處理篩選科目原始成績
                                }

                            } //
                            #endregion 處理科目原始成績加權平均

                            // 處理科目原始成績加權平均平均
                            #region 處理科目原始成績全學期加權平均平均
                            string key6_name = "科目原始成績加權平均平均";//學期對就計入，不管勾選科目
                            string key8_name = "篩選科目原始成績加權平均平均";

                            string key6 = studentID + key6_name;
                            string key8 = studentID + key8_name;

                            selectScore.Add(key6, new studScore());
                            selectScore.Add(key8, new studScore());

                            foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                            {
                                // 當成績不在勾選年級學期，跳過
                                string gs = subjectScore.GradeYear + "" + subjectScore.Semester;
                                if (!setting.useGradeSemesterList.Contains(gs))
                                    continue;
                                decimal score = 0, Sscore;
                                if (decimal.TryParse(subjectScore.Detail.GetAttribute("原始成績"), out Sscore))
                                {
                                    score = Sscore;

                                    #region 處理科目原始成績
                                    if (selectScore.ContainsKey(key6))
                                    {
                                        //列印科目
                                        if (setting.useSubjectPrintList.Contains(subjectScore.Subject)
                                            && (
                                                replace部訂必修專業及實習科目 == false ||
                                                (
                                                    subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                    && subjectScore.Require == true
                                                    && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                )
                                            )
                                        {
                                            // 總分
                                            selectScore[key6].sumScore += score;
                                            // 加總
                                            selectScore[key6].sumScoreA += (score * subjectScore.CreditDec());
                                            // 筆數
                                            selectScore[key6].subjCount++;
                                            // 學分加總
                                            selectScore[key6].sumCredit += subjectScore.CreditDec();
                                        }
                                        // 類別一處理
                                        if (setting.useSubjecOrder1List.Contains(subjectScore.Subject)
                                            && (
                                                replaceTag1部訂必修專業及實習科目 == false ||
                                                (
                                                    subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                    && subjectScore.Require == true
                                                    && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                )
                                            )
                                        {
                                            // 總分
                                            selectScore[key6].sumScoreC1 += score;
                                            // 總分加權
                                            selectScore[key6].sumScoreAC1 += (score * subjectScore.CreditDec());
                                            // 筆數
                                            selectScore[key6].subjCountC1++;
                                            // 學分加總
                                            selectScore[key6].sumCreditC1 += subjectScore.CreditDec();
                                        }
                                        // 類別二處理
                                        if (setting.useSubjecOrder2List.Contains(subjectScore.Subject)
                                            && (
                                                replaceTag2部訂必修專業及實習科目 == false ||
                                                (
                                                    subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                    && subjectScore.Require == true
                                                    && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                )
                                            )
                                        {
                                            // 總分
                                            selectScore[key6].sumScoreC2 += score;
                                            // 總分加權
                                            selectScore[key6].sumScoreAC2 += (score * subjectScore.CreditDec());
                                            // 筆數
                                            selectScore[key6].subjCountC2++;
                                            // 學分加總
                                            selectScore[key6].sumCreditC2 += subjectScore.CreditDec();
                                        }
                                    }
                                    #endregion 處理科目原始成績

                                    #region 處理篩選科目原始成績
                                    if (selectScore.ContainsKey(key8))
                                    {
                                        //列印科目
                                        if (setting.useSubjectPrintList.Contains(subjectScore.Subject)
                                            && (
                                                replace部訂必修專業及實習科目 == false ||
                                                (
                                                    subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                    && subjectScore.Require == true
                                                    && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                )
                                            )
                                        {
                                            // 總分
                                            selectScore[key8].sumScore += score;
                                            // 加總
                                            selectScore[key8].sumScoreA += (score * subjectScore.CreditDec());
                                            // 筆數
                                            selectScore[key8].subjCount++;
                                            // 學分加總
                                            selectScore[key8].sumCredit += subjectScore.CreditDec();
                                        }
                                        // 類別一處理
                                        if (setting.useSubjecOrder1List.Contains(subjectScore.Subject)
                                            && (
                                                replaceTag1部訂必修專業及實習科目 == false ||
                                                (
                                                    subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                    && subjectScore.Require == true
                                                    && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                )
                                            )
                                        {
                                            // 總分
                                            selectScore[key8].sumScoreC1 += score;
                                            // 總分加權
                                            selectScore[key8].sumScoreAC1 += (score * subjectScore.CreditDec());
                                            // 筆數
                                            selectScore[key8].subjCountC1++;
                                            // 學分加總
                                            selectScore[key8].sumCreditC1 += subjectScore.CreditDec();
                                        }
                                        // 類別二處理
                                        if (setting.useSubjecOrder2List.Contains(subjectScore.Subject)
                                            && (
                                                replaceTag2部訂必修專業及實習科目 == false ||
                                                (
                                                    subjectScore.Detail.GetAttribute("修課校部訂") == "部訂"
                                                    && subjectScore.Require == true
                                                    && (subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目" || subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目"))
                                                )
                                            )
                                        {
                                            // 總分
                                            selectScore[key8].sumScoreC2 += score;
                                            // 總分加權
                                            selectScore[key8].sumScoreAC2 += (score * subjectScore.CreditDec());
                                            // 筆數
                                            selectScore[key8].subjCountC2++;
                                            // 學分加總
                                            selectScore[key8].sumCreditC2 += subjectScore.CreditDec();
                                        }
                                    }
                                    #endregion 處理篩選科目原始成績
                                }
                            }

                            #region 處理科目原始成績
                            if (selectScore[key6].subjCount > 0)
                                selectScore[key6].avgScore = (decimal)(selectScore[key6].sumScore / selectScore[key6].subjCount);
                            // 加權平均
                            if (selectScore[key6].sumCredit > 0)
                                selectScore[key6].avgScoreA = (decimal)(selectScore[key6].sumScoreA / selectScore[key6].sumCredit);

                            // 平均(類別一)
                            if (selectScore[key6].subjCountC1 > 0)
                                selectScore[key6].avgScoreC1 = (decimal)(selectScore[key6].sumScoreC1 / selectScore[key6].subjCountC1);
                            // 加權平均(類別一)
                            if (selectScore[key6].sumCreditC1 > 0)
                                selectScore[key6].avgScoreAC1 = (decimal)(selectScore[key6].sumScoreAC1 / selectScore[key6].sumCreditC1);

                            // 平均(類別二)
                            if (selectScore[key6].subjCountC2 > 0)
                                selectScore[key6].avgScoreC2 = (decimal)(selectScore[key6].sumScoreC2 / selectScore[key6].subjCountC2);
                            // 加權平均(類別二)
                            if (selectScore[key6].sumCreditC2 > 0)

                                if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                                {
                                    // 科目原始成績加權平均平均班排名
                                    string key6_1 = "科目原始成績加權平均平均班排名" + studentRec.RefClass.ClassID;
                                    if (!ranks.ContainsKey(key6_1)) ranks.Add(key6_1, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key6_1)) rankStudents.Add(key6_1, new List<string>());
                                    ranks[key6_1].Add(selectScore[key6].avgScoreA);
                                    rankStudents[key6_1].Add(studentID);

                                    //各科目科排名
                                    key6_1 = "科目原始成績加權平均平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key6_1)) ranks.Add(key6_1, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key6_1)) rankStudents.Add(key6_1, new List<string>());
                                    ranks[key6_1].Add(selectScore[key6].avgScoreA);
                                    rankStudents[key6_1].Add(studentID);

                                    key6_1 = "科目原始成績加權平均平均校排名" + gradeyear;
                                    if (!ranks.ContainsKey(key6_1)) ranks.Add(key6_1, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key6_1)) rankStudents.Add(key6_1, new List<string>());
                                    ranks[key6_1].Add(selectScore[key6].avgScoreA);
                                    rankStudents[key6_1].Add(studentID);

                                    if (studentRec.Fields.ContainsKey("tag1"))
                                    {
                                        key6_1 = "科目原始成績加權平均平均類別一" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                        if (!ranks.ContainsKey(key6_1)) ranks.Add(key6_1, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key6_1)) rankStudents.Add(key6_1, new List<string>());
                                        ranks[key6_1].Add(selectScore[key6].avgScoreAC1);
                                        rankStudents[key6_1].Add(studentID);
                                    }
                                    if (studentRec.Fields.ContainsKey("tag2"))
                                    {
                                        key6_1 = "科目原始成績加權平均平均類別二" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                        if (!ranks.ContainsKey(key6_1)) ranks.Add(key6_1, new List<decimal>());
                                        if (!rankStudents.ContainsKey(key6_1)) rankStudents.Add(key6_1, new List<string>());
                                        ranks[key6_1].Add(selectScore[key6].avgScoreAC2);
                                        rankStudents[key6_1].Add(studentID);
                                    }
                                } //
                            #endregion 處理科目原始成績

                            #region 處理篩選科目原始成績
                            // 平均
                            if (selectScore[key8].subjCount > 0)
                                selectScore[key8].avgScore = (decimal)(selectScore[key8].sumScore / selectScore[key8].subjCount);
                            // 加權平均
                            if (selectScore[key8].sumCredit > 0)
                                selectScore[key8].avgScoreA = (decimal)(selectScore[key8].sumScoreA / selectScore[key8].sumCredit);

                            // 平均(類別一)
                            if (selectScore[key8].subjCountC1 > 0)
                                selectScore[key8].avgScoreC1 = (decimal)(selectScore[key8].sumScoreC1 / selectScore[key8].subjCountC1);
                            // 加權平均(類別一)
                            if (selectScore[key8].sumCreditC1 > 0)
                                selectScore[key8].avgScoreAC1 = (decimal)(selectScore[key8].sumScoreAC1 / selectScore[key8].sumCreditC1);

                            // 平均(類別二)
                            if (selectScore[key8].subjCountC2 > 0)
                                selectScore[key8].avgScoreC2 = (decimal)(selectScore[key8].sumScoreC2 / selectScore[key8].subjCountC2);
                            // 加權平均(類別二)
                            if (selectScore[key8].sumCreditC2 > 0)
                                selectScore[key8].avgScoreAC2 = (decimal)(selectScore[key8].sumScoreAC2 / selectScore[key8].sumCreditC2);

                            if (!noRankList.Contains(studentRec.StudentID))//不是不排名學生
                            {
                                // 篩選科目原始成績加權平均平均班排名
                                string key8_1 = "篩選科目原始成績加權平均平均班排名" + studentRec.RefClass.ClassID;
                                if (!ranks.ContainsKey(key8_1)) ranks.Add(key8_1, new List<decimal>());
                                if (!rankStudents.ContainsKey(key8_1)) rankStudents.Add(key8_1, new List<string>());
                                ranks[key8_1].Add(selectScore[key8].avgScoreA);
                                rankStudents[key8_1].Add(studentID);

                                //各科目科排名
                                key8_1 = "篩選科目原始成績加權平均平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                if (!ranks.ContainsKey(key8_1)) ranks.Add(key8_1, new List<decimal>());
                                if (!rankStudents.ContainsKey(key8_1)) rankStudents.Add(key8_1, new List<string>());
                                ranks[key8_1].Add(selectScore[key8].avgScoreA);
                                rankStudents[key8_1].Add(studentID);

                                key8_1 = "篩選科目原始成績加權平均平均校排名" + gradeyear;
                                if (!ranks.ContainsKey(key8_1)) ranks.Add(key8_1, new List<decimal>());
                                if (!rankStudents.ContainsKey(key8_1)) rankStudents.Add(key8_1, new List<string>());
                                ranks[key8_1].Add(selectScore[key8].avgScoreA);
                                rankStudents[key8_1].Add(studentID);

                                if (studentRec.Fields.ContainsKey("tag1"))
                                {
                                    key8_1 = "篩選科目原始成績加權平均平均類別一" + studentRec.Fields["tag1"] + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key8_1)) ranks.Add(key8_1, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key8_1)) rankStudents.Add(key8_1, new List<string>());
                                    ranks[key8_1].Add(selectScore[key8].avgScoreAC1);
                                    rankStudents[key8_1].Add(studentID);
                                }
                                if (studentRec.Fields.ContainsKey("tag2"))
                                {
                                    key8_1 = "篩選科目原始成績加權平均平均類別二" + studentRec.Fields["tag2"] + "^^^" + gradeyear;
                                    if (!ranks.ContainsKey(key8_1)) ranks.Add(key8_1, new List<decimal>());
                                    if (!rankStudents.ContainsKey(key8_1)) rankStudents.Add(key8_1, new List<string>());
                                    ranks[key8_1].Add(selectScore[key8].avgScoreAC2);
                                    rankStudents[key8_1].Add(studentID);
                                }
                            }
                            #endregion 處理篩選科目原始成績
                            #endregion 處理科目原始成績加權平均平均

                        }



                        // 排序(排名次)
                        foreach (var k in ranks.Keys)
                        {
                            var rankscores = ranks[k];
                            //排序
                            rankscores.Sort();
                            rankscores.Reverse();
                        }

                        bkw.ReportProgress(60);
                        // 填資料
                        // 科目排名
                        #region 輸出科目排名
                        subjnameIdx = 1;
                        string shtName = "科目排名" + subjnameIdx;
                        subjCount = 1;
                        BeginColumn = 0;
                        foreach (string subjName in setting.useSubjectPrintList)
                        {
                            int rowIdx = 1;
                            if (subjCount > 6)
                            {
                                subjCount = 1;
                                subjnameIdx++;
                                BeginColumn = 0;
                                shtName = "科目排名" + subjnameIdx;

                            }
                            string tii = gradeyear + "年級" + subjName;
                            foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                            {

                                _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 0].PutValue(tii);
                                _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 1].PutValue(studRec.StudentNumber);
                                _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 2].PutValue(studRec.RefClass.ClassName);
                                _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 3].PutValue(studRec.SeatNo);
                                _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 4].PutValue(studRec.StudentName);


                                //類別一分類
                                if (cat1Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 5].PutValue(cat1Dict[studRec.StudentID]);

                                // 類別二分類
                                if (cat2Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 6].PutValue(cat2Dict[studRec.StudentID]);

                                string id = studRec.StudentID + "^^^" + subjName;
                                if (selectScore.ContainsKey(id))
                                {

                                    if (selectScore[id].gsScore11.HasValue)
                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 7].PutValue(selectScore[id].gsScore11.Value);

                                    if (selectScore[id].gsScore12.HasValue)
                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 8].PutValue(selectScore[id].gsScore12.Value);

                                    if (selectScore[id].gsScore21.HasValue)
                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 9].PutValue(selectScore[id].gsScore21.Value);

                                    if (selectScore[id].gsScore22.HasValue)
                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 10].PutValue(selectScore[id].gsScore22.Value);

                                    if (selectScore[id].gsScore31.HasValue)
                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 11].PutValue(selectScore[id].gsScore31.Value);

                                    if (selectScore[id].gsScore32.HasValue)
                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 12].PutValue(selectScore[id].gsScore32.Value);

                                    // 四上
                                    if (selectScore[id].gsScore41.HasValue)
                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 13].PutValue(selectScore[id].gsScore41.Value);
                                    // 四下
                                    if (selectScore[id].gsScore42.HasValue)
                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 14].PutValue(selectScore[id].gsScore42.Value);

                                    // 總分
                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 15].PutValue(selectScore[id].sumScore);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key1 = "總分班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 16].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "總分科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 17].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "總分全校排名" + gradeyear + "^^^" + subjName;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 18].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key1 = "總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                if (ranks[key1].IndexOf(selectScore[id].sumScoreC1) >= 0)
                                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 19].PutValue(ranks[key1].IndexOf(selectScore[id].sumScoreC1) + 1);
                                            }
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key1 = "總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                if (ranks[key1].IndexOf(selectScore[id].sumScoreC2) >= 0)
                                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 20].PutValue(ranks[key1].IndexOf(selectScore[id].sumScoreC2) + 1);
                                            }
                                        }
                                    }
                                    // 平均
                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 21].PutValue(selectScore[id].avgScore);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key3 = "平均班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                        if (ranks.ContainsKey(key3))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 22].PutValue(ranks[key3].IndexOf(selectScore[id].avgScore) + 1);

                                        key3 = "平均科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                        if (ranks.ContainsKey(key3))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 23].PutValue(ranks[key3].IndexOf(selectScore[id].avgScore) + 1);

                                        key3 = "平均全校排名" + gradeyear + "^^^" + subjName;
                                        if (ranks.ContainsKey(key3))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 24].PutValue(ranks[key3].IndexOf(selectScore[id].avgScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {

                                            key3 = "平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key3))
                                            {
                                                if (ranks[key3].IndexOf(selectScore[id].avgScoreC1) >= 0)
                                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 25].PutValue(ranks[key3].IndexOf(selectScore[id].avgScoreC1) + 1);
                                            }
                                        }
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {

                                            key3 = "平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key3))
                                            {
                                                if (ranks[key3].IndexOf(selectScore[id].avgScoreC2) >= 0)
                                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 26].PutValue(ranks[key3].IndexOf(selectScore[id].avgScoreC2) + 1);
                                            }
                                        }
                                    }
                                    // 加權總分
                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 27].PutValue(selectScore[id].sumScoreA);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key2 = "加權總分班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 28].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1);

                                        key2 = "加權總分科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 29].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1);

                                        key2 = "加權總分全校排名" + gradeyear + "^^^" + subjName;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 30].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1);
                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key2 = "加權總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                if (ranks[key2].IndexOf(selectScore[id].sumScoreAC1) >= 0)
                                                {
                                                    if (ranks[key2].IndexOf(selectScore[id].sumScoreAC1) >= 0)
                                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 31].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreAC1) + 1);
                                                }
                                            }
                                        }
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key2 = "加權總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                if (ranks[key2].IndexOf(selectScore[id].sumScoreAC2) >= 0)
                                                {
                                                    if (ranks[key2].IndexOf(selectScore[id].sumScoreAC2) >= 0)
                                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 32].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreAC2) + 1);
                                                }
                                            }
                                        }
                                    }
                                    // 加權平均
                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 33].PutValue(selectScore[id].avgScoreA);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key4 = "加權平均班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                        if (ranks.ContainsKey(key4))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 34].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1);

                                        key4 = "加權平均科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                        if (ranks.ContainsKey(key4))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 35].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1);

                                        key4 = "加權平均全校排名" + gradeyear + "^^^" + subjName;
                                        if (ranks.ContainsKey(key4))
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 36].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key4 = "加權平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key4))
                                            {
                                                if (ranks[key4].IndexOf(selectScore[id].avgScoreAC1) >= 0)
                                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 37].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreAC1) + 1);
                                            }
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key4 = "加權平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key4))
                                            {
                                                if (ranks[key4].IndexOf(selectScore[id].avgScoreAC2) >= 0)
                                                    _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 38].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreAC2) + 1);
                                            }
                                        }
                                    }
                                }
                                rowIdx++;
                            }
                            subjCount++;
                            BeginColumn += 40;
                        }
                        #endregion 輸出科目排名


                        // 處理學業成績填值
                        if (setting.計算學業成績排名)
                        {
                            #region 輸出學業
                            int rowIdx2 = 1;
                            string shtName2 = "" + gradeyear + "年級學業";
                            foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                            {
                                string id = studRec.StudentID + "學業";

                                _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 0].PutValue(studRec.StudentNumber);
                                _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 1].PutValue(studRec.RefClass.ClassName);
                                _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 2].PutValue(studRec.SeatNo);
                                _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 3].PutValue(studRec.StudentName);
                                //類別一分類
                                if (cat1Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 4].PutValue(cat1Dict[studRec.StudentID]);
                                //類別二分類
                                if (cat2Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 5].PutValue(cat2Dict[studRec.StudentID]);
                                if (selectScore.ContainsKey(id))
                                {
                                    if (selectScore[id].gsScore11.HasValue)
                                        _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 6].PutValue(selectScore[id].gsScore11.Value);
                                    if (selectScore[id].gsScore12.HasValue)
                                        _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 7].PutValue(selectScore[id].gsScore12.Value);
                                    if (selectScore[id].gsScore21.HasValue)
                                        _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 8].PutValue(selectScore[id].gsScore21.Value);
                                    if (selectScore[id].gsScore22.HasValue)
                                        _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 9].PutValue(selectScore[id].gsScore22.Value);
                                    if (selectScore[id].gsScore31.HasValue)
                                        _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 10].PutValue(selectScore[id].gsScore31.Value);
                                    if (selectScore[id].gsScore32.HasValue)
                                        _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 11].PutValue(selectScore[id].gsScore32.Value);
                                    //四上
                                    if (selectScore[id].gsScore41.HasValue)
                                        _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 12].PutValue(selectScore[id].gsScore41.Value);
                                    //四下
                                    if (selectScore[id].gsScore42.HasValue)
                                        _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 13].PutValue(selectScore[id].gsScore42.Value);

                                    // 總分
                                    _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 14].PutValue(selectScore[id].sumScore);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key1 = "學業成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 15].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 16].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業成績總分全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 17].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key1 = "學業成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 18].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key1 = "學業成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 19].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }
                                    }
                                    // 平均
                                    _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 20].PutValue(selectScore[id].avgScore);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key2 = "學業成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 21].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 22].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業成績平均全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 23].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key2 = "學業成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 24].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key2 = "學業成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 25].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                    }
                                }
                                // 暫時沒用到
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 26].PutValue("加權總分");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 27].PutValue("加權總分班排名");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 28].PutValue("加權總分科排名");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 29].PutValue("加權總分校排名");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 30].PutValue("加權總分類別一排名");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 31].PutValue("加權總分類別二排名");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 32].PutValue("加權平均");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 33].PutValue("加權平均班排名");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 34].PutValue("加權平均科排名");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 35].PutValue("加權平均校排名");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 36].PutValue("加權平均類別一排名");
                                //wb.Worksheets[shtName2].Cells[rowIdx2, 37].PutValue("加權平均類別二排名");

                                rowIdx2++;
                            }
                            #endregion 輸出學業
                            // 小郭, 2014/01/02
                            #region 輸出學業(原始)
                            int rowIdx2_5 = 1;
                            string shtName2_5 = "" + gradeyear + "年級學業(原始)";
                            foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                            {
                                string id = studRec.StudentID + "學業原始";

                                _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 0].PutValue(studRec.StudentNumber);
                                _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 1].PutValue(studRec.RefClass.ClassName);
                                _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 2].PutValue(studRec.SeatNo);
                                _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 3].PutValue(studRec.StudentName);
                                //類別一分類
                                if (cat1Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 4].PutValue(cat1Dict[studRec.StudentID]);
                                //類別二分類
                                if (cat2Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 5].PutValue(cat2Dict[studRec.StudentID]);
                                if (selectScore.ContainsKey(id))
                                {
                                    if (selectScore[id].gsScore11.HasValue)
                                        _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 6].PutValue(selectScore[id].gsScore11.Value);
                                    if (selectScore[id].gsScore12.HasValue)
                                        _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 7].PutValue(selectScore[id].gsScore12.Value);
                                    if (selectScore[id].gsScore21.HasValue)
                                        _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 8].PutValue(selectScore[id].gsScore21.Value);
                                    if (selectScore[id].gsScore22.HasValue)
                                        _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 9].PutValue(selectScore[id].gsScore22.Value);
                                    if (selectScore[id].gsScore31.HasValue)
                                        _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 10].PutValue(selectScore[id].gsScore31.Value);
                                    if (selectScore[id].gsScore32.HasValue)
                                        _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 11].PutValue(selectScore[id].gsScore32.Value);
                                    //四上
                                    if (selectScore[id].gsScore41.HasValue)
                                        _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 12].PutValue(selectScore[id].gsScore41.Value);
                                    //四下
                                    if (selectScore[id].gsScore42.HasValue)
                                        _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 13].PutValue(selectScore[id].gsScore42.Value);

                                    // 總分
                                    _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 14].PutValue(selectScore[id].sumScore);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key1 = "學業原始成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 15].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業原始成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 16].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業原始成績總分全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 17].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key1 = "學業原始成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 18].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key1 = "學業原始成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 19].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }
                                    }
                                    // 平均
                                    _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 20].PutValue(selectScore[id].avgScore);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key2 = "學業原始成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 21].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業原始成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 22].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業原始成績平均全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 23].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key2 = "學業原始成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 24].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key2 = "學業原始成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 25].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                    }
                                }

                                rowIdx2_5++;
                            }
                            #endregion 輸出學業(原始)

                            #region 輸出體育
                            int rowIdx2_1 = 1;
                            string shtName2_1 = "" + gradeyear + "年級學業體育";
                            foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                            {
                                string id = studRec.StudentID + "學業體育";
                                _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 0].PutValue(studRec.StudentNumber);
                                _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 1].PutValue(studRec.RefClass.ClassName);
                                _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 2].PutValue(studRec.SeatNo);
                                _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 3].PutValue(studRec.StudentName);
                                //類別一分類
                                if (cat1Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 4].PutValue(cat1Dict[studRec.StudentID]);
                                //類別二分類
                                if (cat2Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 5].PutValue(cat2Dict[studRec.StudentID]);
                                if (selectScore.ContainsKey(id))
                                {
                                    if (selectScore[id].gsScore11.HasValue)
                                        _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 6].PutValue(selectScore[id].gsScore11.Value);
                                    if (selectScore[id].gsScore12.HasValue)
                                        _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 7].PutValue(selectScore[id].gsScore12.Value);
                                    if (selectScore[id].gsScore21.HasValue)
                                        _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 8].PutValue(selectScore[id].gsScore21.Value);
                                    if (selectScore[id].gsScore22.HasValue)
                                        _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 9].PutValue(selectScore[id].gsScore22.Value);
                                    if (selectScore[id].gsScore31.HasValue)
                                        _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 10].PutValue(selectScore[id].gsScore31.Value);
                                    if (selectScore[id].gsScore32.HasValue)
                                        _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 11].PutValue(selectScore[id].gsScore32.Value);
                                    //四上
                                    if (selectScore[id].gsScore41.HasValue)
                                        _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 12].PutValue(selectScore[id].gsScore41.Value);
                                    //四下
                                    if (selectScore[id].gsScore42.HasValue)
                                        _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 13].PutValue(selectScore[id].gsScore42.Value);

                                    // 總分
                                    _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 14].PutValue(selectScore[id].sumScore);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key1 = "學業體育成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 15].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業體育成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 16].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業體育成績總分全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 17].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key1 = "學業體育成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 18].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key1 = "學業體育成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 19].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }
                                    }
                                    // 平均
                                    _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 20].PutValue(selectScore[id].avgScore);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key2 = "學業體育成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 21].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業體育成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 22].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業體育成績平均全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 23].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key2 = "學業體育成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 24].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key2 = "學業體育成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 25].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                    }
                                }

                                rowIdx2_1++;
                            }
                            #endregion 輸出體育

                            #region 輸出健康與護理
                            int rowIdx2_2 = 1;
                            string shtName2_2 = "" + gradeyear + "年級學業健康與護理";
                            foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                            {

                                string id = studRec.StudentID + "學業健康與護理";
                                _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 0].PutValue(studRec.StudentNumber);
                                _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 1].PutValue(studRec.RefClass.ClassName);
                                _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 2].PutValue(studRec.SeatNo);
                                _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 3].PutValue(studRec.StudentName);
                                //類別一分類
                                if (cat1Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 4].PutValue(cat1Dict[studRec.StudentID]);
                                //類別二分類
                                if (cat2Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 5].PutValue(cat2Dict[studRec.StudentID]);
                                if (selectScore.ContainsKey(id))
                                {
                                    if (selectScore[id].gsScore11.HasValue)
                                        _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 6].PutValue(selectScore[id].gsScore11.Value);
                                    if (selectScore[id].gsScore12.HasValue)
                                        _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 7].PutValue(selectScore[id].gsScore12.Value);
                                    if (selectScore[id].gsScore21.HasValue)
                                        _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 8].PutValue(selectScore[id].gsScore21.Value);
                                    if (selectScore[id].gsScore22.HasValue)
                                        _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 9].PutValue(selectScore[id].gsScore22.Value);
                                    if (selectScore[id].gsScore31.HasValue)
                                        _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 10].PutValue(selectScore[id].gsScore31.Value);
                                    if (selectScore[id].gsScore32.HasValue)
                                        _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 11].PutValue(selectScore[id].gsScore32.Value);
                                    //四上
                                    if (selectScore[id].gsScore41.HasValue)
                                        _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 12].PutValue(selectScore[id].gsScore41.Value);
                                    //四下
                                    if (selectScore[id].gsScore42.HasValue)
                                        _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 13].PutValue(selectScore[id].gsScore42.Value);

                                    // 總分
                                    _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 14].PutValue(selectScore[id].sumScore);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key1 = "學業健康與護理成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 15].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業健康與護理成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 16].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業健康與護理成績總分全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 17].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key1 = "學業健康與護理成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 18].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key1 = "學業健康與護理成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 19].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }
                                    }
                                    // 平均
                                    _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 20].PutValue(selectScore[id].avgScore);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key2 = "學業健康與護理成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 21].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業健康與護理成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 22].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業健康與護理成績平均全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 23].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key2 = "學業健康與護理成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 24].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key2 = "學業健康與護理成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 25].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                    }
                                }

                                rowIdx2_2++;
                            }
                            #endregion 輸出健康與護理

                            #region 輸出國防通識
                            int rowIdx2_3 = 1;
                            string shtName2_3 = "" + gradeyear + "年級學業國防通識";
                            foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                            {
                                string id = studRec.StudentID + "學業國防通識";
                                _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 0].PutValue(studRec.StudentNumber);
                                _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 1].PutValue(studRec.RefClass.ClassName);
                                _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 2].PutValue(studRec.SeatNo);
                                _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 3].PutValue(studRec.StudentName);
                                //類別一分類
                                if (cat1Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 4].PutValue(cat1Dict[studRec.StudentID]);
                                //類別二分類
                                if (cat2Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 5].PutValue(cat2Dict[studRec.StudentID]);
                                if (selectScore.ContainsKey(id))
                                {
                                    if (selectScore[id].gsScore11.HasValue)
                                        _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 6].PutValue(selectScore[id].gsScore11.Value);
                                    if (selectScore[id].gsScore12.HasValue)
                                        _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 7].PutValue(selectScore[id].gsScore12.Value);
                                    if (selectScore[id].gsScore21.HasValue)
                                        _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 8].PutValue(selectScore[id].gsScore21.Value);
                                    if (selectScore[id].gsScore22.HasValue)
                                        _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 9].PutValue(selectScore[id].gsScore22.Value);
                                    if (selectScore[id].gsScore31.HasValue)
                                        _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 10].PutValue(selectScore[id].gsScore31.Value);
                                    if (selectScore[id].gsScore32.HasValue)
                                        _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 11].PutValue(selectScore[id].gsScore32.Value);
                                    //四上
                                    if (selectScore[id].gsScore41.HasValue)
                                        _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 12].PutValue(selectScore[id].gsScore41.Value);
                                    //四下
                                    if (selectScore[id].gsScore42.HasValue)
                                        _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 13].PutValue(selectScore[id].gsScore42.Value);

                                    // 總分
                                    _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 14].PutValue(selectScore[id].sumScore);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key1 = "學業國防通識成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 15].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業國防通識成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 16].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業國防通識成績總分全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 17].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key1 = "學業國防通識成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 18].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key1 = "學業國防通識成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 19].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }
                                    }
                                    // 平均
                                    _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 20].PutValue(selectScore[id].avgScore);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key2 = "學業國防通識成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 21].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業國防通識成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 22].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業國防通識成績平均全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 23].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key2 = "學業國防通識成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 24].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key2 = "學業國防通識成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 25].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                    }
                                }

                                rowIdx2_3++;
                            }
                            #endregion 輸出國防通識

                            #region 輸出實習科目
                            int rowIdx2_4 = 1;
                            string shtName2_4 = "" + gradeyear + "年級學業實習科目";
                            foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                            {
                                string id = studRec.StudentID + "學業實習科目";
                                _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 0].PutValue(studRec.StudentNumber);
                                _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 1].PutValue(studRec.RefClass.ClassName);
                                _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 2].PutValue(studRec.SeatNo);
                                _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 3].PutValue(studRec.StudentName);
                                //類別一分類
                                if (cat1Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 4].PutValue(cat1Dict[studRec.StudentID]);
                                //類別二分類
                                if (cat2Dict.ContainsKey(studRec.StudentID))
                                    _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 5].PutValue(cat2Dict[studRec.StudentID]);
                                if (selectScore.ContainsKey(id))
                                {
                                    if (selectScore[id].gsScore11.HasValue)
                                        _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 6].PutValue(selectScore[id].gsScore11.Value);
                                    if (selectScore[id].gsScore12.HasValue)
                                        _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 7].PutValue(selectScore[id].gsScore12.Value);
                                    if (selectScore[id].gsScore21.HasValue)
                                        _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 8].PutValue(selectScore[id].gsScore21.Value);
                                    if (selectScore[id].gsScore22.HasValue)
                                        _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 9].PutValue(selectScore[id].gsScore22.Value);
                                    if (selectScore[id].gsScore31.HasValue)
                                        _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 10].PutValue(selectScore[id].gsScore31.Value);
                                    if (selectScore[id].gsScore32.HasValue)
                                        _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 11].PutValue(selectScore[id].gsScore32.Value);
                                    //四上
                                    if (selectScore[id].gsScore41.HasValue)
                                        _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 12].PutValue(selectScore[id].gsScore41.Value);
                                    //四下
                                    if (selectScore[id].gsScore42.HasValue)
                                        _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 13].PutValue(selectScore[id].gsScore42.Value);

                                    // 總分
                                    _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 14].PutValue(selectScore[id].sumScore);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key1 = "學業實習科目成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 15].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業實習科目成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 16].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "學業實習科目成績總分全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 17].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key1 = "學業實習科目成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 18].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key1 = "學業實習科目成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 19].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                        }
                                    }
                                    // 平均
                                    _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 20].PutValue(selectScore[id].avgScore);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key2 = "學業實習科目成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 21].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業實習科目成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 22].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        key2 = "學業實習科目成績平均全校排名" + gradeyear + "^^^";
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 23].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key2 = "學業實習科目成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 24].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key2 = "學業實習科目成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 25].PutValue(ranks[key2].IndexOf(selectScore[id].avgScore) + 1);
                                        }
                                    }
                                }

                                rowIdx2_4++;
                            }
                            #endregion 輸出實習科目

                        }
                        bkw.ReportProgress(80);
                        // 處理總計排名並填值
                        #region 處理總計排名並填值
                        int rowIdx3 = 1;
                        string shtName3 = "" + gradeyear + "年級總計成績";

                        foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                        {

                            string id = studRec.StudentID + "總計成績";
                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 0].PutValue(studRec.StudentNumber);
                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 1].PutValue(studRec.RefClass.ClassName);
                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 2].PutValue(studRec.SeatNo);
                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 3].PutValue(studRec.StudentName);
                            //類別一分類
                            if (cat1Dict.ContainsKey(studRec.StudentID))
                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 4].PutValue(cat1Dict[studRec.StudentID]);
                            // 類別二分類
                            if (cat2Dict.ContainsKey(studRec.StudentID))
                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 5].PutValue(cat2Dict[studRec.StudentID]);

                            if (selectScore.ContainsKey(id))
                            {
                                string key1 = "";
                                // 總計總分類別1
                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 6].PutValue(selectScore[id].sumScoreC1);
                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                {
                                    if (studRec.Fields.ContainsKey("tag1"))
                                    {
                                        key1 = "總計總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 7].PutValue(ranks[key1].IndexOf(selectScore[id].sumScoreC1) + 1);
                                    }
                                }

                                // 總計總分類別2
                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 14].PutValue(selectScore[id].sumScoreC2);
                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                {
                                    if (studRec.Fields.ContainsKey("tag2"))
                                    {
                                        key1 = "總計總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 15].PutValue(ranks[key1].IndexOf(selectScore[id].sumScoreC1) + 1);
                                    }
                                }
                                // 總分
                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 22].PutValue(selectScore[id].sumScore);

                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                {
                                    key1 = "總計總分班排名" + studRec.RefClass.ClassID;
                                    if (ranks.ContainsKey(key1))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 23].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                    key1 = "總計總分科排名" + studRec.Department + "^^^" + gradeyear;
                                    if (ranks.ContainsKey(key1))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 24].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                    key1 = "總計總分全校排名" + gradeyear;
                                    if (ranks.ContainsKey(key1))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 25].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);
                                }
                                // 平均
                                string key3 = "";
                                // 平均(類別一)
                                if (studRec.Fields.ContainsKey("tag1"))
                                {
                                    key3 = "總計平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                    _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 8].PutValue(selectScore[id].avgScoreC1);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        if (ranks.ContainsKey(key3))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 9].PutValue(ranks[key3].IndexOf(selectScore[id].avgScoreC1) + 1);
                                    }
                                }
                                // 平均(類別二)
                                if (studRec.Fields.ContainsKey("tag2"))
                                {
                                    key3 = "總計平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                    _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 16].PutValue(selectScore[id].avgScoreC2);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        if (ranks.ContainsKey(key3))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 17].PutValue(ranks[key3].IndexOf(selectScore[id].avgScoreC2) + 1);
                                    }
                                }
                                // 平均
                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 26].PutValue(selectScore[id].avgScore);

                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                {
                                    key3 = "總計平均班排名" + studRec.RefClass.ClassID;
                                    if (ranks.ContainsKey(key3))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 27].PutValue(ranks[key3].IndexOf(selectScore[id].avgScore) + 1);

                                    key3 = "總計平均科排名" + studRec.Department + "^^^" + gradeyear;

                                    if (ranks.ContainsKey(key3))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 28].PutValue(ranks[key3].IndexOf(selectScore[id].avgScore) + 1);

                                    key3 = "總計平均全校排名" + gradeyear;
                                    if (ranks.ContainsKey(key3))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 29].PutValue(ranks[key3].IndexOf(selectScore[id].avgScore) + 1);
                                }

                                string key2 = "";
                                if (studRec.Fields.ContainsKey("tag1"))
                                {
                                    key2 = "總計加權總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                    // 加權總分(類別一)
                                    _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 10].PutValue(selectScore[id].sumScoreAC1);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 11].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreAC1) + 1);
                                    }
                                }

                                if (studRec.Fields.ContainsKey("tag2"))
                                {
                                    key2 = "總計加權總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                    // 加權總分(類別二)
                                    _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 18].PutValue(selectScore[id].sumScoreAC2);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 19].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreAC2) + 1);
                                    }
                                }
                                // 加權總分
                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 30].PutValue(selectScore[id].sumScoreA);
                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                {
                                    key2 = "總計加權總分班排名" + studRec.RefClass.ClassID;
                                    if (ranks.ContainsKey(key2))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 31].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1);

                                    key2 = "總計加權總分科排名" + studRec.Department + "^^^" + gradeyear;
                                    if (ranks.ContainsKey(key2))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 32].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1);

                                    key2 = "總計加權總分全校排名" + gradeyear;
                                    if (ranks.ContainsKey(key2))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 33].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1);
                                }

                                string key4 = "";

                                if (studRec.Fields.ContainsKey("tag1"))
                                {
                                    key4 = "總計加權平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                    _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 12].PutValue(selectScore[id].avgScoreAC1);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        if (ranks.ContainsKey(key4))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 13].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreAC1) + 1);
                                    }
                                }

                                if (studRec.Fields.ContainsKey("tag2"))
                                {
                                    key4 = "總計加權平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                    _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 20].PutValue(selectScore[id].avgScoreAC2);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        if (ranks.ContainsKey(key4))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 21].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreAC2) + 1);
                                    }
                                }

                                // 加權平均
                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 34].PutValue(selectScore[id].avgScoreA);
                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                {
                                    key4 = "總計加權平均班排名" + studRec.RefClass.ClassID;
                                    if (ranks.ContainsKey(key4))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 35].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1);
                                    key4 = "總計加權平均科排名" + studRec.Department + "^^^" + gradeyear;
                                    if (ranks.ContainsKey(key4))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 36].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1);
                                    key4 = "總計加權平均全校排名" + gradeyear;
                                    if (ranks.ContainsKey(key4))
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 37].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1);
                                }
                            }

                            rowIdx3++;
                        }
                        #endregion 處理總計排名並填值

                        // Excel 存檔
                        #region Save Excl
                        #region 處理超過樣版支援名單
                        int overflowCount = 1;
                        foreach (KeyValuePair<string, List<string>> data in _studSubjCountDict)
                        {
                            if (data.Value.Count > setting.SubjectLimit)
                            {
                                string errName = "學生科目數超出範本科目數清單";
                                if (!_ExcelCellDict.ContainsKey(errName))
                                {
                                    _ExcelCellDict.Add(errName, new Aspose.Cells.Workbook());
                                    _ExcelCellDict[errName].Worksheets[0].Name = errName;
                                    var overflowSheet = _ExcelCellDict[errName].Worksheets[0];
                                    overflowSheet.Cells[0, 0].PutValue("學號");
                                    overflowSheet.Cells[0, 1].PutValue("班級");
                                    overflowSheet.Cells[0, 2].PutValue("座號");
                                    overflowSheet.Cells[0, 3].PutValue("姓名");
                                    overflowSheet.Cells[0, 4].PutValue("成績科目數");
                                    overflowSheet.Cells[0, 5].PutValue("樣版科目數");
                                }
                                else
                                {
                                    var overflowSheet = _ExcelCellDict[errName].Worksheets[0];
                                    overflowSheet.Cells[overflowCount, 0].PutValue(accessHelper.StudentHelper.GetStudent(data.Key).StudentNumber);
                                    overflowSheet.Cells[overflowCount, 1].PutValue(accessHelper.StudentHelper.GetStudent(data.Key).RefClass == null ? "" : accessHelper.StudentHelper.GetStudent(data.Key).RefClass.ClassName);
                                    overflowSheet.Cells[overflowCount, 2].PutValue(accessHelper.StudentHelper.GetStudent(data.Key).SeatNo);
                                    overflowSheet.Cells[overflowCount, 3].PutValue(accessHelper.StudentHelper.GetStudent(data.Key).StudentName);
                                    overflowSheet.Cells[overflowCount, 4].PutValue(data.Value);
                                    overflowSheet.Cells[overflowCount, 5].PutValue(setting.SubjectLimit);
                                    overflowCount++;
                                }
                            }
                        }
                        #endregion

                        if (!Directory.Exists(Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports", FolderName)))
                            Directory.CreateDirectory(Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports", FolderName));
                        foreach (string key in _ExcelCellDict.Keys)
                        {
                            Aspose.Cells.Workbook data = _ExcelCellDict[key];

                            try
                            {
                                foreach (Aspose.Cells.Worksheet sheet in data.Worksheets)
                                {
                                    if (sheet.Cells.Rows.Count > 0)
                                        sheet.FreezePanes(1, 0, 1, sheet.Cells.MaxColumn);
                                }
                            }
                            catch (Exception ex)
                            {

                            }

                            string inputReportName = "多學期科目成績固定排名";
                            string reportName = "E_" + key + inputReportName;


                            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports", FolderName, "Detials");
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
                                foreach (memoText mt in _memoText)
                                {
                                    if (mt.WorkSheetName == key)
                                        mt.FileAddress = path;
                                }

                                data.Save(path, Aspose.Cells.FileFormatType.Excel97To2003);
                            }
                            catch (OutOfMemoryException exo1)
                            {
                                exc = exo1;
                            }

                        }
                        #endregion

                        #region Excel檔案索引
                        if (_memoText.Count > 0)
                        {
                            Aspose.Cells.Workbook wbm = new Aspose.Cells.Workbook();
                            wbm.Worksheets[0].Cells[0, 0].PutValue("名稱");
                            wbm.Worksheets[0].Cells[0, 1].PutValue("連結");
                            var linkStyle = wbm.Styles[wbm.Styles.Add()];
                            linkStyle.Font.Color = wbm.GetMatchingColor(System.Drawing.Color.Blue);
                            int idx = 1;

                            foreach (memoText mt in _memoText)
                            {
                                if (!string.IsNullOrEmpty(mt.Memo))
                                {
                                    wbm.Worksheets[0].Cells[idx, 0].PutValue(mt.Memo);
                                    wbm.Worksheets[0].Cells[idx, 1].PutValue("連結");
                                    //wbm.Worksheets[0].Cells[idx, 1].Style = linkStyle;
                                    wbm.Worksheets[0].Cells[idx, 1].SetStyle(linkStyle);
                                    wbm.Worksheets[0].Hyperlinks.Add(idx, 1, 1, 1, mt.FileAddress);
                                    idx++;
                                }
                            }
                            wbm.Worksheets[0].AutoFitColumns();
                            string pathaa = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports", FolderName);
                            wbm.Save(pathaa + "\\_索引_詳細成績資料.xls", Aspose.Cells.FileFormatType.Excel97To2003);
                        }
                        #endregion

                        _ExcelCellDict.Clear();
                        GC.Collect();

                        // 處理Word
                        // 依班級名稱分割
                        _WordDocDict.Clear();


                        // 取得班級名稱 
                        List<string> classNameList = new List<string>();
                        foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                            if (!classNameList.Contains(studRec.RefClass.ClassName))
                                classNameList.Add(studRec.RefClass.ClassName);

                        string SchoolName = K12.Data.School.ChineseName;

                        _table.Columns.Clear();
                        _table.Clear();
                        #region 合併欄位使用
                        // 固定
                        _table.Columns.Add("學校名稱");
                        _table.Columns.Add("科別");
                        _table.Columns.Add("班級");
                        _table.Columns.Add("座號");
                        _table.Columns.Add("學號");
                        _table.Columns.Add("學生系統編號");
                        _table.Columns.Add("教師系統編號");
                        _table.Columns.Add("教師姓名");
                        _table.Columns.Add("姓名");
                        _table.Columns.Add("類別一分類");
                        _table.Columns.Add("類別二分類");
                        _table.Columns.Add("類別排名1");
                        _table.Columns.Add("類別排名2");
                        _table.Columns.Add("一年級學年度");
                        _table.Columns.Add("二年級學年度");
                        _table.Columns.Add("三年級學年度");
                        _table.Columns.Add("四年級學年度");
                        // 學業類 學業
                        _table.Columns.Add("一上學業成績");
                        _table.Columns.Add("一下學業成績");
                        _table.Columns.Add("二上學業成績");
                        _table.Columns.Add("二下學業成績");
                        _table.Columns.Add("三上學業成績");
                        _table.Columns.Add("三下學業成績");
                        _table.Columns.Add("四上學業成績");
                        _table.Columns.Add("四下學業成績");
                        _table.Columns.Add("一上學業班排名");
                        _table.Columns.Add("一下學業班排名");
                        _table.Columns.Add("二上學業班排名");
                        _table.Columns.Add("二下學業班排名");
                        _table.Columns.Add("三上學業班排名");
                        _table.Columns.Add("三下學業班排名");
                        _table.Columns.Add("四上學業班排名");
                        _table.Columns.Add("四下學業班排名");
                        _table.Columns.Add("一上學業科排名");
                        _table.Columns.Add("一下學業科排名");
                        _table.Columns.Add("二上學業科排名");
                        _table.Columns.Add("二下學業科排名");
                        _table.Columns.Add("三上學業科排名");
                        _table.Columns.Add("三下學業科排名");
                        _table.Columns.Add("四上學業科排名");
                        _table.Columns.Add("四下學業科排名");
                        _table.Columns.Add("一上學業校排名");
                        _table.Columns.Add("一下學業校排名");
                        _table.Columns.Add("二上學業校排名");
                        _table.Columns.Add("二下學業校排名");
                        _table.Columns.Add("三上學業校排名");
                        _table.Columns.Add("三下學業校排名");
                        _table.Columns.Add("四上學業校排名");
                        _table.Columns.Add("四下學業校排名");
                        _table.Columns.Add("一上學業類1排名");
                        _table.Columns.Add("一下學業類1排名");
                        _table.Columns.Add("二上學業類1排名");
                        _table.Columns.Add("二下學業類1排名");
                        _table.Columns.Add("三上學業類1排名");
                        _table.Columns.Add("三下學業類1排名");
                        _table.Columns.Add("四上學業類1排名");
                        _table.Columns.Add("四下學業類1排名");
                        _table.Columns.Add("一上學業班排名母數");
                        _table.Columns.Add("一下學業班排名母數");
                        _table.Columns.Add("二上學業班排名母數");
                        _table.Columns.Add("二下學業班排名母數");
                        _table.Columns.Add("三上學業班排名母數");
                        _table.Columns.Add("三下學業班排名母數");
                        _table.Columns.Add("四上學業班排名母數");
                        _table.Columns.Add("四下學業班排名母數");
                        _table.Columns.Add("一上學業科排名母數");
                        _table.Columns.Add("一下學業科排名母數");
                        _table.Columns.Add("二上學業科排名母數");
                        _table.Columns.Add("二下學業科排名母數");
                        _table.Columns.Add("三上學業科排名母數");
                        _table.Columns.Add("三下學業科排名母數");
                        _table.Columns.Add("四上學業科排名母數");
                        _table.Columns.Add("四下學業科排名母數");
                        _table.Columns.Add("一上學業校排名母數");
                        _table.Columns.Add("一下學業校排名母數");
                        _table.Columns.Add("二上學業校排名母數");
                        _table.Columns.Add("二下學業校排名母數");
                        _table.Columns.Add("三上學業校排名母數");
                        _table.Columns.Add("三下學業校排名母數");
                        _table.Columns.Add("四上學業校排名母數");
                        _table.Columns.Add("四下學業校排名母數");
                        _table.Columns.Add("一上學業類1排名母數");
                        _table.Columns.Add("一下學業類1排名母數");
                        _table.Columns.Add("二上學業類1排名母數");
                        _table.Columns.Add("二下學業類1排名母數");
                        _table.Columns.Add("三上學業類1排名母數");
                        _table.Columns.Add("三下學業類1排名母數");
                        _table.Columns.Add("四上學業類1排名母數");
                        _table.Columns.Add("四下學業類1排名母數");
                        _table.Columns.Add("一上學業班排名百分比");
                        _table.Columns.Add("一下學業班排名百分比");
                        _table.Columns.Add("二上學業班排名百分比");
                        _table.Columns.Add("二下學業班排名百分比");
                        _table.Columns.Add("三上學業班排名百分比");
                        _table.Columns.Add("三下學業班排名百分比");
                        _table.Columns.Add("四上學業班排名百分比");
                        _table.Columns.Add("四下學業班排名百分比");
                        _table.Columns.Add("一上學業科排名百分比");
                        _table.Columns.Add("一下學業科排名百分比");
                        _table.Columns.Add("二上學業科排名百分比");
                        _table.Columns.Add("二下學業科排名百分比");
                        _table.Columns.Add("三上學業科排名百分比");
                        _table.Columns.Add("三下學業科排名百分比");
                        _table.Columns.Add("四上學業科排名百分比");
                        _table.Columns.Add("四下學業科排名百分比");
                        _table.Columns.Add("一上學業校排名百分比");
                        _table.Columns.Add("一下學業校排名百分比");
                        _table.Columns.Add("二上學業校排名百分比");
                        _table.Columns.Add("二下學業校排名百分比");
                        _table.Columns.Add("三上學業校排名百分比");
                        _table.Columns.Add("三下學業校排名百分比");
                        _table.Columns.Add("四上學業校排名百分比");
                        _table.Columns.Add("四下學業校排名百分比");
                        _table.Columns.Add("一上學業類1排名百分比");
                        _table.Columns.Add("一下學業類1排名百分比");
                        _table.Columns.Add("二上學業類1排名百分比");
                        _table.Columns.Add("二下學業類1排名百分比");
                        _table.Columns.Add("三上學業類1排名百分比");
                        _table.Columns.Add("三下學業類1排名百分比");
                        _table.Columns.Add("四上學業類1排名百分比");
                        _table.Columns.Add("四下學業類1排名百分比");

                        _table.Columns.Add("學業平均");
                        _table.Columns.Add("學業平均科排名");
                        _table.Columns.Add("學業平均科排名母數");
                        _table.Columns.Add("學業平均校排名");
                        _table.Columns.Add("學業平均校排名母數");
                        _table.Columns.Add("學業平均班排名");
                        _table.Columns.Add("學業平均班排名母數");
                        _table.Columns.Add("學業平均類別一排名");
                        _table.Columns.Add("學業平均類別一排名母數");
                        _table.Columns.Add("學業平均類別二排名");
                        _table.Columns.Add("學業平均類別二排名母數");
                        _table.Columns.Add("學業總分");
                        _table.Columns.Add("學業總分科排名");
                        _table.Columns.Add("學業總分科排名母數");
                        _table.Columns.Add("學業總分校排名");
                        _table.Columns.Add("學業總分校排名母數");
                        _table.Columns.Add("學業總分班排名");
                        _table.Columns.Add("學業總分班排名母數");
                        _table.Columns.Add("學業總分類別一排名");
                        _table.Columns.Add("學業總分類別一排名母數");
                        _table.Columns.Add("學業總分類別二排名");
                        _table.Columns.Add("學業總分類別二排名母數");
                        _table.Columns.Add("學業平均科排名百分比");
                        _table.Columns.Add("學業平均校排名百分比");
                        _table.Columns.Add("學業平均班排名百分比");
                        _table.Columns.Add("學業平均類別一排名百分比");
                        _table.Columns.Add("學業平均類別二排名百分比");

                        // 學業類 學業原始
                        _table.Columns.Add("一上學業原始成績");
                        _table.Columns.Add("一下學業原始成績");
                        _table.Columns.Add("二上學業原始成績");
                        _table.Columns.Add("二下學業原始成績");
                        _table.Columns.Add("三上學業原始成績");
                        _table.Columns.Add("三下學業原始成績");
                        _table.Columns.Add("四上學業原始成績");
                        _table.Columns.Add("四下學業原始成績");
                        _table.Columns.Add("學業原始平均");
                        _table.Columns.Add("學業原始平均科排名");
                        _table.Columns.Add("學業原始平均科排名母數");
                        _table.Columns.Add("學業原始平均校排名");
                        _table.Columns.Add("學業原始平均校排名母數");
                        _table.Columns.Add("學業原始平均班排名");
                        _table.Columns.Add("學業原始平均班排名母數");
                        _table.Columns.Add("學業原始平均類別一排名");
                        _table.Columns.Add("學業原始平均類別一排名母數");
                        _table.Columns.Add("學業原始平均類別二排名");
                        _table.Columns.Add("學業原始平均類別二排名母數");
                        _table.Columns.Add("學業原始總分");
                        _table.Columns.Add("學業原始總分科排名");
                        _table.Columns.Add("學業原始總分科排名母數");
                        _table.Columns.Add("學業原始總分校排名");
                        _table.Columns.Add("學業原始總分校排名母數");
                        _table.Columns.Add("學業原始總分班排名");
                        _table.Columns.Add("學業原始總分班排名母數");
                        _table.Columns.Add("學業原始總分類別一排名");
                        _table.Columns.Add("學業原始總分類別一排名母數");
                        _table.Columns.Add("學業原始總分類別二排名");
                        _table.Columns.Add("學業原始總分類別二排名母數");
                        _table.Columns.Add("學業原始平均科排名百分比");
                        _table.Columns.Add("學業原始平均校排名百分比");
                        _table.Columns.Add("學業原始平均班排名百分比");
                        _table.Columns.Add("學業原始平均類別一排名百分比");
                        _table.Columns.Add("學業原始平均類別二排名百分比");

                        _table.Columns.Add("一上學業體育成績");
                        _table.Columns.Add("一下學業體育成績");
                        _table.Columns.Add("二上學業體育成績");
                        _table.Columns.Add("二下學業體育成績");
                        _table.Columns.Add("三上學業體育成績");
                        _table.Columns.Add("三下學業體育成績");
                        _table.Columns.Add("四上學業體育成績");
                        _table.Columns.Add("四下學業體育成績");
                        _table.Columns.Add("學業體育平均");
                        _table.Columns.Add("學業體育平均科排名");
                        _table.Columns.Add("學業體育平均科排名母數");
                        _table.Columns.Add("學業體育平均校排名");
                        _table.Columns.Add("學業體育平均校排名母數");
                        _table.Columns.Add("學業體育平均班排名");
                        _table.Columns.Add("學業體育平均班排名母數");
                        _table.Columns.Add("學業體育平均類別一排名");
                        _table.Columns.Add("學業體育平均類別一排名母數");
                        _table.Columns.Add("學業體育平均類別二排名");
                        _table.Columns.Add("學業體育平均類別二排名母數");
                        _table.Columns.Add("學業體育總分");
                        _table.Columns.Add("學業體育總分科排名");
                        _table.Columns.Add("學業體育總分科排名母數");
                        _table.Columns.Add("學業體育總分校排名");
                        _table.Columns.Add("學業體育總分校排名母數");
                        _table.Columns.Add("學業體育總分班排名");
                        _table.Columns.Add("學業體育總分班排名母數");
                        _table.Columns.Add("學業體育總分類別一排名");
                        _table.Columns.Add("學業體育總分類別一排名母數");
                        _table.Columns.Add("學業體育總分類別二排名");
                        _table.Columns.Add("學業體育總分類別二排名母數");
                        _table.Columns.Add("學業體育平均科排名百分比");
                        _table.Columns.Add("學業體育平均校排名百分比");
                        _table.Columns.Add("學業體育平均班排名百分比");
                        _table.Columns.Add("學業體育平均類別一排名百分比");
                        _table.Columns.Add("學業體育平均類別二排名百分比");
                        _table.Columns.Add("一上學業健康與護理成績");
                        _table.Columns.Add("一下學業健康與護理成績");
                        _table.Columns.Add("二上學業健康與護理成績");
                        _table.Columns.Add("二下學業健康與護理成績");
                        _table.Columns.Add("三上學業健康與護理成績");
                        _table.Columns.Add("三下學業健康與護理成績");
                        _table.Columns.Add("四上學業健康與護理成績");
                        _table.Columns.Add("四下學業健康與護理成績");
                        _table.Columns.Add("學業健康與護理平均");
                        _table.Columns.Add("學業健康與護理平均科排名");
                        _table.Columns.Add("學業健康與護理平均科排名母數");
                        _table.Columns.Add("學業健康與護理平均校排名");
                        _table.Columns.Add("學業健康與護理平均校排名母數");
                        _table.Columns.Add("學業健康與護理平均班排名");
                        _table.Columns.Add("學業健康與護理平均班排名母數");
                        _table.Columns.Add("學業健康與護理平均類別一排名");
                        _table.Columns.Add("學業健康與護理平均類別一排名母數");
                        _table.Columns.Add("學業健康與護理平均類別二排名");
                        _table.Columns.Add("學業健康與護理平均類別二排名母數");
                        _table.Columns.Add("學業健康與護理總分");
                        _table.Columns.Add("學業健康與護理總分科排名");
                        _table.Columns.Add("學業健康與護理總分科排名母數");
                        _table.Columns.Add("學業健康與護理總分校排名");
                        _table.Columns.Add("學業健康與護理總分校排名母數");
                        _table.Columns.Add("學業健康與護理總分班排名");
                        _table.Columns.Add("學業健康與護理總分班排名母數");
                        _table.Columns.Add("學業健康與護理總分類別一排名");
                        _table.Columns.Add("學業健康與護理總分類別一排名母數");
                        _table.Columns.Add("學業健康與護理總分類別二排名");
                        _table.Columns.Add("學業健康與護理總分類別二排名母數");
                        _table.Columns.Add("學業健康與護理平均科排名百分比");
                        _table.Columns.Add("學業健康與護理平均校排名百分比");
                        _table.Columns.Add("學業健康與護理平均班排名百分比");
                        _table.Columns.Add("學業健康與護理平均類別一排名百分比");
                        _table.Columns.Add("學業健康與護理平均類別二排名百分比");
                        _table.Columns.Add("一上學業國防通識成績");
                        _table.Columns.Add("一下學業國防通識成績");
                        _table.Columns.Add("二上學業國防通識成績");
                        _table.Columns.Add("二下學業國防通識成績");
                        _table.Columns.Add("三上學業國防通識成績");
                        _table.Columns.Add("三下學業國防通識成績");
                        _table.Columns.Add("四上學業國防通識成績");
                        _table.Columns.Add("四下學業國防通識成績");
                        _table.Columns.Add("學業國防通識平均");
                        _table.Columns.Add("學業國防通識平均科排名");
                        _table.Columns.Add("學業國防通識平均科排名母數");
                        _table.Columns.Add("學業國防通識平均校排名");
                        _table.Columns.Add("學業國防通識平均校排名母數");
                        _table.Columns.Add("學業國防通識平均班排名");
                        _table.Columns.Add("學業國防通識平均班排名母數");
                        _table.Columns.Add("學業國防通識平均類別一排名");
                        _table.Columns.Add("學業國防通識平均類別一排名母數");
                        _table.Columns.Add("學業國防通識平均類別二排名");
                        _table.Columns.Add("學業國防通識平均類別二排名母數");
                        _table.Columns.Add("學業國防通識總分");
                        _table.Columns.Add("學業國防通識總分科排名");
                        _table.Columns.Add("學業國防通識總分科排名母數");
                        _table.Columns.Add("學業國防通識總分校排名");
                        _table.Columns.Add("學業國防通識總分校排名母數");
                        _table.Columns.Add("學業國防通識總分班排名");
                        _table.Columns.Add("學業國防通識總分班排名母數");
                        _table.Columns.Add("學業國防通識總分類別一排名");
                        _table.Columns.Add("學業國防通識總分類別一排名母數");
                        _table.Columns.Add("學業國防通識總分類別二排名");
                        _table.Columns.Add("學業國防通識總分類別二排名母數");
                        _table.Columns.Add("學業國防通識平均科排名百分比");
                        _table.Columns.Add("學業國防通識平均校排名百分比");
                        _table.Columns.Add("學業國防通識平均班排名百分比");
                        _table.Columns.Add("學業國防通識平均類別一排名百分比");
                        _table.Columns.Add("學業國防通識平均類別二排名百分比");
                        _table.Columns.Add("一上學業實習科目成績");
                        _table.Columns.Add("一下學業實習科目成績");
                        _table.Columns.Add("二上學業實習科目成績");
                        _table.Columns.Add("二下學業實習科目成績");
                        _table.Columns.Add("三上學業實習科目成績");
                        _table.Columns.Add("三下學業實習科目成績");
                        _table.Columns.Add("四上學業實習科目成績");
                        _table.Columns.Add("四下學業實習科目成績");
                        _table.Columns.Add("學業實習科目平均");
                        _table.Columns.Add("學業實習科目平均科排名");
                        _table.Columns.Add("學業實習科目平均科排名母數");
                        _table.Columns.Add("學業實習科目平均校排名");
                        _table.Columns.Add("學業實習科目平均校排名母數");
                        _table.Columns.Add("學業實習科目平均班排名");
                        _table.Columns.Add("學業實習科目平均班排名母數");
                        _table.Columns.Add("學業實習科目平均類別一排名");
                        _table.Columns.Add("學業實習科目平均類別一排名母數");
                        _table.Columns.Add("學業實習科目平均類別二排名");
                        _table.Columns.Add("學業實習科目平均類別二排名母數");
                        _table.Columns.Add("學業實習科目總分");
                        _table.Columns.Add("學業實習科目總分科排名");
                        _table.Columns.Add("學業實習科目總分科排名母數");
                        _table.Columns.Add("學業實習科目總分校排名");
                        _table.Columns.Add("學業實習科目總分校排名母數");
                        _table.Columns.Add("學業實習科目總分班排名");
                        _table.Columns.Add("學業實習科目總分班排名母數");
                        _table.Columns.Add("學業實習科目總分類別一排名");
                        _table.Columns.Add("學業實習科目總分類別一排名母數");
                        _table.Columns.Add("學業實習科目總分類別二排名");
                        _table.Columns.Add("學業實習科目總分類別二排名母數");
                        _table.Columns.Add("學業實習科目平均科排名百分比");
                        _table.Columns.Add("學業實習科目平均校排名百分比");
                        _table.Columns.Add("學業實習科目平均班排名百分比");
                        _table.Columns.Add("學業實習科目平均類別一排名百分比");
                        _table.Columns.Add("學業實習科目平均類別二排名百分比");


                        _table.Columns.Add("總計加權平均");
                        _table.Columns.Add("總計加權平均科排名");
                        _table.Columns.Add("總計加權平均科排名母數");
                        _table.Columns.Add("總計加權平均校排名");
                        _table.Columns.Add("總計加權平均校排名母數");
                        _table.Columns.Add("總計加權平均班排名");
                        _table.Columns.Add("總計加權平均班排名母數");
                        _table.Columns.Add("總計加權平均類別一");
                        _table.Columns.Add("總計加權平均類別一排名");
                        _table.Columns.Add("總計加權平均類別一排名母數");
                        _table.Columns.Add("總計加權平均類別二");
                        _table.Columns.Add("總計加權平均類別二排名");
                        _table.Columns.Add("總計加權平均類別二排名母數");
                        _table.Columns.Add("總計加權總分");
                        _table.Columns.Add("總計加權總分科排名");
                        _table.Columns.Add("總計加權總分科排名母數");
                        _table.Columns.Add("總計加權總分校排名");
                        _table.Columns.Add("總計加權總分校排名母數");
                        _table.Columns.Add("總計加權總分班排名");
                        _table.Columns.Add("總計加權總分班排名母數");
                        _table.Columns.Add("總計加權總分類別一");
                        _table.Columns.Add("總計加權總分類別一排名");
                        _table.Columns.Add("總計加權總分類別一排名母數");
                        _table.Columns.Add("總計加權總分類別二");
                        _table.Columns.Add("總計加權總分類別二排名");
                        _table.Columns.Add("總計加權總分類別二排名母數");
                        _table.Columns.Add("總計平均");
                        _table.Columns.Add("總計平均科排名");
                        _table.Columns.Add("總計平均科排名母數");
                        _table.Columns.Add("總計平均校排名");
                        _table.Columns.Add("總計平均校排名母數");
                        _table.Columns.Add("總計平均班排名");
                        _table.Columns.Add("總計平均班排名母數");
                        _table.Columns.Add("總計平均類別一");
                        _table.Columns.Add("總計平均類別一排名");
                        _table.Columns.Add("總計平均類別一排名母數");
                        _table.Columns.Add("總計平均類別二");
                        _table.Columns.Add("總計平均類別二排名");
                        _table.Columns.Add("總計平均類別二排名母數");
                        _table.Columns.Add("總計總分");
                        _table.Columns.Add("總計總分科排名");
                        _table.Columns.Add("總計總分科排名母數");
                        _table.Columns.Add("總計總分校排名");
                        _table.Columns.Add("總計總分校排名母數");
                        _table.Columns.Add("總計總分班排名");
                        _table.Columns.Add("總計總分班排名母數");
                        _table.Columns.Add("總計總分類別一");
                        _table.Columns.Add("總計總分類別一排名");
                        _table.Columns.Add("總計總分類別一排名母數");
                        _table.Columns.Add("總計總分類別二");
                        _table.Columns.Add("總計總分類別二排名");
                        _table.Columns.Add("總計總分類別二排名母數");

                        for (int g = 1; g <= 4; g++)
                        {
                            string str = "";

                            for (int s = 1; s <= 2; s++)
                            {
                                if (g == 1 && s == 1) str = "一上";
                                if (g == 1 && s == 2) str = "一下";
                                if (g == 2 && s == 1) str = "二上";
                                if (g == 2 && s == 2) str = "二下";
                                if (g == 3 && s == 1) str = "三上";
                                if (g == 3 && s == 2) str = "三下";
                                if (g == 4 && s == 1) str = "四上";
                                if (g == 4 && s == 2) str = "四下";

                                _table.Columns.Add(str + "科目原始成績加權平均");
                                _table.Columns.Add(str + "科目原始成績加權平均班排名");
                                _table.Columns.Add(str + "科目原始成績加權平均班排名母數");
                                _table.Columns.Add(str + "科目原始成績加權平均班排名百分比");
                                _table.Columns.Add(str + "科目原始成績加權平均科排名");
                                _table.Columns.Add(str + "科目原始成績加權平均科排名母數");
                                _table.Columns.Add(str + "科目原始成績加權平均科排名百分比");
                                _table.Columns.Add(str + "科目原始成績加權平均校排名");
                                _table.Columns.Add(str + "科目原始成績加權平均校排名母數");
                                _table.Columns.Add(str + "科目原始成績加權平均校排名百分比");
                                _table.Columns.Add(str + "科目原始成績加權平均類別一");
                                _table.Columns.Add(str + "科目原始成績加權平均類別一排名");
                                _table.Columns.Add(str + "科目原始成績加權平均類別一排名母數");
                                _table.Columns.Add(str + "科目原始成績加權平均類別一排名百分比");
                                _table.Columns.Add(str + "科目原始成績加權平均類別二");
                                _table.Columns.Add(str + "科目原始成績加權平均類別二排名");
                                _table.Columns.Add(str + "科目原始成績加權平均類別二排名母數");
                                _table.Columns.Add(str + "科目原始成績加權平均類別二排名百分比");

                                _table.Columns.Add(str + "篩選科目原始成績加權平均");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均班排名");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均班排名母數");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均班排名百分比");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均科排名");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均科排名母數");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均科排名百分比");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均校排名");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均校排名母數");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均校排名百分比");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均類別一");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均類別一排名");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均類別一排名母數");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均類別一排名百分比");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均類別二");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均類別二排名");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均類別二排名母數");
                                _table.Columns.Add(str + "篩選科目原始成績加權平均類別二排名百分比");
                            }
                        }

                        //  
                        _table.Columns.Add("科目原始成績加權平均平均");
                        _table.Columns.Add("科目原始成績加權平均平均班排名");
                        _table.Columns.Add("科目原始成績加權平均平均班排名母數");
                        _table.Columns.Add("科目原始成績加權平均平均班排名百分比");
                        _table.Columns.Add("科目原始成績加權平均平均科排名");
                        _table.Columns.Add("科目原始成績加權平均平均科排名母數");
                        _table.Columns.Add("科目原始成績加權平均平均科排名百分比");
                        _table.Columns.Add("科目原始成績加權平均平均校排名");
                        _table.Columns.Add("科目原始成績加權平均平均校排名母數");
                        _table.Columns.Add("科目原始成績加權平均平均校排名百分比");
                        _table.Columns.Add("科目原始成績加權平均平均類別一");
                        _table.Columns.Add("科目原始成績加權平均平均類別一排名");
                        _table.Columns.Add("科目原始成績加權平均平均類別一排名母數");
                        _table.Columns.Add("科目原始成績加權平均平均類別一排名百分比");
                        _table.Columns.Add("科目原始成績加權平均平均類別二");
                        _table.Columns.Add("科目原始成績加權平均平均類別二排名");
                        _table.Columns.Add("科目原始成績加權平均平均類別二排名母數");
                        _table.Columns.Add("科目原始成績加權平均平均類別二排名百分比");

                        _table.Columns.Add("篩選科目原始成績加權平均平均");
                        _table.Columns.Add("篩選科目原始成績加權平均平均班排名");
                        _table.Columns.Add("篩選科目原始成績加權平均平均班排名母數");
                        _table.Columns.Add("篩選科目原始成績加權平均平均班排名百分比");
                        _table.Columns.Add("篩選科目原始成績加權平均平均科排名");
                        _table.Columns.Add("篩選科目原始成績加權平均平均科排名母數");
                        _table.Columns.Add("篩選科目原始成績加權平均平均科排名百分比");
                        _table.Columns.Add("篩選科目原始成績加權平均平均校排名");
                        _table.Columns.Add("篩選科目原始成績加權平均平均校排名母數");
                        _table.Columns.Add("篩選科目原始成績加權平均平均校排名百分比");
                        _table.Columns.Add("篩選科目原始成績加權平均平均類別一");
                        _table.Columns.Add("篩選科目原始成績加權平均平均類別一排名");
                        _table.Columns.Add("篩選科目原始成績加權平均平均類別一排名母數");
                        _table.Columns.Add("篩選科目原始成績加權平均平均類別一排名百分比");
                        _table.Columns.Add("篩選科目原始成績加權平均平均類別二");
                        _table.Columns.Add("篩選科目原始成績加權平均平均類別二排名");
                        _table.Columns.Add("篩選科目原始成績加權平均平均類別二排名母數");
                        _table.Columns.Add("篩選科目原始成績加權平均平均類別二排名百分比");

                        // 動態科目類
                        int subjCot = setting.useSubjectPrintList.Count;
                        if (setting.SubjectLimit > subjCot)
                            subjCot = setting.SubjectLimit;

                        for (int i = 1; i <= subjCot; i++)
                        {
                            _table.Columns.Add("科目名稱" + i);
                            _table.Columns.Add("一上科目成績" + i);
                            _table.Columns.Add("一上科目學分數" + i);
                            _table.Columns.Add("一上科目排名" + i);
                            _table.Columns.Add("一下科目成績" + i);
                            _table.Columns.Add("一下科目學分數" + i);
                            _table.Columns.Add("一下科目排名" + i);
                            _table.Columns.Add("二上科目成績" + i);
                            _table.Columns.Add("二上科目學分數" + i);
                            _table.Columns.Add("二上科目排名" + i);
                            _table.Columns.Add("二下科目成績" + i);
                            _table.Columns.Add("二下科目學分數" + i);
                            _table.Columns.Add("二下科目排名" + i);
                            _table.Columns.Add("三上科目成績" + i);
                            _table.Columns.Add("三上科目學分數" + i);
                            _table.Columns.Add("三上科目排名" + i);
                            _table.Columns.Add("三下科目成績" + i);
                            _table.Columns.Add("三下科目學分數" + i);
                            _table.Columns.Add("三下科目排名" + i);
                            _table.Columns.Add("四上科目成績" + i);
                            _table.Columns.Add("四上科目學分數" + i);
                            _table.Columns.Add("四上科目排名" + i);
                            _table.Columns.Add("四下科目成績" + i);
                            _table.Columns.Add("四下科目學分數" + i);
                            _table.Columns.Add("四下科目排名" + i);
                            _table.Columns.Add("科目加權平均" + i);
                            _table.Columns.Add("科目加權平均科排名" + i);
                            _table.Columns.Add("科目加權平均科排名母數" + i);
                            _table.Columns.Add("科目加權平均校排名" + i);
                            _table.Columns.Add("科目加權平均校排名母數" + i);
                            _table.Columns.Add("科目加權平均班排名" + i);
                            _table.Columns.Add("科目加權平均班排名母數" + i);
                            _table.Columns.Add("科目加權平均類別一排名" + i);
                            _table.Columns.Add("科目加權平均類別一排名母數" + i);
                            _table.Columns.Add("科目加權平均類別二排名" + i);
                            _table.Columns.Add("科目加權平均類別二排名母數" + i);
                            _table.Columns.Add("科目加權總分" + i);
                            _table.Columns.Add("科目加權總分科排名" + i);
                            _table.Columns.Add("科目加權總分科排名母數" + i);
                            _table.Columns.Add("科目加權總分校排名" + i);
                            _table.Columns.Add("科目加權總分校排名母數" + i);
                            _table.Columns.Add("科目加權總分班排名" + i);
                            _table.Columns.Add("科目加權總分班排名母數" + i);
                            _table.Columns.Add("科目加權總分類別一排名" + i);
                            _table.Columns.Add("科目加權總分類別一排名母數" + i);
                            _table.Columns.Add("科目加權總分類別二排名" + i);
                            _table.Columns.Add("科目加權總分類別二排名母數" + i);
                            _table.Columns.Add("科目平均" + i);
                            _table.Columns.Add("科目平均科排名" + i);
                            _table.Columns.Add("科目平均科排名母數" + i);
                            _table.Columns.Add("科目平均校排名" + i);
                            _table.Columns.Add("科目平均校排名母數" + i);
                            _table.Columns.Add("科目平均班排名" + i);
                            _table.Columns.Add("科目平均班排名母數" + i);
                            _table.Columns.Add("科目平均類別一排名" + i);
                            _table.Columns.Add("科目平均類別一排名母數" + i);
                            _table.Columns.Add("科目平均類別二排名" + i);
                            _table.Columns.Add("科目平均類別二排名母數" + i);
                            _table.Columns.Add("科目總分" + i);
                            _table.Columns.Add("科目總分科排名" + i);
                            _table.Columns.Add("科目總分科排名母數" + i);
                            _table.Columns.Add("科目總分校排名" + i);
                            _table.Columns.Add("科目總分校排名母數" + i);
                            _table.Columns.Add("科目總分班排名" + i);
                            _table.Columns.Add("科目總分班排名母數" + i);
                            _table.Columns.Add("科目總分類別一排名" + i);
                            _table.Columns.Add("科目總分類別一排名母數" + i);
                            _table.Columns.Add("科目總分類別二排名" + i);
                            _table.Columns.Add("科目總分類別二排名母數" + i);

                            _table.Columns.Add("科目平均科排名百分比" + i);
                            _table.Columns.Add("科目平均校排名百分比" + i);
                            _table.Columns.Add("科目平均班排名百分比" + i);
                            _table.Columns.Add("科目總分科排名百分比" + i);
                            _table.Columns.Add("科目總分校排名百分比" + i);
                            _table.Columns.Add("科目總分班排名百分比" + i);
                            _table.Columns.Add("科目加權平均科排名百分比" + i);
                            _table.Columns.Add("科目加權平均校排名百分比" + i);
                            _table.Columns.Add("科目加權平均班排名百分比" + i);
                            _table.Columns.Add("科目加權總分科排名百分比" + i);
                            _table.Columns.Add("科目加權總分校排名百分比" + i);
                            _table.Columns.Add("科目加權總分班排名百分比" + i);
                            _table.Columns.Add("科目平均類別一排名百分比" + i);
                            _table.Columns.Add("科目平均類別二排名百分比" + i);
                            _table.Columns.Add("科目總分類別一排名百分比" + i);
                            _table.Columns.Add("科目總分類別二排名百分比" + i);
                            _table.Columns.Add("科目加權平均類別一排名百分比" + i);
                            _table.Columns.Add("科目加權平均類別二排名百分比" + i);
                            _table.Columns.Add("科目加權總分類別一排名百分比" + i);
                            _table.Columns.Add("科目加權總分類別二排名百分比" + i);

                            _table.Columns.Add("一上科目班排名" + i);
                            _table.Columns.Add("一上科目班排名母數" + i);
                            _table.Columns.Add("一上科目班排名百分比" + i);
                            _table.Columns.Add("一上科目科排名" + i);
                            _table.Columns.Add("一上科目科排名母數" + i);
                            _table.Columns.Add("一上科目科排名百分比" + i);
                            _table.Columns.Add("一上科目校排名" + i);
                            _table.Columns.Add("一上科目校排名母數" + i);
                            _table.Columns.Add("一上科目校排名百分比" + i);
                            _table.Columns.Add("一上科目類1排名" + i);
                            _table.Columns.Add("一上科目類1排名母數" + i);
                            _table.Columns.Add("一上科目類1排名百分比" + i);
                            _table.Columns.Add("一下科目班排名" + i);
                            _table.Columns.Add("一下科目班排名母數" + i);
                            _table.Columns.Add("一下科目班排名百分比" + i);
                            _table.Columns.Add("一下科目科排名" + i);
                            _table.Columns.Add("一下科目科排名母數" + i);
                            _table.Columns.Add("一下科目科排名百分比" + i);
                            _table.Columns.Add("一下科目校排名" + i);
                            _table.Columns.Add("一下科目校排名母數" + i);
                            _table.Columns.Add("一下科目校排名百分比" + i);
                            _table.Columns.Add("一下科目類1排名" + i);
                            _table.Columns.Add("一下科目類1排名母數" + i);
                            _table.Columns.Add("一下科目類1排名百分比" + i);
                            _table.Columns.Add("二上科目班排名" + i);
                            _table.Columns.Add("二上科目班排名母數" + i);
                            _table.Columns.Add("二上科目班排名百分比" + i);
                            _table.Columns.Add("二上科目科排名" + i);
                            _table.Columns.Add("二上科目科排名母數" + i);
                            _table.Columns.Add("二上科目科排名百分比" + i);
                            _table.Columns.Add("二上科目校排名" + i);
                            _table.Columns.Add("二上科目校排名母數" + i);
                            _table.Columns.Add("二上科目校排名百分比" + i);
                            _table.Columns.Add("二上科目類1排名" + i);
                            _table.Columns.Add("二上科目類1排名母數" + i);
                            _table.Columns.Add("二上科目類1排名百分比" + i);
                            _table.Columns.Add("二下科目班排名" + i);
                            _table.Columns.Add("二下科目班排名母數" + i);
                            _table.Columns.Add("二下科目班排名百分比" + i);
                            _table.Columns.Add("二下科目科排名" + i);
                            _table.Columns.Add("二下科目科排名母數" + i);
                            _table.Columns.Add("二下科目科排名百分比" + i);
                            _table.Columns.Add("二下科目校排名" + i);
                            _table.Columns.Add("二下科目校排名母數" + i);
                            _table.Columns.Add("二下科目校排名百分比" + i);
                            _table.Columns.Add("二下科目類1排名" + i);
                            _table.Columns.Add("二下科目類1排名母數" + i);
                            _table.Columns.Add("二下科目類1排名百分比" + i);
                            _table.Columns.Add("三上科目班排名" + i);
                            _table.Columns.Add("三上科目班排名母數" + i);
                            _table.Columns.Add("三上科目班排名百分比" + i);
                            _table.Columns.Add("三上科目科排名" + i);
                            _table.Columns.Add("三上科目科排名母數" + i);
                            _table.Columns.Add("三上科目科排名百分比" + i);
                            _table.Columns.Add("三上科目校排名" + i);
                            _table.Columns.Add("三上科目校排名母數" + i);
                            _table.Columns.Add("三上科目校排名百分比" + i);
                            _table.Columns.Add("三上科目類1排名" + i);
                            _table.Columns.Add("三上科目類1排名母數" + i);
                            _table.Columns.Add("三上科目類1排名百分比" + i);
                            _table.Columns.Add("三下科目班排名" + i);
                            _table.Columns.Add("三下科目班排名母數" + i);
                            _table.Columns.Add("三下科目班排名百分比" + i);
                            _table.Columns.Add("三下科目科排名" + i);
                            _table.Columns.Add("三下科目科排名母數" + i);
                            _table.Columns.Add("三下科目科排名百分比" + i);
                            _table.Columns.Add("三下科目校排名" + i);
                            _table.Columns.Add("三下科目校排名母數" + i);
                            _table.Columns.Add("三下科目校排名百分比" + i);
                            _table.Columns.Add("三下科目類1排名" + i);
                            _table.Columns.Add("三下科目類1排名母數" + i);
                            _table.Columns.Add("三下科目類1排名百分比" + i);
                            _table.Columns.Add("四上科目班排名" + i);
                            _table.Columns.Add("四上科目班排名母數" + i);
                            _table.Columns.Add("四上科目班排名百分比" + i);
                            _table.Columns.Add("四上科目科排名" + i);
                            _table.Columns.Add("四上科目科排名母數" + i);
                            _table.Columns.Add("四上科目科排名百分比" + i);
                            _table.Columns.Add("四上科目校排名" + i);
                            _table.Columns.Add("四上科目校排名母數" + i);
                            _table.Columns.Add("四上科目校排名百分比" + i);
                            _table.Columns.Add("四上科目類1排名" + i);
                            _table.Columns.Add("四上科目類1排名母數" + i);
                            _table.Columns.Add("四上科目類1排名百分比" + i);
                            _table.Columns.Add("四下科目班排名" + i);
                            _table.Columns.Add("四下科目班排名母數" + i);
                            _table.Columns.Add("四下科目班排名百分比" + i);
                            _table.Columns.Add("四下科目科排名" + i);
                            _table.Columns.Add("四下科目科排名母數" + i);
                            _table.Columns.Add("四下科目科排名百分比" + i);
                            _table.Columns.Add("四下科目校排名" + i);
                            _table.Columns.Add("四下科目校排名母數" + i);
                            _table.Columns.Add("四下科目校排名百分比" + i);
                            _table.Columns.Add("四下科目類1排名" + i);
                            _table.Columns.Add("四下科目類1排名母數" + i);
                            _table.Columns.Add("四下科目類1排名百分比" + i);

                        }

                        for (int i = 1; i <= 60; i++)
                        {
                            _table.Columns.Add("回歸科目名稱" + i);
                            _table.Columns.Add("回歸一上科目成績" + i);
                            _table.Columns.Add("回歸一上科目學分數" + i);
                            _table.Columns.Add("回歸一上科目排名" + i);
                            _table.Columns.Add("回歸一下科目成績" + i);
                            _table.Columns.Add("回歸一下科目學分數" + i);
                            _table.Columns.Add("回歸一下科目排名" + i);
                            _table.Columns.Add("回歸二上科目成績" + i);
                            _table.Columns.Add("回歸二上科目學分數" + i);
                            _table.Columns.Add("回歸二上科目排名" + i);
                            _table.Columns.Add("回歸二下科目成績" + i);
                            _table.Columns.Add("回歸二下科目學分數" + i);
                            _table.Columns.Add("回歸二下科目排名" + i);
                            _table.Columns.Add("回歸三上科目成績" + i);
                            _table.Columns.Add("回歸三上科目學分數" + i);
                            _table.Columns.Add("回歸三上科目排名" + i);
                            _table.Columns.Add("回歸三下科目成績" + i);
                            _table.Columns.Add("回歸三下科目學分數" + i);
                            _table.Columns.Add("回歸三下科目排名" + i);
                            _table.Columns.Add("回歸四上科目成績" + i);
                            _table.Columns.Add("回歸四上科目學分數" + i);
                            _table.Columns.Add("回歸四上科目排名" + i);
                            _table.Columns.Add("回歸四下科目成績" + i);
                            _table.Columns.Add("回歸四下科目學分數" + i);
                            _table.Columns.Add("回歸四下科目排名" + i);
                            _table.Columns.Add("回歸科目加權平均" + i);
                            _table.Columns.Add("回歸科目加權平均科排名" + i);
                            _table.Columns.Add("回歸科目加權平均科排名母數" + i);
                            _table.Columns.Add("回歸科目加權平均校排名" + i);
                            _table.Columns.Add("回歸科目加權平均校排名母數" + i);
                            _table.Columns.Add("回歸科目加權平均班排名" + i);
                            _table.Columns.Add("回歸科目加權平均班排名母數" + i);
                            _table.Columns.Add("回歸科目加權平均類別一排名" + i);
                            _table.Columns.Add("回歸科目加權平均類別一排名母數" + i);
                            _table.Columns.Add("回歸科目加權平均類別二排名" + i);
                            _table.Columns.Add("回歸科目加權平均類別二排名母數" + i);
                            _table.Columns.Add("回歸科目加權總分" + i);
                            _table.Columns.Add("回歸科目加權總分科排名" + i);
                            _table.Columns.Add("回歸科目加權總分科排名母數" + i);
                            _table.Columns.Add("回歸科目加權總分校排名" + i);
                            _table.Columns.Add("回歸科目加權總分校排名母數" + i);
                            _table.Columns.Add("回歸科目加權總分班排名" + i);
                            _table.Columns.Add("回歸科目加權總分班排名母數" + i);
                            _table.Columns.Add("回歸科目加權總分類別一排名" + i);
                            _table.Columns.Add("回歸科目加權總分類別一排名母數" + i);
                            _table.Columns.Add("回歸科目加權總分類別二排名" + i);
                            _table.Columns.Add("回歸科目加權總分類別二排名母數" + i);
                            _table.Columns.Add("回歸科目平均" + i);
                            _table.Columns.Add("回歸科目平均科排名" + i);
                            _table.Columns.Add("回歸科目平均科排名母數" + i);
                            _table.Columns.Add("回歸科目平均校排名" + i);
                            _table.Columns.Add("回歸科目平均校排名母數" + i);
                            _table.Columns.Add("回歸科目平均班排名" + i);
                            _table.Columns.Add("回歸科目平均班排名母數" + i);
                            _table.Columns.Add("回歸科目平均類別一排名" + i);
                            _table.Columns.Add("回歸科目平均類別一排名母數" + i);
                            _table.Columns.Add("回歸科目平均類別二排名" + i);
                            _table.Columns.Add("回歸科目平均類別二排名母數" + i);
                            _table.Columns.Add("回歸科目總分" + i);
                            _table.Columns.Add("回歸科目總分科排名" + i);
                            _table.Columns.Add("回歸科目總分科排名母數" + i);
                            _table.Columns.Add("回歸科目總分校排名" + i);
                            _table.Columns.Add("回歸科目總分校排名母數" + i);
                            _table.Columns.Add("回歸科目總分班排名" + i);
                            _table.Columns.Add("回歸科目總分班排名母數" + i);
                            _table.Columns.Add("回歸科目總分類別一排名" + i);
                            _table.Columns.Add("回歸科目總分類別一排名母數" + i);
                            _table.Columns.Add("回歸科目總分類別二排名" + i);
                            _table.Columns.Add("回歸科目總分類別二排名母數" + i);

                            _table.Columns.Add("回歸科目平均科排名百分比" + i);
                            _table.Columns.Add("回歸科目平均校排名百分比" + i);
                            _table.Columns.Add("回歸科目平均班排名百分比" + i);
                            _table.Columns.Add("回歸科目總分科排名百分比" + i);
                            _table.Columns.Add("回歸科目總分校排名百分比" + i);
                            _table.Columns.Add("回歸科目總分班排名百分比" + i);
                            _table.Columns.Add("回歸科目加權平均科排名百分比" + i);
                            _table.Columns.Add("回歸科目加權平均校排名百分比" + i);
                            _table.Columns.Add("回歸科目加權平均班排名百分比" + i);
                            _table.Columns.Add("回歸科目加權總分科排名百分比" + i);
                            _table.Columns.Add("回歸科目加權總分校排名百分比" + i);
                            _table.Columns.Add("回歸科目加權總分班排名百分比" + i);
                            _table.Columns.Add("回歸科目平均類別一排名百分比" + i);
                            _table.Columns.Add("回歸科目平均類別二排名百分比" + i);
                            _table.Columns.Add("回歸科目總分類別一排名百分比" + i);
                            _table.Columns.Add("回歸科目總分類別二排名百分比" + i);
                            _table.Columns.Add("回歸科目加權平均類別一排名百分比" + i);
                            _table.Columns.Add("回歸科目加權平均類別二排名百分比" + i);
                            _table.Columns.Add("回歸科目加權總分類別一排名百分比" + i);
                            _table.Columns.Add("回歸科目加權總分類別二排名百分比" + i);

                            _table.Columns.Add("回歸一上科目班排名" + i);
                            _table.Columns.Add("回歸一上科目班排名母數" + i);
                            _table.Columns.Add("回歸一上科目班排名百分比" + i);
                            _table.Columns.Add("回歸一上科目科排名" + i);
                            _table.Columns.Add("回歸一上科目科排名母數" + i);
                            _table.Columns.Add("回歸一上科目科排名百分比" + i);
                            _table.Columns.Add("回歸一上科目校排名" + i);
                            _table.Columns.Add("回歸一上科目校排名母數" + i);
                            _table.Columns.Add("回歸一上科目校排名百分比" + i);
                            _table.Columns.Add("回歸一上科目類1排名" + i);
                            _table.Columns.Add("回歸一上科目類1排名母數" + i);
                            _table.Columns.Add("回歸一上科目類1排名百分比" + i);
                            _table.Columns.Add("回歸一下科目班排名" + i);
                            _table.Columns.Add("回歸一下科目班排名母數" + i);
                            _table.Columns.Add("回歸一下科目班排名百分比" + i);
                            _table.Columns.Add("回歸一下科目科排名" + i);
                            _table.Columns.Add("回歸一下科目科排名母數" + i);
                            _table.Columns.Add("回歸一下科目科排名百分比" + i);
                            _table.Columns.Add("回歸一下科目校排名" + i);
                            _table.Columns.Add("回歸一下科目校排名母數" + i);
                            _table.Columns.Add("回歸一下科目校排名百分比" + i);
                            _table.Columns.Add("回歸一下科目類1排名" + i);
                            _table.Columns.Add("回歸一下科目類1排名母數" + i);
                            _table.Columns.Add("回歸一下科目類1排名百分比" + i);
                            _table.Columns.Add("回歸二上科目班排名" + i);
                            _table.Columns.Add("回歸二上科目班排名母數" + i);
                            _table.Columns.Add("回歸二上科目班排名百分比" + i);
                            _table.Columns.Add("回歸二上科目科排名" + i);
                            _table.Columns.Add("回歸二上科目科排名母數" + i);
                            _table.Columns.Add("回歸二上科目科排名百分比" + i);
                            _table.Columns.Add("回歸二上科目校排名" + i);
                            _table.Columns.Add("回歸二上科目校排名母數" + i);
                            _table.Columns.Add("回歸二上科目校排名百分比" + i);
                            _table.Columns.Add("回歸二上科目類1排名" + i);
                            _table.Columns.Add("回歸二上科目類1排名母數" + i);
                            _table.Columns.Add("回歸二上科目類1排名百分比" + i);
                            _table.Columns.Add("回歸二下科目班排名" + i);
                            _table.Columns.Add("回歸二下科目班排名母數" + i);
                            _table.Columns.Add("回歸二下科目班排名百分比" + i);
                            _table.Columns.Add("回歸二下科目科排名" + i);
                            _table.Columns.Add("回歸二下科目科排名母數" + i);
                            _table.Columns.Add("回歸二下科目科排名百分比" + i);
                            _table.Columns.Add("回歸二下科目校排名" + i);
                            _table.Columns.Add("回歸二下科目校排名母數" + i);
                            _table.Columns.Add("回歸二下科目校排名百分比" + i);
                            _table.Columns.Add("回歸二下科目類1排名" + i);
                            _table.Columns.Add("回歸二下科目類1排名母數" + i);
                            _table.Columns.Add("回歸二下科目類1排名百分比" + i);
                            _table.Columns.Add("回歸三上科目班排名" + i);
                            _table.Columns.Add("回歸三上科目班排名母數" + i);
                            _table.Columns.Add("回歸三上科目班排名百分比" + i);
                            _table.Columns.Add("回歸三上科目科排名" + i);
                            _table.Columns.Add("回歸三上科目科排名母數" + i);
                            _table.Columns.Add("回歸三上科目科排名百分比" + i);
                            _table.Columns.Add("回歸三上科目校排名" + i);
                            _table.Columns.Add("回歸三上科目校排名母數" + i);
                            _table.Columns.Add("回歸三上科目校排名百分比" + i);
                            _table.Columns.Add("回歸三上科目類1排名" + i);
                            _table.Columns.Add("回歸三上科目類1排名母數" + i);
                            _table.Columns.Add("回歸三上科目類1排名百分比" + i);
                            _table.Columns.Add("回歸三下科目班排名" + i);
                            _table.Columns.Add("回歸三下科目班排名母數" + i);
                            _table.Columns.Add("回歸三下科目班排名百分比" + i);
                            _table.Columns.Add("回歸三下科目科排名" + i);
                            _table.Columns.Add("回歸三下科目科排名母數" + i);
                            _table.Columns.Add("回歸三下科目科排名百分比" + i);
                            _table.Columns.Add("回歸三下科目校排名" + i);
                            _table.Columns.Add("回歸三下科目校排名母數" + i);
                            _table.Columns.Add("回歸三下科目校排名百分比" + i);
                            _table.Columns.Add("回歸三下科目類1排名" + i);
                            _table.Columns.Add("回歸三下科目類1排名母數" + i);
                            _table.Columns.Add("回歸三下科目類1排名百分比" + i);
                            _table.Columns.Add("回歸四上科目班排名" + i);
                            _table.Columns.Add("回歸四上科目班排名母數" + i);
                            _table.Columns.Add("回歸四上科目班排名百分比" + i);
                            _table.Columns.Add("回歸四上科目科排名" + i);
                            _table.Columns.Add("回歸四上科目科排名母數" + i);
                            _table.Columns.Add("回歸四上科目科排名百分比" + i);
                            _table.Columns.Add("回歸四上科目校排名" + i);
                            _table.Columns.Add("回歸四上科目校排名母數" + i);
                            _table.Columns.Add("回歸四上科目校排名百分比" + i);
                            _table.Columns.Add("回歸四上科目類1排名" + i);
                            _table.Columns.Add("回歸四上科目類1排名母數" + i);
                            _table.Columns.Add("回歸四上科目類1排名百分比" + i);
                            _table.Columns.Add("回歸四下科目班排名" + i);
                            _table.Columns.Add("回歸四下科目班排名母數" + i);
                            _table.Columns.Add("回歸四下科目班排名百分比" + i);
                            _table.Columns.Add("回歸四下科目科排名" + i);
                            _table.Columns.Add("回歸四下科目科排名母數" + i);
                            _table.Columns.Add("回歸四下科目科排名百分比" + i);
                            _table.Columns.Add("回歸四下科目校排名" + i);
                            _table.Columns.Add("回歸四下科目校排名母數" + i);
                            _table.Columns.Add("回歸四下科目校排名百分比" + i);
                            _table.Columns.Add("回歸四下科目類1排名" + i);
                            _table.Columns.Add("回歸四下科目類1排名母數" + i);
                            _table.Columns.Add("回歸四下科目類1排名百分比" + i);

                            _table.Columns.Add("回歸一上科目類2排名" + i);
                            _table.Columns.Add("回歸一上科目類2排名母數" + i);
                            _table.Columns.Add("回歸一上科目類2排名百分比" + i);

                            _table.Columns.Add("回歸一下科目類2排名" + i);
                            _table.Columns.Add("回歸一下科目類2排名母數" + i);
                            _table.Columns.Add("回歸一下科目類2排名百分比" + i);

                            _table.Columns.Add("回歸二上科目類2排名" + i);
                            _table.Columns.Add("回歸二上科目類2排名母數" + i);
                            _table.Columns.Add("回歸二上科目類2排名百分比" + i);

                            _table.Columns.Add("回歸二下科目類2排名" + i);
                            _table.Columns.Add("回歸二下科目類2排名母數" + i);
                            _table.Columns.Add("回歸二下科目類2排名百分比" + i);

                            _table.Columns.Add("回歸三上科目類2排名" + i);
                            _table.Columns.Add("回歸三上科目類2排名母數" + i);
                            _table.Columns.Add("回歸三上科目類2排名百分比" + i);

                            _table.Columns.Add("回歸三下科目類2排名" + i);
                            _table.Columns.Add("回歸三下科目類2排名母數" + i);
                            _table.Columns.Add("回歸三下科目類2排名百分比" + i);

                            _table.Columns.Add("回歸四上科目類2排名" + i);
                            _table.Columns.Add("回歸四上科目類2排名母數" + i);
                            _table.Columns.Add("回歸四上科目類2排名百分比" + i);

                            _table.Columns.Add("回歸四下科目類2排名" + i);
                            _table.Columns.Add("回歸四下科目類2排名母數" + i);
                            _table.Columns.Add("回歸四下科目類2排名百分比" + i);
                        }



                        #endregion 合併欄位使用



                        bkw.ReportProgress(80);


                        StringBuilder sbPDFErr = new StringBuilder();

                        // 取得學生ID
                        List<string> StudIDList = new List<string>();
                        foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                            StudIDList.Add(studRec.StudentID);

                        // 讀取各學業排名
                        Dictionary<string, List<StudSemsEntryRating>> StudSemsEntryRatingDict = Utility.GetStudSemsEntryRatingByStudentID(StudIDList);

                        // 讀取各學期各科目排名， 2018/3/31 穎驊註解，若一口氣查詢的學生筆數太多(ex:新民高中一屆1640人)此方法有可能造成資料庫記憶體爆掉，而回傳不回結果，可以進去看方法，我有最新改寫，一次500人批次處理。
                        Dictionary<string, List<StudSemsSubjRating>> StudSemsSubjRatingDict = Utility.GetStudSemsSubjRatingByStudentID(StudIDList);

                        // 判斷是否產生 PDF
                        if (setting.CheckExportPDF)
                        {
                            // 取得 serNo
                            Dictionary<string, string> StudSATSerNoDict = Utility.GetStudentSATSerNoByStudentIDList(StudIDList);
                            int StudCountStart = 0, StudSumCount = gradeyearStudents[gradeyear].Count;

                            #region 產生 PDF 檔案
                            foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                            {
                                //2018/3/31 穎驊註解，先在每一輪loop 開頭 就先濾掉不排名學生，可以避免資源浪費，以及
                                // 可能會造成#6517 DataRow row = _table.NewRow(); 記憶體洩漏的問題，
                                // 因為在一個DataRow row 產生後，其來源table 會預期它會被加入，
                                //但若後來造印製表才判斷此份資料不需被加入，就會產生記憶體洩漏，資料比數一大，就會造成系統當機。
                                if (noRankList.Contains(studRec.StudentID))
                                    continue;

                                string FileKey = "_" + studRec.StudentID;
                                string ErrMsg = "";

                                // 判斷存檔方式,Y 身分證號
                                if (setting.CheckUseIDNumber)
                                {
                                    if (!string.IsNullOrWhiteSpace(studRec.IDNumber))
                                        FileKey = studRec.IDNumber;
                                    else
                                    {
                                        ErrMsg = "檔案編號：_" + studRec.StudentID + ",學號：" + studRec.StudentNumber + ",學生姓名：" + studRec.StudentName + ", 沒有身分證號。";
                                    }
                                }
                                else
                                {
                                    if (StudSATSerNoDict.ContainsKey(studRec.StudentID))
                                    {
                                        // 有值
                                        if (!string.IsNullOrWhiteSpace(StudSATSerNoDict[studRec.StudentID]))
                                        {
                                            FileKey = StudSATSerNoDict[studRec.StudentID];
                                        }
                                        else
                                        {
                                            ErrMsg = "檔案編號：_" + studRec.StudentID + ",學號：" + studRec.StudentNumber + ", 學生姓名：" + studRec.StudentName + ", 沒有報名序號。";
                                        }
                                    }
                                    else
                                    {
                                        ErrMsg = "檔案編號：_" + studRec.StudentID + ",學號：" + studRec.StudentNumber + " ,學生姓名：" + studRec.StudentName + ", 沒有報名序號。";
                                    }
                                }

                                if (ErrMsg != "")
                                    sbPDFErr.AppendLine(ErrMsg);

                                // 產生 DataRow 資料
                                List<int> g1List = new List<int>();
                                List<int> g2List = new List<int>();
                                List<int> g3List = new List<int>();
                                List<int> g4List = new List<int>();

                                // Word 套表
                                // 樣板
                                Aspose.Words.Document doc = new Aspose.Words.Document();
                                Aspose.Words.Document docTemplate = setting.Template;
                                if (docTemplate == null)
                                    docTemplate = new Aspose.Words.Document(new MemoryStream(Properties.Resources.高中多學期學生成績證明書));

                                doc.Sections.Clear();
                                //DataTable _d = new DataTable();
                                //_d = _table.Clone();
                                DataRow row = _table.NewRow();
                                row["學校名稱"] = SchoolName;
                                row["班級"] = studRec.RefClass.ClassName;
                                row["座號"] = studRec.SeatNo;
                                row["學號"] = studRec.StudentNumber;
                                row["學生系統編號"] = studRec.StudentID;
                                row["教師系統編號"] = "";
                                row["教師姓名"] = "";
                                if (studRec.RefClass.RefTeacher != null)
                                {
                                    row["教師系統編號"] = studRec.RefClass.RefTeacher.TeacherID;
                                    row["教師姓名"] = studRec.RefClass.RefTeacher.TeacherName;
                                }
                                row["姓名"] = studRec.StudentName;
                                row["科別"] = studRec.Department;
                                row["類別一分類"] = (cat1Dict.ContainsKey(studRec.StudentID)) ? cat1Dict[studRec.StudentID] : "";
                                row["類別二分類"] = (cat2Dict.ContainsKey(studRec.StudentID)) ? cat2Dict[studRec.StudentID] : "";
                                row["類別排名1"] = (cat1Dict.ContainsKey(studRec.StudentID)) ? cat1Dict[studRec.StudentID] : "";
                                row["類別排名2"] = (cat2Dict.ContainsKey(studRec.StudentID)) ? cat2Dict[studRec.StudentID] : "";
                                row["一年級學年度"] = "";
                                row["二年級學年度"] = "";
                                row["三年級學年度"] = "";
                                row["四年級學年度"] = "";
                                g1List.Clear();
                                g2List.Clear();
                                g3List.Clear();
                                g4List.Clear();

                                int subjIndex = 1;
                                // 處理科目
                                #region 處理科目
                                foreach (string subjName in setting.useSubjectPrintList)
                                {
                                    string id = studRec.StudentID + "^^^" + subjName;

                                    if (selectScore.ContainsKey(id))
                                    {
                                        row["科目名稱" + subjIndex] = subjName;

                                        if (selectScore[id].gsSchoolYear11.HasValue)
                                        {
                                            if (!g1List.Contains(selectScore[id].gsSchoolYear11.Value))
                                                g1List.Add(selectScore[id].gsSchoolYear11.Value);
                                        }
                                        if (selectScore[id].gsSchoolYear12.HasValue)
                                        {
                                            if (!g1List.Contains(selectScore[id].gsSchoolYear12.Value))
                                                g1List.Add(selectScore[id].gsSchoolYear12.Value);
                                        }

                                        if (selectScore[id].gsSchoolYear21.HasValue)
                                        {
                                            if (!g2List.Contains(selectScore[id].gsSchoolYear21.Value))
                                                g2List.Add(selectScore[id].gsSchoolYear21.Value);
                                        }

                                        if (selectScore[id].gsSchoolYear22.HasValue)
                                        {
                                            if (!g2List.Contains(selectScore[id].gsSchoolYear22.Value))
                                                g2List.Add(selectScore[id].gsSchoolYear22.Value);
                                        }

                                        if (selectScore[id].gsSchoolYear31.HasValue)
                                        {
                                            if (!g3List.Contains(selectScore[id].gsSchoolYear31.Value))
                                                g3List.Add(selectScore[id].gsSchoolYear31.Value);
                                        }
                                        if (selectScore[id].gsSchoolYear32.HasValue)
                                        {
                                            if (!g3List.Contains(selectScore[id].gsSchoolYear32.Value))
                                                g3List.Add(selectScore[id].gsSchoolYear32.Value);
                                        }

                                        if (selectScore[id].gsSchoolYear41.HasValue)
                                        {
                                            if (!g4List.Contains(selectScore[id].gsSchoolYear41.Value))
                                                g4List.Add(selectScore[id].gsSchoolYear41.Value);
                                        }

                                        if (selectScore[id].gsSchoolYear42.HasValue)
                                        {
                                            if (!g4List.Contains(selectScore[id].gsSchoolYear42.Value))
                                                g4List.Add(selectScore[id].gsSchoolYear42.Value);
                                        }

                                        if (selectScore[id].gsScore11.HasValue)
                                            row["一上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore11.Value);
                                        if (selectScore[id].gsCredit11.HasValue)
                                            row["一上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit11.Value);
                                        if (selectScore[id].gsScore12.HasValue)
                                            row["一下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore12.Value);

                                        if (selectScore[id].gsCredit12.HasValue)
                                            row["一下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit12.Value);
                                        if (selectScore[id].gsScore21.HasValue)
                                            row["二上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore21.Value);

                                        if (selectScore[id].gsCredit21.HasValue)
                                            row["二上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit21.Value);

                                        if (selectScore[id].gsScore22.HasValue)
                                            row["二下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore22.Value);

                                        if (selectScore[id].gsCredit22.HasValue)
                                            row["二下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit22.Value);

                                        if (selectScore[id].gsScore31.HasValue)
                                            row["三上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore31.Value);

                                        if (selectScore[id].gsCredit31.HasValue)
                                            row["三上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit31.Value);

                                        if (selectScore[id].gsScore32.HasValue)
                                            row["三下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore32.Value);

                                        if (selectScore[id].gsCredit32.HasValue)
                                            row["三下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit32.Value);

                                        if (selectScore[id].gsScore41.HasValue)
                                            row["四上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore41.Value);

                                        if (selectScore[id].gsCredit41.HasValue)
                                            row["四上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit41.Value);

                                        if (selectScore[id].gsScore42.HasValue)
                                            row["四下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore42.Value);

                                        if (selectScore[id].gsCredit42.HasValue)
                                            row["四下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit42.Value);

                                        row["科目平均" + subjIndex] = Utility.NoRound(selectScore[id].avgScore);
                                        row["科目總分" + subjIndex] = Utility.NoRound(selectScore[id].sumScore);
                                        row["科目加權平均" + subjIndex] = Utility.NoRound(selectScore[id].avgScoreA);
                                        row["科目加權總分" + subjIndex] = Utility.NoRound(selectScore[id].sumScoreA);

                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key4 = "加權平均班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key4))
                                            {
                                                int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                row["科目加權平均班排名" + subjIndex] = rr;
                                                row["科目加權平均班排名母數" + subjIndex] = ranks[key4].Count;
                                                row["科目加權平均班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                            }

                                            key4 = "加權平均科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key4))
                                            {
                                                int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                row["科目加權平均科排名" + subjIndex] = rr;
                                                row["科目加權平均科排名母數" + subjIndex] = ranks[key4].Count;
                                                row["科目加權平均科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);

                                            }
                                            key4 = "加權平均全校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key4))
                                            {
                                                int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                row["科目加權平均校排名" + subjIndex] = rr;
                                                row["科目加權平均校排名母數" + subjIndex] = ranks[key4].Count;
                                                row["科目加權平均校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                            }
                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key4 = "加權平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key4))
                                                {
                                                    if (ranks[key4].IndexOf(selectScore[id].avgScoreAC1) >= 0)
                                                    {
                                                        int rr = ranks[key4].IndexOf(selectScore[id].avgScoreAC1) + 1;
                                                        row["科目加權平均類別一排名" + subjIndex] = rr;
                                                        row["科目加權平均類別一排名母數" + subjIndex] = ranks[key4].Count;
                                                        row["科目加權平均類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                    }
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                if (ranks.ContainsKey(key4))
                                                {
                                                    if (ranks[key4].IndexOf(selectScore[id].avgScoreAC2) >= 0)
                                                    {
                                                        int rr = ranks[key4].IndexOf(selectScore[id].avgScoreAC2) + 1;
                                                        row["科目加權平均類別二排名" + subjIndex] = rr;
                                                        row["科目加權平均類別二排名母數" + subjIndex] = ranks[key4].Count;
                                                        row["科目加權平均類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                    }
                                                }
                                            }

                                            string key2 = "加權總分班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                row["科目加權總分班排名" + subjIndex] = rr;
                                                row["科目加權總分班排名母數" + subjIndex] = ranks[key2].Count;
                                                row["科目加權總分班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }
                                            key2 = "加權總分科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                row["科目加權總分科排名" + subjIndex] = rr;
                                                row["科目加權總分科排名母數" + subjIndex] = ranks[key2].Count;
                                                row["科目加權總分科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                            }

                                            key2 = "加權總分全校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                row["科目加權總分校排名" + subjIndex] = rr;
                                                row["科目加權總分校排名母數" + subjIndex] = ranks[key2].Count;
                                                row["科目加權總分校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key2 = "加權總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    if (ranks[key2].IndexOf(selectScore[id].sumScoreAC1) >= 0)
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id].sumScoreAC1) + 1;
                                                        row["科目加權總分類別一排名" + subjIndex] = rr;
                                                        row["科目加權總分類別一排名母數" + subjIndex] = ranks[key2].Count;
                                                        row["科目加權總分類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key2 = "加權總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    if (ranks[key2].IndexOf(selectScore[id].sumScoreAC2) >= 0)
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id].sumScoreAC2) + 1;
                                                        row["科目加權總分類別二排名" + subjIndex] = rr;
                                                        row["科目加權總分類別二排名母數" + subjIndex] = ranks[key2].Count;
                                                        row["科目加權總分類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }
                                                }
                                            }

                                            string key3 = "平均班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key3))
                                            {
                                                int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                row["科目平均班排名" + subjIndex] = rr;
                                                row["科目平均班排名母數" + subjIndex] = ranks[key3].Count;
                                                row["科目平均班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                            }

                                            key3 = "平均科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key3))
                                            {
                                                int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                row["科目平均科排名" + subjIndex] = rr;
                                                row["科目平均科排名母數" + subjIndex] = ranks[key3].Count;
                                                row["科目平均科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                            }
                                            key3 = "平均全校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key3))
                                            {
                                                int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                row["科目平均校排名" + subjIndex] = rr;
                                                row["科目平均校排名母數" + subjIndex] = ranks[key3].Count;
                                                row["科目平均校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key3 = "平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key3))
                                                {
                                                    if (ranks[key3].IndexOf(selectScore[id].avgScoreC1) >= 0)
                                                    {
                                                        int rr = ranks[key3].IndexOf(selectScore[id].avgScoreC1) + 1;
                                                        row["科目平均類別一排名" + subjIndex] = rr;
                                                        row["科目平均類別一排名母數" + subjIndex] = ranks[key3].Count;
                                                        row["科目平均類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                    }
                                                }
                                            }

                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key3 = "平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key3))
                                                {
                                                    if (ranks[key3].IndexOf(selectScore[id].avgScoreC2) >= 0)
                                                    {
                                                        int rr = ranks[key3].IndexOf(selectScore[id].avgScoreC2) + 1;
                                                        row["科目平均類別二排名" + subjIndex] = rr;
                                                        row["科目平均類別二排名母數" + subjIndex] = ranks[key3].Count;
                                                        row["科目平均類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                    }
                                                }
                                            }
                                            string key1 = "總分班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                row["科目總分班排名" + subjIndex] = rr;
                                                row["科目總分班排名母數" + subjIndex] = ranks[key1].Count;
                                                row["科目總分班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                            }
                                            key1 = "總分科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                row["科目總分科排名" + subjIndex] = rr;
                                                row["科目總分科排名母數" + subjIndex] = ranks[key1].Count;
                                                row["科目總分科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                            }
                                            key1 = "總分全校排名" + gradeyear + "^^^" + subjName;

                                            if (ranks.ContainsKey(key1))
                                            {
                                                int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                row["科目總分校排名" + subjIndex] = rr;
                                                row["科目總分校排名母數" + subjIndex] = ranks[key1].Count;
                                                row["科目總分校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key1 = "總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    if (ranks[key1].IndexOf(selectScore[id].sumScoreC1) >= 0)
                                                    {
                                                        int rr = ranks[key1].IndexOf(selectScore[id].sumScoreC1) + 1;
                                                        row["科目總分類別一排名" + subjIndex] = rr;
                                                        row["科目總分類別一排名母數" + subjIndex] = ranks[key1].Count;
                                                        row["科目總分類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                    }
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key1 = "總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    if (ranks[key1].IndexOf(selectScore[id].sumScoreC2) >= 0)
                                                    {
                                                        int rr = ranks[key1].IndexOf(selectScore[id].sumScoreC2) + 1;
                                                        row["科目總分類別二排名" + subjIndex] = rr;
                                                        row["科目總分類別二排名母數" + subjIndex] = ranks[key1].Count;
                                                        row["科目總分類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                    }
                                                }
                                            }
                                        }

                                        subjIndex++;
                                    }
                                }
                                #endregion 處理科目

                                subjIndex = 1;
                                // 處理科目回歸
                                #region 處理科目回歸
                                foreach (string subjName in SubjMappingDict.Keys)
                                {
                                    string id = studRec.StudentID + "^^^回歸" + subjName;

                                    if (selectScore.ContainsKey(id))
                                    {
                                        row["回歸科目名稱" + subjIndex] = subjName;

                                        if (selectScore[id].gsSchoolYear11.HasValue)
                                        {
                                            if (!g1List.Contains(selectScore[id].gsSchoolYear11.Value))
                                                g1List.Add(selectScore[id].gsSchoolYear11.Value);
                                        }
                                        if (selectScore[id].gsSchoolYear12.HasValue)
                                        {
                                            if (!g1List.Contains(selectScore[id].gsSchoolYear12.Value))
                                                g1List.Add(selectScore[id].gsSchoolYear12.Value);
                                        }

                                        if (selectScore[id].gsSchoolYear21.HasValue)
                                        {
                                            if (!g2List.Contains(selectScore[id].gsSchoolYear21.Value))
                                                g2List.Add(selectScore[id].gsSchoolYear21.Value);
                                        }

                                        if (selectScore[id].gsSchoolYear22.HasValue)
                                        {
                                            if (!g2List.Contains(selectScore[id].gsSchoolYear22.Value))
                                                g2List.Add(selectScore[id].gsSchoolYear22.Value);
                                        }

                                        if (selectScore[id].gsSchoolYear31.HasValue)
                                        {
                                            if (!g3List.Contains(selectScore[id].gsSchoolYear31.Value))
                                                g3List.Add(selectScore[id].gsSchoolYear31.Value);
                                        }
                                        if (selectScore[id].gsSchoolYear32.HasValue)
                                        {
                                            if (!g3List.Contains(selectScore[id].gsSchoolYear32.Value))
                                                g3List.Add(selectScore[id].gsSchoolYear32.Value);
                                        }

                                        if (selectScore[id].gsSchoolYear41.HasValue)
                                        {
                                            if (!g4List.Contains(selectScore[id].gsSchoolYear41.Value))
                                                g4List.Add(selectScore[id].gsSchoolYear41.Value);
                                        }

                                        if (selectScore[id].gsSchoolYear42.HasValue)
                                        {
                                            if (!g4List.Contains(selectScore[id].gsSchoolYear42.Value))
                                                g4List.Add(selectScore[id].gsSchoolYear42.Value);
                                        }

                                        if (selectScore[id].gsScore11.HasValue)
                                            row["回歸一上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore11.Value);
                                        if (selectScore[id].gsCredit11.HasValue)
                                            row["回歸一上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit11.Value);
                                        if (selectScore[id].gsScore12.HasValue)
                                            row["回歸一下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore12.Value);

                                        if (selectScore[id].gsCredit12.HasValue)
                                            row["回歸一下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit12.Value);
                                        if (selectScore[id].gsScore21.HasValue)
                                            row["回歸二上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore21.Value);

                                        if (selectScore[id].gsCredit21.HasValue)
                                            row["回歸二上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit21.Value);

                                        if (selectScore[id].gsScore22.HasValue)
                                            row["回歸二下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore22.Value);

                                        if (selectScore[id].gsCredit22.HasValue)
                                            row["回歸二下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit22.Value);

                                        if (selectScore[id].gsScore31.HasValue)
                                            row["回歸三上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore31.Value);

                                        if (selectScore[id].gsCredit31.HasValue)
                                            row["回歸三上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit31.Value);

                                        if (selectScore[id].gsScore32.HasValue)
                                            row["回歸三下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore32.Value);

                                        if (selectScore[id].gsCredit32.HasValue)
                                            row["回歸三下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit32.Value);

                                        if (selectScore[id].gsScore41.HasValue)
                                            row["回歸四上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore41.Value);

                                        row["回歸科目平均" + subjIndex] = Utility.NoRound(selectScore[id].avgScore);
                                        row["回歸科目總分" + subjIndex] = Utility.NoRound(selectScore[id].sumScore);
                                        row["回歸科目加權平均" + subjIndex] = Utility.NoRound(selectScore[id].avgScoreA);
                                        row["回歸科目加權總分" + subjIndex] = Utility.NoRound(selectScore[id].sumScoreA);

                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {

                                            #region 寫入各科目成績班排名
                                            string key11 = "回歸一上科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key11))
                                            {
                                                if (selectScore[id].gsScore11.HasValue)
                                                {
                                                    int rr = ranks[key11].IndexOf(selectScore[id].gsScore11.Value) + 1;
                                                    row["回歸一上科目班排名" + subjIndex] = rr;
                                                    row["回歸一上科目班排名母數" + subjIndex] = ranks[key11].Count;
                                                    row["回歸一上科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key11].Count);
                                                }
                                            }

                                            string key12 = "回歸一下科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key12))
                                            {
                                                if (selectScore[id].gsScore12.HasValue)
                                                {
                                                    int rr = ranks[key12].IndexOf(selectScore[id].gsScore12.Value) + 1;
                                                    row["回歸一下科目班排名" + subjIndex] = rr;
                                                    row["回歸一下科目班排名母數" + subjIndex] = ranks[key12].Count;
                                                    row["回歸一下科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key12].Count);
                                                }
                                            }

                                            string key21 = "回歸二上科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key21))
                                            {
                                                if (selectScore[id].gsScore21.HasValue)
                                                {
                                                    int rr = ranks[key21].IndexOf(selectScore[id].gsScore21.Value) + 1;
                                                    row["回歸二上科目班排名" + subjIndex] = rr;
                                                    row["回歸二上科目班排名母數" + subjIndex] = ranks[key21].Count;
                                                    row["回歸二上科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key21].Count);
                                                }
                                            }

                                            string key22 = "回歸二下科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key22))
                                            {
                                                if (selectScore[id].gsScore22.HasValue)
                                                {
                                                    int rr = ranks[key22].IndexOf(selectScore[id].gsScore22.Value) + 1;
                                                    row["回歸二下科目班排名" + subjIndex] = rr;
                                                    row["回歸二下科目班排名母數" + subjIndex] = ranks[key22].Count;
                                                    row["回歸二下科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key22].Count);
                                                }
                                            }

                                            string key31 = "回歸三上科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key31))
                                            {
                                                if (selectScore[id].gsScore31.HasValue)
                                                {
                                                    int rr = ranks[key31].IndexOf(selectScore[id].gsScore31.Value) + 1;
                                                    row["回歸三上科目班排名" + subjIndex] = rr;
                                                    row["回歸三上科目班排名母數" + subjIndex] = ranks[key31].Count;
                                                    row["回歸三上科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key31].Count);
                                                }
                                            }

                                            string key32 = "回歸三下科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key32))
                                            {
                                                if (selectScore[id].gsScore32.HasValue)
                                                {
                                                    int rr = ranks[key32].IndexOf(selectScore[id].gsScore32.Value) + 1;
                                                    row["回歸三下科目班排名" + subjIndex] = rr;
                                                    row["回歸三下科目班排名母數" + subjIndex] = ranks[key32].Count;
                                                    row["回歸三下科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key32].Count);
                                                }
                                            }

                                            string key41 = "回歸四上科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key41))
                                            {
                                                if (selectScore[id].gsScore41.HasValue)
                                                {
                                                    int rr = ranks[key41].IndexOf(selectScore[id].gsScore41.Value) + 1;
                                                    row["回歸四上科目班排名" + subjIndex] = rr;
                                                    row["回歸四上科目班排名母數" + subjIndex] = ranks[key41].Count;
                                                    row["回歸四上科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key41].Count);
                                                }
                                            }

                                            string key42 = "回歸四下科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key42))
                                            {
                                                if (selectScore[id].gsScore42.HasValue)
                                                {
                                                    int rr = ranks[key42].IndexOf(selectScore[id].gsScore42.Value) + 1;
                                                    row["回歸四下科目班排名" + subjIndex] = rr;
                                                    row["回歸四下科目班排名母數" + subjIndex] = ranks[key42].Count;
                                                    row["回歸四下科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key42].Count);
                                                }
                                            }

                                            #endregion

                                            #region 寫入各科目成績科排名
                                            key11 = "回歸一上科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key11))
                                            {
                                                if (selectScore[id].gsScore11.HasValue)
                                                {
                                                    int rr = ranks[key11].IndexOf(selectScore[id].gsScore11.Value) + 1;
                                                    row["回歸一上科目科排名" + subjIndex] = rr;
                                                    row["回歸一上科目科排名母數" + subjIndex] = ranks[key11].Count;
                                                    row["回歸一上科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key11].Count);
                                                }
                                            }

                                            key12 = "回歸一下科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key12))
                                            {
                                                if (selectScore[id].gsScore12.HasValue)
                                                {
                                                    int rr = ranks[key12].IndexOf(selectScore[id].gsScore12.Value) + 1;
                                                    row["回歸一下科目科排名" + subjIndex] = rr;
                                                    row["回歸一下科目科排名母數" + subjIndex] = ranks[key12].Count;
                                                    row["回歸一下科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key12].Count);
                                                }
                                            }

                                            key21 = "回歸二上科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key21))
                                            {
                                                if (selectScore[id].gsScore21.HasValue)
                                                {
                                                    int rr = ranks[key21].IndexOf(selectScore[id].gsScore21.Value) + 1;
                                                    row["回歸二上科目科排名" + subjIndex] = rr;
                                                    row["回歸二上科目科排名母數" + subjIndex] = ranks[key21].Count;
                                                    row["回歸二上科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key21].Count);
                                                }
                                            }

                                            key22 = "回歸二下科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key22))
                                            {
                                                if (selectScore[id].gsScore22.HasValue)
                                                {
                                                    int rr = ranks[key22].IndexOf(selectScore[id].gsScore22.Value) + 1;
                                                    row["回歸二下科目科排名" + subjIndex] = rr;
                                                    row["回歸二下科目科排名母數" + subjIndex] = ranks[key22].Count;
                                                    row["回歸二下科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key22].Count);
                                                }
                                            }

                                            key31 = "回歸三上科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key31))
                                            {
                                                if (selectScore[id].gsScore31.HasValue)
                                                {
                                                    int rr = ranks[key31].IndexOf(selectScore[id].gsScore31.Value) + 1;
                                                    row["回歸三上科目科排名" + subjIndex] = rr;
                                                    row["回歸三上科目科排名母數" + subjIndex] = ranks[key31].Count;
                                                    row["回歸三上科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key31].Count);
                                                }
                                            }

                                            key32 = "回歸三下科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key32))
                                            {
                                                if (selectScore[id].gsScore32.HasValue)
                                                {
                                                    int rr = ranks[key32].IndexOf(selectScore[id].gsScore32.Value) + 1;
                                                    row["回歸三下科目科排名" + subjIndex] = rr;
                                                    row["回歸三下科目科排名母數" + subjIndex] = ranks[key32].Count;
                                                    row["回歸三下科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key32].Count);
                                                }
                                            }

                                            key41 = "回歸四上科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key41))
                                            {
                                                if (selectScore[id].gsScore41.HasValue)
                                                {
                                                    int rr = ranks[key41].IndexOf(selectScore[id].gsScore41.Value) + 1;
                                                    row["回歸四上科目科排名" + subjIndex] = rr;
                                                    row["回歸四上科目科排名母數" + subjIndex] = ranks[key41].Count;
                                                    row["回歸四上科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key41].Count);
                                                }
                                            }

                                            key42 = "回歸四下科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key42))
                                            {
                                                if (selectScore[id].gsScore42.HasValue)
                                                {
                                                    int rr = ranks[key42].IndexOf(selectScore[id].gsScore42.Value) + 1;
                                                    row["回歸四下科目科排名" + subjIndex] = rr;
                                                    row["回歸四下科目科排名母數" + subjIndex] = ranks[key42].Count;
                                                    row["回歸四下科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key42].Count);
                                                }
                                            }

                                            #endregion

                                            #region 寫入各科目成績校排名
                                            key11 = "回歸一上科目校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key11))
                                            {
                                                if (selectScore[id].gsScore11.HasValue)
                                                {
                                                    int rr = ranks[key11].IndexOf(selectScore[id].gsScore11.Value) + 1;
                                                    row["回歸一上科目校排名" + subjIndex] = rr;
                                                    row["回歸一上科目校排名母數" + subjIndex] = ranks[key11].Count;
                                                    row["回歸一上科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key11].Count);
                                                }
                                            }

                                            key12 = "回歸一下科目校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key12))
                                            {
                                                if (selectScore[id].gsScore12.HasValue)
                                                {
                                                    int rr = ranks[key12].IndexOf(selectScore[id].gsScore12.Value) + 1;
                                                    row["回歸一下科目校排名" + subjIndex] = rr;
                                                    row["回歸一下科目校排名母數" + subjIndex] = ranks[key12].Count;
                                                    row["回歸一下科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key12].Count);
                                                }
                                            }

                                            key21 = "回歸二上科目校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key21))
                                            {
                                                if (selectScore[id].gsScore21.HasValue)
                                                {
                                                    int rr = ranks[key21].IndexOf(selectScore[id].gsScore21.Value) + 1;
                                                    row["回歸二上科目校排名" + subjIndex] = rr;
                                                    row["回歸二上科目校排名母數" + subjIndex] = ranks[key21].Count;
                                                    row["回歸二上科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key21].Count);
                                                }
                                            }

                                            key22 = "回歸二下科目校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key22))
                                            {
                                                if (selectScore[id].gsScore22.HasValue)
                                                {
                                                    int rr = ranks[key22].IndexOf(selectScore[id].gsScore22.Value) + 1;
                                                    row["回歸二下科目校排名" + subjIndex] = rr;
                                                    row["回歸二下科目校排名母數" + subjIndex] = ranks[key22].Count;
                                                    row["回歸二下科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key22].Count);
                                                }
                                            }

                                            key31 = "回歸三上科目校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key31))
                                            {
                                                if (selectScore[id].gsScore31.HasValue)
                                                {
                                                    int rr = ranks[key31].IndexOf(selectScore[id].gsScore31.Value) + 1;
                                                    row["回歸三上科目校排名" + subjIndex] = rr;
                                                    row["回歸三上科目校排名母數" + subjIndex] = ranks[key31].Count;
                                                    row["回歸三上科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key31].Count);
                                                }
                                            }

                                            key32 = "回歸三下科目校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key32))
                                            {
                                                if (selectScore[id].gsScore32.HasValue)
                                                {
                                                    int rr = ranks[key32].IndexOf(selectScore[id].gsScore32.Value) + 1;
                                                    row["回歸三下科目校排名" + subjIndex] = rr;
                                                    row["回歸三下科目校排名母數" + subjIndex] = ranks[key32].Count;
                                                    row["回歸三下科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key32].Count);
                                                }
                                            }

                                            key41 = "回歸四上科目校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key41))
                                            {
                                                if (selectScore[id].gsScore41.HasValue)
                                                {
                                                    int rr = ranks[key41].IndexOf(selectScore[id].gsScore41.Value) + 1;
                                                    row["回歸四上科目校排名" + subjIndex] = rr;
                                                    row["回歸四上科目校排名母數" + subjIndex] = ranks[key41].Count;
                                                    row["回歸四上科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key41].Count);
                                                }
                                            }

                                            key42 = "回歸四下科目校排名" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key42))
                                            {
                                                if (selectScore[id].gsScore42.HasValue)
                                                {
                                                    int rr = ranks[key42].IndexOf(selectScore[id].gsScore42.Value) + 1;
                                                    row["回歸四下科目校排名" + subjIndex] = rr;
                                                    row["回歸四下科目校排名母數" + subjIndex] = ranks[key42].Count;
                                                    row["回歸四下科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key42].Count);
                                                }
                                            }

                                            #endregion

                                            #region 寫入各科目成績類1排名
                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key11 = "回歸一上科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key11))
                                                {
                                                    if (selectScore[id].gsScore11.HasValue)
                                                    {
                                                        int rr = ranks[key11].IndexOf(selectScore[id].gsScore11.Value) + 1;
                                                        row["回歸一上科目類1排名" + subjIndex] = rr;
                                                        row["回歸一上科目類1排名母數" + subjIndex] = ranks[key11].Count;
                                                        row["回歸一上科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key11].Count);
                                                    }
                                                }

                                                key12 = "回歸一下科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key12))
                                                {
                                                    if (selectScore[id].gsScore12.HasValue)
                                                    {
                                                        int rr = ranks[key12].IndexOf(selectScore[id].gsScore12.Value) + 1;
                                                        row["回歸一下科目類1排名" + subjIndex] = rr;
                                                        row["回歸一下科目類1排名母數" + subjIndex] = ranks[key12].Count;
                                                        row["回歸一下科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key12].Count);
                                                    }
                                                }

                                                key21 = "回歸二上科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key21))
                                                {
                                                    if (selectScore[id].gsScore21.HasValue)
                                                    {
                                                        int rr = ranks[key21].IndexOf(selectScore[id].gsScore21.Value) + 1;
                                                        row["回歸二上科目類1排名" + subjIndex] = rr;
                                                        row["回歸二上科目類1排名母數" + subjIndex] = ranks[key21].Count;
                                                        row["回歸二上科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key21].Count);
                                                    }
                                                }

                                                key22 = "回歸二下科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key22))
                                                {
                                                    if (selectScore[id].gsScore22.HasValue)
                                                    {
                                                        int rr = ranks[key22].IndexOf(selectScore[id].gsScore22.Value) + 1;
                                                        row["回歸二下科目類1排名" + subjIndex] = rr;
                                                        row["回歸二下科目類1排名母數" + subjIndex] = ranks[key22].Count;
                                                        row["回歸二下科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key22].Count);
                                                    }
                                                }

                                                key31 = "回歸三上科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key31))
                                                {
                                                    if (selectScore[id].gsScore31.HasValue)
                                                    {
                                                        int rr = ranks[key31].IndexOf(selectScore[id].gsScore31.Value) + 1;
                                                        row["回歸三上科目類1排名" + subjIndex] = rr;
                                                        row["回歸三上科目類1排名母數" + subjIndex] = ranks[key31].Count;
                                                        row["回歸三上科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key31].Count);
                                                    }
                                                }

                                                key32 = "回歸三下科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key32))
                                                {
                                                    if (selectScore[id].gsScore32.HasValue)
                                                    {
                                                        int rr = ranks[key32].IndexOf(selectScore[id].gsScore32.Value) + 1;
                                                        row["回歸三下科目類1排名" + subjIndex] = rr;
                                                        row["回歸三下科目類1排名母數" + subjIndex] = ranks[key32].Count;
                                                        row["回歸三下科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key32].Count);
                                                    }
                                                }

                                                key41 = "回歸四上科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key41))
                                                {
                                                    if (selectScore[id].gsScore41.HasValue)
                                                    {
                                                        int rr = ranks[key41].IndexOf(selectScore[id].gsScore41.Value) + 1;
                                                        row["回歸四上科目類1排名" + subjIndex] = rr;
                                                        row["回歸四上科目類1排名母數" + subjIndex] = ranks[key41].Count;
                                                        row["回歸四上科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key41].Count);
                                                    }
                                                }

                                                key42 = "回歸四下科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key42))
                                                {
                                                    if (selectScore[id].gsScore42.HasValue)
                                                    {
                                                        int rr = ranks[key42].IndexOf(selectScore[id].gsScore42.Value) + 1;
                                                        row["回歸四下科目類1排名" + subjIndex] = rr;
                                                        row["回歸四下科目類1排名母數" + subjIndex] = ranks[key42].Count;
                                                        row["回歸四下科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key42].Count);
                                                    }
                                                }
                                            }

                                            #endregion

                                            #region 寫入各科目成績類2排名

                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key11 = "回歸一上科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key11))
                                                {
                                                    if (selectScore[id].gsScore11.HasValue)
                                                    {
                                                        int rr = ranks[key11].IndexOf(selectScore[id].gsScore11.Value) + 1;
                                                        row["回歸一上科目類2排名" + subjIndex] = rr;
                                                        row["回歸一上科目類2排名母數" + subjIndex] = ranks[key11].Count;
                                                        row["回歸一上科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key11].Count);
                                                    }
                                                }

                                                key12 = "回歸一下科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key12))
                                                {
                                                    if (selectScore[id].gsScore12.HasValue)
                                                    {
                                                        int rr = ranks[key12].IndexOf(selectScore[id].gsScore12.Value) + 1;
                                                        row["回歸一下科目類2排名" + subjIndex] = rr;
                                                        row["回歸一下科目類2排名母數" + subjIndex] = ranks[key12].Count;
                                                        row["回歸一下科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key12].Count);
                                                    }
                                                }

                                                key21 = "回歸二上科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key21))
                                                {
                                                    if (selectScore[id].gsScore21.HasValue)
                                                    {
                                                        int rr = ranks[key21].IndexOf(selectScore[id].gsScore21.Value) + 1;
                                                        row["回歸二上科目類2排名" + subjIndex] = rr;
                                                        row["回歸二上科目類2排名母數" + subjIndex] = ranks[key21].Count;
                                                        row["回歸二上科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key21].Count);
                                                    }
                                                }

                                                key22 = "回歸二下科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key22))
                                                {
                                                    if (selectScore[id].gsScore22.HasValue)
                                                    {
                                                        int rr = ranks[key22].IndexOf(selectScore[id].gsScore22.Value) + 1;
                                                        row["回歸二下科目類2排名" + subjIndex] = rr;
                                                        row["回歸二下科目類2排名母數" + subjIndex] = ranks[key22].Count;
                                                        row["回歸二下科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key22].Count);
                                                    }
                                                }

                                                key31 = "回歸三上科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key31))
                                                {
                                                    if (selectScore[id].gsScore31.HasValue)
                                                    {
                                                        int rr = ranks[key31].IndexOf(selectScore[id].gsScore31.Value) + 1;
                                                        row["回歸三上科目類2排名" + subjIndex] = rr;
                                                        row["回歸三上科目類2排名母數" + subjIndex] = ranks[key31].Count;
                                                        row["回歸三上科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key31].Count);
                                                    }
                                                }

                                                key32 = "回歸三下科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key32))
                                                {
                                                    if (selectScore[id].gsScore32.HasValue)
                                                    {
                                                        int rr = ranks[key32].IndexOf(selectScore[id].gsScore32.Value) + 1;
                                                        row["回歸三下科目類2排名" + subjIndex] = rr;
                                                        row["回歸三下科目類2排名母數" + subjIndex] = ranks[key32].Count;
                                                        row["回歸三下科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key32].Count);
                                                    }
                                                }

                                                key41 = "回歸四上科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key41))
                                                {
                                                    if (selectScore[id].gsScore41.HasValue)
                                                    {
                                                        int rr = ranks[key41].IndexOf(selectScore[id].gsScore41.Value) + 1;
                                                        row["回歸四上科目類2排名" + subjIndex] = rr;
                                                        row["回歸四上科目類2排名母數" + subjIndex] = ranks[key41].Count;
                                                        row["回歸四上科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key41].Count);
                                                    }
                                                }

                                                key42 = "回歸四下科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key42))
                                                {
                                                    if (selectScore[id].gsScore42.HasValue)
                                                    {
                                                        int rr = ranks[key42].IndexOf(selectScore[id].gsScore42.Value) + 1;
                                                        row["回歸四下科目類2排名" + subjIndex] = rr;
                                                        row["回歸四下科目類2排名母數" + subjIndex] = ranks[key42].Count;
                                                        row["回歸四下科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key42].Count);
                                                    }
                                                }
                                            }

                                            #endregion


                                            string key4 = "加權平均班排名_回歸_" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key4))
                                            {
                                                int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                row["回歸科目加權平均班排名" + subjIndex] = rr;
                                                row["回歸科目加權平均班排名母數" + subjIndex] = ranks[key4].Count;
                                                row["回歸科目加權平均班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                            }

                                            key4 = "加權平均科排名_回歸_" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key4))
                                            {
                                                int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                row["回歸科目加權平均科排名" + subjIndex] = rr;
                                                row["回歸科目加權平均科排名母數" + subjIndex] = ranks[key4].Count;
                                                row["回歸科目加權平均科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);

                                            }
                                            key4 = "加權平均全校排名_回歸_" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key4))
                                            {
                                                int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                row["回歸科目加權平均校排名" + subjIndex] = rr;
                                                row["回歸科目加權平均校排名母數" + subjIndex] = ranks[key4].Count;
                                                row["回歸科目加權平均校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                            }
                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key4 = "加權平均類別1排名_回歸_" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key4))
                                                {
                                                    if (ranks[key4].IndexOf(selectScore[id].avgScoreAC1) >= 0)
                                                    {
                                                        int rr = ranks[key4].IndexOf(selectScore[id].avgScoreAC1) + 1;
                                                        row["回歸科目加權平均類別一排名" + subjIndex] = rr;
                                                        row["回歸科目加權平均類別一排名母數" + subjIndex] = ranks[key4].Count;
                                                        row["回歸科目加權平均類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                    }
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                if (ranks.ContainsKey(key4))
                                                {
                                                    if (ranks[key4].IndexOf(selectScore[id].avgScoreAC2) >= 0)
                                                    {
                                                        int rr = ranks[key4].IndexOf(selectScore[id].avgScoreAC2) + 1;
                                                        row["回歸科目加權平均類別二排名" + subjIndex] = rr;
                                                        row["回歸科目加權平均類別二排名母數" + subjIndex] = ranks[key4].Count;
                                                        row["回歸科目加權平均類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                    }
                                                }
                                            }

                                            string key2 = "加權總分班排名_回歸_" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                row["回歸科目加權總分班排名" + subjIndex] = rr;
                                                row["回歸科目加權總分班排名母數" + subjIndex] = ranks[key2].Count;
                                                row["回歸科目加權總分班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }
                                            key2 = "加權總分科排名_回歸_" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                row["回歸科目加權總分科排名" + subjIndex] = rr;
                                                row["回歸科目加權總分科排名母數" + subjIndex] = ranks[key2].Count;
                                                row["回歸科目加權總分科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                            }

                                            key2 = "加權總分全校排名_回歸_" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                row["回歸科目加權總分校排名" + subjIndex] = rr;
                                                row["回歸科目加權總分校排名母數" + subjIndex] = ranks[key2].Count;
                                                row["回歸科目加權總分校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key2 = "加權總分類別1排名_回歸_" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    if (ranks[key2].IndexOf(selectScore[id].sumScoreAC1) >= 0)
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id].sumScoreAC1) + 1;
                                                        row["回歸科目加權總分類別一排名" + subjIndex] = rr;
                                                        row["回歸科目加權總分類別一排名母數" + subjIndex] = ranks[key2].Count;
                                                        row["回歸科目加權總分類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key2 = "加權總分類別2排名_回歸_" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    if (ranks[key2].IndexOf(selectScore[id].sumScoreAC2) >= 0)
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id].sumScoreAC2) + 1;
                                                        row["回歸科目加權總分類別二排名" + subjIndex] = rr;
                                                        row["回歸科目加權總分類別二排名母數" + subjIndex] = ranks[key2].Count;
                                                        row["回歸科目加權總分類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }
                                                }
                                            }

                                            string key3 = "平均班排名_回歸_" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key3))
                                            {
                                                int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                row["回歸科目平均班排名" + subjIndex] = rr;
                                                row["回歸科目平均班排名母數" + subjIndex] = ranks[key3].Count;
                                                row["回歸科目平均班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                            }

                                            key3 = "平均科排名_回歸_" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key3))
                                            {
                                                int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                row["回歸科目平均科排名" + subjIndex] = rr;
                                                row["回歸科目平均科排名母數" + subjIndex] = ranks[key3].Count;
                                                row["回歸科目平均科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                            }
                                            key3 = "平均全校排名_回歸_" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key3))
                                            {
                                                int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                row["回歸科目平均校排名" + subjIndex] = rr;
                                                row["回歸科目平均校排名母數" + subjIndex] = ranks[key3].Count;
                                                row["回歸科目平均校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key3 = "平均類別1排名_回歸_" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key3))
                                                {
                                                    if (ranks[key3].IndexOf(selectScore[id].avgScoreC1) >= 0)
                                                    {
                                                        int rr = ranks[key3].IndexOf(selectScore[id].avgScoreC1) + 1;
                                                        row["回歸科目平均類別一排名" + subjIndex] = rr;
                                                        row["回歸科目平均類別一排名母數" + subjIndex] = ranks[key3].Count;
                                                        row["回歸科目平均類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                    }
                                                }
                                            }

                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key3 = "平均類別2排名_回歸_" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key3))
                                                {
                                                    if (ranks[key3].IndexOf(selectScore[id].avgScoreC2) >= 0)
                                                    {
                                                        int rr = ranks[key3].IndexOf(selectScore[id].avgScoreC2) + 1;
                                                        row["回歸科目平均類別二排名" + subjIndex] = rr;
                                                        row["回歸科目平均類別二排名母數" + subjIndex] = ranks[key3].Count;
                                                        row["回歸科目平均類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                    }
                                                }
                                            }
                                            string key1 = "總分班排名_回歸_" + studRec.RefClass.ClassID + "^^^" + subjName;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                row["回歸科目總分班排名" + subjIndex] = rr;
                                                row["回歸科目總分班排名母數" + subjIndex] = ranks[key1].Count;
                                                row["回歸科目總分班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                            }
                                            key1 = "總分科排名_回歸_" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                row["回歸科目總分科排名" + subjIndex] = rr;
                                                row["回歸科目總分科排名母數" + subjIndex] = ranks[key1].Count;
                                                row["回歸科目總分科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                            }
                                            key1 = "總分全校排名_回歸_" + gradeyear + "^^^" + subjName;

                                            if (ranks.ContainsKey(key1))
                                            {
                                                int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                row["回歸科目總分校排名" + subjIndex] = rr;
                                                row["回歸科目總分校排名母數" + subjIndex] = ranks[key1].Count;
                                                row["回歸科目總分校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key1 = "總分類別1排名_回歸_" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    if (ranks[key1].IndexOf(selectScore[id].sumScoreC1) >= 0)
                                                    {
                                                        int rr = ranks[key1].IndexOf(selectScore[id].sumScoreC1) + 1;
                                                        row["回歸科目總分類別一排名" + subjIndex] = rr;
                                                        row["回歸科目總分類別一排名母數" + subjIndex] = ranks[key1].Count;
                                                        row["回歸科目總分類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                    }
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key1 = "總分類別2排名_回歸_" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    if (ranks[key1].IndexOf(selectScore[id].sumScoreC2) >= 0)
                                                    {
                                                        int rr = ranks[key1].IndexOf(selectScore[id].sumScoreC2) + 1;
                                                        row["回歸科目總分類別二排名" + subjIndex] = rr;
                                                        row["回歸科目總分類別二排名母數" + subjIndex] = ranks[key1].Count;
                                                        row["回歸科目總分類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                    }
                                                }
                                            }
                                        }

                                        subjIndex++;
                                    }
                                }
                                #endregion 處理科目回歸



                                // 處理科目年級學年度
                                row["一年級學年度"] = string.Join("/", g1List.ToArray());
                                row["二年級學年度"] = string.Join("/", g2List.ToArray());
                                row["三年級學年度"] = string.Join("/", g3List.ToArray());
                                row["四年級學年度"] = string.Join("/", g4List.ToArray());

                                // 處理各學期各科目排名
                                if (StudSemsSubjRatingDict.ContainsKey(studRec.StudentID))
                                {
                                    foreach (StudSemsSubjRating ssr in StudSemsSubjRatingDict[studRec.StudentID])
                                    {

                                        #region 舊
                                        if (ssr.GradeYear == "1" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["一上科目排名" + i] = ssr.GetClassRank(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "1" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["一下科目排名" + i] = ssr.GetClassRank(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "2" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["二上科目排名" + i] = ssr.GetClassRank(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "2" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["二下科目排名" + i] = ssr.GetClassRank(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "3" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["三上科目排名" + i] = ssr.GetClassRank(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "3" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["三下科目排名" + i] = ssr.GetClassRank(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "4" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["四上科目排名" + i] = ssr.GetClassRank(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "4" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["四下科目排名" + i] = ssr.GetClassRank(subjName);
                                            }
                                        }
                                        #endregion

                                        #region 班排
                                        if (ssr.GradeYear == "1" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["一上科目班排名" + i] = ssr.GetClassRank(subjName);
                                                row["一上科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                row["一上科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "1" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["一下科目班排名" + i] = ssr.GetClassRank(subjName);
                                                row["一下科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                row["一下科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "2" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["二上科目班排名" + i] = ssr.GetClassRank(subjName);
                                                row["二上科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                row["二上科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "2" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["二下科目班排名" + i] = ssr.GetClassRank(subjName);
                                                row["二下科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                row["二下科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "3" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["三上科目班排名" + i] = ssr.GetClassRank(subjName);
                                                row["三上科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                row["三上科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "3" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["三下科目班排名" + i] = ssr.GetClassRank(subjName);
                                                row["三下科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                row["三下科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "4" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["四上科目班排名" + i] = ssr.GetClassRank(subjName);
                                                row["四上科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                row["四上科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "4" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["四下科目班排名" + i] = ssr.GetClassRank(subjName);
                                                row["四下科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                row["四下科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                            }
                                        }
                                        #endregion

                                        #region 科排
                                        if (ssr.GradeYear == "1" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["一上科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                row["一上科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                row["一上科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "1" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["一下科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                row["一下科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                row["一下科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "2" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["二上科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                row["二上科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                row["二上科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "2" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["二下科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                row["二下科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                row["二下科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "3" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["三上科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                row["三上科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                row["三上科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "3" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["三下科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                row["三下科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                row["三下科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "4" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["四上科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                row["四上科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                row["四上科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "4" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["四下科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                row["四下科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                row["四下科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                            }
                                        }
                                        #endregion                

                                        #region 校排
                                        if (ssr.GradeYear == "1" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["一上科目校排名" + i] = ssr.GetYearRank(subjName);
                                                row["一上科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                row["一上科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "1" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["一下科目校排名" + i] = ssr.GetYearRank(subjName);
                                                row["一下科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                row["一下科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "2" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["二上科目校排名" + i] = ssr.GetYearRank(subjName);
                                                row["二上科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                row["二上科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "2" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["二下科目校排名" + i] = ssr.GetYearRank(subjName);
                                                row["二下科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                row["二下科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "3" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["三上科目校排名" + i] = ssr.GetYearRank(subjName);
                                                row["三上科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                row["三上科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "3" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["三下科目校排名" + i] = ssr.GetYearRank(subjName);
                                                row["三下科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                row["三下科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "4" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["四上科目校排名" + i] = ssr.GetYearRank(subjName);
                                                row["四上科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                row["四上科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "4" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["四下科目校排名" + i] = ssr.GetYearRank(subjName);
                                                row["四下科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                row["四下科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                            }
                                        }
                                        #endregion                

                                        #region 類1
                                        if (ssr.GradeYear == "1" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["一上科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                row["一上科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                row["一上科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "1" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["一下科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                row["一下科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                row["一下科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "2" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["二上科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                row["二上科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                row["二上科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "2" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["二下科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                row["二下科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                row["二下科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "3" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["三上科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                row["三上科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                row["三上科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "3" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["三下科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                row["三下科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                row["三下科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                            }
                                        }

                                        if (ssr.GradeYear == "4" && ssr.Semester == "1")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["四上科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                row["四上科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                row["四上科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                            }
                                        }
                                        if (ssr.GradeYear == "4" && ssr.Semester == "2")
                                        {
                                            for (int i = 1; i <= subjCot; i++)
                                            {
                                                string subjName = row["科目名稱" + i].ToString();
                                                row["四下科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                row["四下科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                row["四下科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                            }
                                        }
                                        #endregion                    
                                    }
                                }
                                if (setting.計算學業成績排名)
                                {
                                    // 處理學業
                                    #region 處理學業
                                    string id1 = studRec.StudentID + "學業";

                                    #region 讀取各學期學業排名
                                    if (StudSemsEntryRatingDict.ContainsKey(studRec.StudentID))
                                    {
                                        foreach (StudSemsEntryRating sser in StudSemsEntryRatingDict[studRec.StudentID])
                                        {
                                            if (sser.GradeYear == "1" && sser.Semester == "1")
                                            {
                                                if (sser.ClassRank.HasValue)
                                                {
                                                    row["一上學業班排名"] = sser.ClassRank.Value;
                                                    if (sser.ClassCount.HasValue)
                                                    {
                                                        row["一上學業班排名母數"] = sser.ClassCount.Value;
                                                        row["一上學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                    }
                                                }

                                                if (sser.DeptRank.HasValue)
                                                {
                                                    row["一上學業科排名"] = sser.DeptRank.Value;
                                                    if (sser.DeptCount.HasValue)
                                                    {
                                                        row["一上學業科排名母數"] = sser.DeptCount.Value;
                                                        row["一上學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                    }
                                                }

                                                if (sser.YearRank.HasValue)
                                                {
                                                    row["一上學業校排名"] = sser.YearRank.Value;
                                                    if (sser.YearCount.HasValue)
                                                    {
                                                        row["一上學業校排名母數"] = sser.YearCount.Value;
                                                        row["一上學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                    }
                                                }

                                                if (sser.Group1Rank.HasValue)
                                                {
                                                    row["一上學業類1排名"] = sser.Group1Rank.Value;
                                                    if (sser.Group1Count.HasValue)
                                                    {
                                                        row["一上學業類1排名母數"] = sser.Group1Count.Value;
                                                        row["一上學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                    }
                                                }
                                            }
                                            if (sser.GradeYear == "1" && sser.Semester == "2")
                                            {
                                                if (sser.ClassRank.HasValue)
                                                {
                                                    row["一下學業班排名"] = sser.ClassRank.Value;
                                                    if (sser.ClassCount.HasValue)
                                                    {
                                                        row["一下學業班排名母數"] = sser.ClassCount.Value;
                                                        row["一下學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                    }
                                                }

                                                if (sser.DeptRank.HasValue)
                                                {
                                                    row["一下學業科排名"] = sser.DeptRank.Value;
                                                    if (sser.DeptCount.HasValue)
                                                    {
                                                        row["一下學業科排名母數"] = sser.DeptCount.Value;
                                                        row["一下學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                    }
                                                }

                                                if (sser.YearRank.HasValue)
                                                {
                                                    row["一下學業校排名"] = sser.YearRank.Value;
                                                    if (sser.YearCount.HasValue)
                                                    {
                                                        row["一下學業校排名母數"] = sser.YearCount.Value;
                                                        row["一下學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                    }
                                                }

                                                if (sser.Group1Rank.HasValue)
                                                {
                                                    row["一下學業類1排名"] = sser.Group1Rank.Value;
                                                    if (sser.Group1Count.HasValue)
                                                    {
                                                        row["一下學業類1排名母數"] = sser.Group1Count.Value;
                                                        row["一下學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                    }
                                                }
                                            }
                                            if (sser.GradeYear == "2" && sser.Semester == "1")
                                            {
                                                if (sser.ClassRank.HasValue)
                                                {
                                                    row["二上學業班排名"] = sser.ClassRank.Value;
                                                    if (sser.ClassCount.HasValue)
                                                    {
                                                        row["二上學業班排名母數"] = sser.ClassCount.Value;
                                                        row["二上學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                    }
                                                }

                                                if (sser.DeptRank.HasValue)
                                                {
                                                    row["二上學業科排名"] = sser.DeptRank.Value;
                                                    if (sser.DeptCount.HasValue)
                                                    {
                                                        row["二上學業科排名母數"] = sser.DeptCount.Value;
                                                        row["二上學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                    }
                                                }

                                                if (sser.YearRank.HasValue)
                                                {
                                                    row["二上學業校排名"] = sser.YearRank.Value;
                                                    if (sser.YearCount.HasValue)
                                                    {
                                                        row["二上學業校排名母數"] = sser.YearCount.Value;
                                                        row["二上學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                    }
                                                }

                                                if (sser.Group1Rank.HasValue)
                                                {
                                                    row["二上學業類1排名"] = sser.Group1Rank.Value;
                                                    if (sser.Group1Count.HasValue)
                                                    {
                                                        row["二上學業類1排名母數"] = sser.Group1Count.Value;
                                                        row["二上學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                    }
                                                }
                                            }
                                            if (sser.GradeYear == "2" && sser.Semester == "2")
                                            {
                                                if (sser.ClassRank.HasValue)
                                                {
                                                    row["二下學業班排名"] = sser.ClassRank.Value;
                                                    if (sser.ClassCount.HasValue)
                                                    {
                                                        row["二下學業班排名母數"] = sser.ClassCount.Value;
                                                        row["二下學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                    }
                                                }

                                                if (sser.DeptRank.HasValue)
                                                {
                                                    row["二下學業科排名"] = sser.DeptRank.Value;
                                                    if (sser.DeptCount.HasValue)
                                                    {
                                                        row["二下學業科排名母數"] = sser.DeptCount.Value;
                                                        row["二下學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                    }
                                                }

                                                if (sser.YearRank.HasValue)
                                                {
                                                    row["二下學業校排名"] = sser.YearRank.Value;
                                                    if (sser.YearCount.HasValue)
                                                    {
                                                        row["二下學業校排名母數"] = sser.YearCount.Value;
                                                        row["二下學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                    }
                                                }

                                                if (sser.Group1Rank.HasValue)
                                                {
                                                    row["二下學業類1排名"] = sser.Group1Rank.Value;
                                                    if (sser.Group1Count.HasValue)
                                                    {
                                                        row["二下學業類1排名母數"] = sser.Group1Count.Value;
                                                        row["二下學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                    }
                                                }
                                            }
                                            if (sser.GradeYear == "3" && sser.Semester == "1")
                                            {
                                                if (sser.ClassRank.HasValue)
                                                {
                                                    row["三上學業班排名"] = sser.ClassRank.Value;
                                                    if (sser.ClassCount.HasValue)
                                                    {
                                                        row["三上學業班排名母數"] = sser.ClassCount.Value;
                                                        row["三上學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                    }
                                                }

                                                if (sser.DeptRank.HasValue)
                                                {
                                                    row["三上學業科排名"] = sser.DeptRank.Value;
                                                    if (sser.DeptCount.HasValue)
                                                    {
                                                        row["三上學業科排名母數"] = sser.DeptCount.Value;
                                                        row["三上學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                    }
                                                }

                                                if (sser.YearRank.HasValue)
                                                {
                                                    row["三上學業校排名"] = sser.YearRank.Value;
                                                    if (sser.YearCount.HasValue)
                                                    {
                                                        row["三上學業校排名母數"] = sser.YearCount.Value;
                                                        row["三上學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                    }
                                                }

                                                if (sser.Group1Rank.HasValue)
                                                {
                                                    row["三上學業類1排名"] = sser.Group1Rank.Value;
                                                    if (sser.Group1Count.HasValue)
                                                    {
                                                        row["三上學業類1排名母數"] = sser.Group1Count.Value;
                                                        row["三上學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                    }
                                                }
                                            }
                                            if (sser.GradeYear == "3" && sser.Semester == "2")
                                            {
                                                if (sser.ClassRank.HasValue)
                                                {
                                                    row["三下學業班排名"] = sser.ClassRank.Value;
                                                    if (sser.ClassCount.HasValue)
                                                    {
                                                        row["三下學業班排名母數"] = sser.ClassCount.Value;
                                                        row["三下學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                    }
                                                }

                                                if (sser.DeptRank.HasValue)
                                                {
                                                    row["三下學業科排名"] = sser.DeptRank.Value;
                                                    if (sser.DeptCount.HasValue)
                                                    {
                                                        row["三下學業科排名母數"] = sser.DeptCount.Value;
                                                        row["三下學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                    }
                                                }

                                                if (sser.YearRank.HasValue)
                                                {
                                                    row["三下學業校排名"] = sser.YearRank.Value;
                                                    if (sser.YearCount.HasValue)
                                                    {
                                                        row["三下學業校排名母數"] = sser.YearCount.Value;
                                                        row["三下學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                    }
                                                }

                                                if (sser.Group1Rank.HasValue)
                                                {
                                                    row["三下學業類1排名"] = sser.Group1Rank.Value;
                                                    if (sser.Group1Count.HasValue)
                                                    {
                                                        row["三下學業類1排名母數"] = sser.Group1Count.Value;
                                                        row["三下學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                    }
                                                }
                                            }
                                            if (sser.GradeYear == "4" && sser.Semester == "1")
                                            {
                                                if (sser.ClassRank.HasValue)
                                                {
                                                    row["四上學業班排名"] = sser.ClassRank.Value;
                                                    if (sser.ClassCount.HasValue)
                                                    {
                                                        row["四上學業班排名母數"] = sser.ClassCount.Value;
                                                        row["四上學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                    }
                                                }

                                                if (sser.DeptRank.HasValue)
                                                {
                                                    row["四上學業科排名"] = sser.DeptRank.Value;
                                                    if (sser.DeptCount.HasValue)
                                                    {
                                                        row["四上學業科排名母數"] = sser.DeptCount.Value;
                                                        row["四上學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                    }
                                                }

                                                if (sser.YearRank.HasValue)
                                                {
                                                    row["四上學業校排名"] = sser.YearRank.Value;
                                                    if (sser.YearCount.HasValue)
                                                    {
                                                        row["四上學業校排名母數"] = sser.YearCount.Value;
                                                        row["四上學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                    }
                                                }

                                                if (sser.Group1Rank.HasValue)
                                                {
                                                    row["四上學業類1排名"] = sser.Group1Rank.Value;
                                                    if (sser.Group1Count.HasValue)
                                                    {
                                                        row["四上學業類1排名母數"] = sser.Group1Count.Value;
                                                        row["四上學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                    }
                                                }
                                            }
                                            if (sser.GradeYear == "4" && sser.Semester == "2")
                                            {
                                                if (sser.ClassRank.HasValue)
                                                {
                                                    row["四下學業班排名"] = sser.ClassRank.Value;
                                                    if (sser.ClassCount.HasValue)
                                                    {
                                                        row["四下學業班排名母數"] = sser.ClassCount.Value;
                                                        row["四下學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                    }
                                                }

                                                if (sser.DeptRank.HasValue)
                                                {
                                                    row["四下學業科排名"] = sser.DeptRank.Value;
                                                    if (sser.DeptCount.HasValue)
                                                    {
                                                        row["四下學業科排名母數"] = sser.DeptCount.Value;
                                                        row["四下學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                    }
                                                }

                                                if (sser.YearRank.HasValue)
                                                {
                                                    row["四下學業校排名"] = sser.YearRank.Value;
                                                    if (sser.YearCount.HasValue)
                                                    {
                                                        row["四下學業校排名母數"] = sser.YearCount.Value;
                                                        row["四下學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                    }
                                                }

                                                if (sser.Group1Rank.HasValue)
                                                {
                                                    row["四下學業類1排名"] = sser.Group1Rank.Value;
                                                    if (sser.Group1Count.HasValue)
                                                    {
                                                        row["四下學業類1排名母數"] = sser.Group1Count.Value;
                                                        row["四下學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                    #endregion

                                    if (selectScore.ContainsKey(id1))
                                    {
                                        if (selectScore[id1].gsScore11.HasValue)
                                            row["一上學業成績"] = Utility.NoRound(selectScore[id1].gsScore11.Value);

                                        if (selectScore[id1].gsScore12.HasValue)
                                            row["一下學業成績"] = Utility.NoRound(selectScore[id1].gsScore12.Value);

                                        if (selectScore[id1].gsScore21.HasValue)
                                            row["二上學業成績"] = Utility.NoRound(selectScore[id1].gsScore21.Value);

                                        if (selectScore[id1].gsScore22.HasValue)
                                            row["二下學業成績"] = Utility.NoRound(selectScore[id1].gsScore22.Value);

                                        if (selectScore[id1].gsScore31.HasValue)
                                            row["三上學業成績"] = Utility.NoRound(selectScore[id1].gsScore31.Value);

                                        if (selectScore[id1].gsScore32.HasValue)
                                            row["三下學業成績"] = Utility.NoRound(selectScore[id1].gsScore32.Value);

                                        if (selectScore[id1].gsScore41.HasValue)
                                            row["四上學業成績"] = Utility.NoRound(selectScore[id1].gsScore41.Value);

                                        if (selectScore[id1].gsScore42.HasValue)
                                            row["四下學業成績"] = Utility.NoRound(selectScore[id1].gsScore42.Value);


                                        row["學業平均"] = Utility.NoRound(selectScore[id1].avgScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key2 = "學業成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1].avgScore) + 1;
                                                row["學業平均班排名"] = rr;
                                                row["學業平均班排名母數"] = ranks[key2].Count;
                                                row["學業平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1].avgScore) + 1;
                                                row["學業平均科排名"] = rr;
                                                row["學業平均科排名母數"] = ranks[key2].Count;
                                                row["學業平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業成績平均全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1].avgScore) + 1;
                                                row["學業平均校排名"] = rr;
                                                row["學業平均校排名母數"] = ranks[key2].Count;
                                                row["學業平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key2 = "學業成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1].avgScoreC1) + 1;
                                                    row["學業平均類別一排名"] = rr;
                                                    row["學業平均類別一排名母數"] = ranks[key2].Count;
                                                    row["學業平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key2 = "學業成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1].avgScoreC2) + 1;
                                                    row["學業平均類別二排名"] = rr;
                                                    row["學業平均類別二排名母數"] = ranks[key2].Count;
                                                    row["學業平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                }
                                            }
                                        }
                                        row["學業總分"] = Utility.NoRound(selectScore[id1].sumScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key1 = "學業成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業總分班排名"] = ranks[key1].IndexOf(selectScore[id1].sumScore) + 1;
                                                row["學業總分班排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業總分科排名"] = ranks[key1].IndexOf(selectScore[id1].sumScore) + 1;
                                                row["學業總分科排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業成績總分全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業總分校排名"] = ranks[key1].IndexOf(selectScore[id1].sumScore) + 1;
                                                row["學業總分校排名母數"] = ranks[key1].Count;
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key1 = "學業成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1].sumScoreC1) + 1;
                                                    row["學業總分類別一排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key1 = "學業成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1].sumScoreC2) + 1;
                                                    row["學業總分類別二排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                        }
                                    }
                                    #endregion 處理學業

                                    // 處理學業原始
                                    #region 處理學業原始
                                    string id1_5 = studRec.StudentID + "學業原始";

                                    if (selectScore.ContainsKey(id1_5))
                                    {
                                        if (selectScore[id1_5].gsScore11.HasValue)
                                            row["一上學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore11.Value);

                                        if (selectScore[id1_5].gsScore12.HasValue)
                                            row["一下學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore12.Value);

                                        if (selectScore[id1_5].gsScore21.HasValue)
                                            row["二上學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore21.Value);

                                        if (selectScore[id1_5].gsScore22.HasValue)
                                            row["二下學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore22.Value);

                                        if (selectScore[id1_5].gsScore31.HasValue)
                                            row["三上學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore31.Value);

                                        if (selectScore[id1_5].gsScore32.HasValue)
                                            row["三下學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore32.Value);

                                        if (selectScore[id1_5].gsScore41.HasValue)
                                            row["四上學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore41.Value);

                                        if (selectScore[id1_5].gsScore42.HasValue)
                                            row["四下學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore42.Value);


                                        row["學業原始平均"] = Utility.NoRound(selectScore[id1_5].avgScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key2 = "學業原始成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_5].avgScore) + 1;
                                                row["學業原始平均班排名"] = rr;
                                                row["學業原始平均班排名母數"] = ranks[key2].Count;
                                                row["學業原始平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業原始成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_5].avgScore) + 1;
                                                row["學業原始平均科排名"] = rr;
                                                row["學業原始平均科排名母數"] = ranks[key2].Count;
                                                row["學業原始平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業原始成績平均全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_5].avgScore) + 1;
                                                row["學業原始平均校排名"] = rr;
                                                row["學業原始平均校排名母數"] = ranks[key2].Count;
                                                row["學業原始平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key2 = "學業原始成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1_5].avgScoreC1) + 1;
                                                    row["學業原始平均類別一排名"] = rr;
                                                    row["學業原始平均類別一排名母數"] = ranks[key2].Count;
                                                    row["學業原始平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key2 = "學業原始成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1_5].avgScoreC2) + 1;
                                                    row["學業原始平均類別二排名"] = rr;
                                                    row["學業原始平均類別二排名母數"] = ranks[key2].Count;
                                                    row["學業原始平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                }
                                            }
                                        }
                                        row["學業原始總分"] = Utility.NoRound(selectScore[id1_5].sumScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key1 = "學業原始成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業原始總分班排名"] = ranks[key1].IndexOf(selectScore[id1_5].sumScore) + 1;
                                                row["學業原始總分班排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業原始成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業原始總分科排名"] = ranks[key1].IndexOf(selectScore[id1_5].sumScore) + 1;
                                                row["學業原始總分科排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業原始成績總分全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業原始總分校排名"] = ranks[key1].IndexOf(selectScore[id1_5].sumScore) + 1;
                                                row["學業原始總分校排名母數"] = ranks[key1].Count;
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key1 = "學業原始成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業原始總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1_5].sumScoreC1) + 1;
                                                    row["學業原始總分類別一排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key1 = "學業原始成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業原始總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1_5].sumScoreC2) + 1;
                                                    row["學業原始總分類別二排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                        }
                                    }
                                    #endregion 處理學業原始

                                    // 處理學業體育
                                    #region 處理學業體育
                                    string id1_1 = studRec.StudentID + "學業體育";

                                    if (selectScore.ContainsKey(id1_1))
                                    {
                                        if (selectScore[id1_1].gsScore11.HasValue)
                                            row["一上學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore11.Value);

                                        if (selectScore[id1_1].gsScore12.HasValue)
                                            row["一下學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore12.Value);

                                        if (selectScore[id1_1].gsScore21.HasValue)
                                            row["二上學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore21.Value);

                                        if (selectScore[id1_1].gsScore22.HasValue)
                                            row["二下學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore22.Value);

                                        if (selectScore[id1_1].gsScore31.HasValue)
                                            row["三上學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore31.Value);

                                        if (selectScore[id1_1].gsScore32.HasValue)
                                            row["三下學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore32.Value);

                                        if (selectScore[id1_1].gsScore41.HasValue)
                                            row["四上學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore41.Value);

                                        if (selectScore[id1_1].gsScore42.HasValue)
                                            row["四下學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore42.Value);


                                        row["學業體育平均"] = Utility.NoRound(selectScore[id1_1].avgScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key2 = "學業體育成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_1].avgScore) + 1;
                                                row["學業體育平均班排名"] = rr;
                                                row["學業體育平均班排名母數"] = ranks[key2].Count;
                                                row["學業體育平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業體育成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_1].avgScore) + 1;
                                                row["學業體育平均科排名"] = rr;
                                                row["學業體育平均科排名母數"] = ranks[key2].Count;
                                                row["學業體育平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業體育成績平均全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_1].avgScore) + 1;
                                                row["學業體育平均校排名"] = rr;
                                                row["學業體育平均校排名母數"] = ranks[key2].Count;
                                                row["學業體育平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key2 = "學業體育成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1_1].avgScoreC1) + 1;
                                                    row["學業體育平均類別一排名"] = rr;
                                                    row["學業體育平均類別一排名母數"] = ranks[key2].Count;
                                                    row["學業體育平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key2 = "學業體育成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1_1].avgScoreC2) + 1;
                                                    row["學業體育平均類別二排名"] = rr;
                                                    row["學業體育平均類別二排名母數"] = ranks[key2].Count;
                                                    row["學業體育平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                }
                                            }
                                        }
                                        row["學業體育總分"] = Utility.NoRound(selectScore[id1_1].sumScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key1 = "學業體育成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業體育總分班排名"] = ranks[key1].IndexOf(selectScore[id1_1].sumScore) + 1;
                                                row["學業體育總分班排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業體育成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業體育總分科排名"] = ranks[key1].IndexOf(selectScore[id1_1].sumScore) + 1;
                                                row["學業體育總分科排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業體育成績總分全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業體育總分校排名"] = ranks[key1].IndexOf(selectScore[id1_1].sumScore) + 1;
                                                row["學業體育總分校排名母數"] = ranks[key1].Count;
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key1 = "學業體育成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業體育總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1_1].sumScoreC1) + 1;
                                                    row["學業體育總分類別一排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key1 = "學業體育成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業體育總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1_1].sumScoreC2) + 1;
                                                    row["學業體育總分類別二排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                        }
                                    }
                                    #endregion 處理學業體育

                                    // 處理學業健康與護理
                                    #region 處理學業健康與護理
                                    string id1_2 = studRec.StudentID + "學業健康與護理";

                                    if (selectScore.ContainsKey(id1_2))
                                    {
                                        if (selectScore[id1_2].gsScore11.HasValue)
                                            row["一上學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore11.Value);

                                        if (selectScore[id1_2].gsScore12.HasValue)
                                            row["一下學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore12.Value);

                                        if (selectScore[id1_2].gsScore21.HasValue)
                                            row["二上學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore21.Value);

                                        if (selectScore[id1_2].gsScore22.HasValue)
                                            row["二下學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore22.Value);

                                        if (selectScore[id1_2].gsScore31.HasValue)
                                            row["三上學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore31.Value);

                                        if (selectScore[id1_2].gsScore32.HasValue)
                                            row["三下學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore32.Value);

                                        if (selectScore[id1_2].gsScore41.HasValue)
                                            row["四上學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore41.Value);

                                        if (selectScore[id1_2].gsScore42.HasValue)
                                            row["四下學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore42.Value);


                                        row["學業健康與護理平均"] = Utility.NoRound(selectScore[id1_2].avgScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key2 = "學業健康與護理成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_2].avgScore) + 1;
                                                row["學業健康與護理平均班排名"] = rr;
                                                row["學業健康與護理平均班排名母數"] = ranks[key2].Count;
                                                row["學業健康與護理平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業健康與護理成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_2].avgScore) + 1;
                                                row["學業健康與護理平均科排名"] = rr;
                                                row["學業健康與護理平均科排名母數"] = ranks[key2].Count;
                                                row["學業健康與護理平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業健康與護理成績平均全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_2].avgScore) + 1;
                                                row["學業健康與護理平均校排名"] = rr;
                                                row["學業健康與護理平均校排名母數"] = ranks[key2].Count;
                                                row["學業健康與護理平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key2 = "學業健康與護理成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1_2].avgScoreC1) + 1;
                                                    row["學業健康與護理平均類別一排名"] = rr;
                                                    row["學業健康與護理平均類別一排名母數"] = ranks[key2].Count;
                                                    row["學業健康與護理平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key2 = "學業健康與護理成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1_2].avgScoreC2) + 1;
                                                    row["學業健康與護理平均類別二排名"] = rr;
                                                    row["學業健康與護理平均類別二排名母數"] = ranks[key2].Count;
                                                    row["學業健康與護理平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                }
                                            }
                                        }
                                        row["學業健康與護理總分"] = Utility.NoRound(selectScore[id1_2].sumScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key1 = "學業健康與護理成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業健康與護理總分班排名"] = ranks[key1].IndexOf(selectScore[id1_2].sumScore) + 1;
                                                row["學業健康與護理總分班排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業健康與護理成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業健康與護理總分科排名"] = ranks[key1].IndexOf(selectScore[id1_2].sumScore) + 1;
                                                row["學業健康與護理總分科排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業健康與護理成績總分全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業健康與護理總分校排名"] = ranks[key1].IndexOf(selectScore[id1_2].sumScore) + 1;
                                                row["學業健康與護理總分校排名母數"] = ranks[key1].Count;
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key1 = "學業健康與護理成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業健康與護理總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1_2].sumScoreC1) + 1;
                                                    row["學業健康與護理總分類別一排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key1 = "學業健康與護理成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業健康與護理總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1_2].sumScoreC2) + 1;
                                                    row["學業健康與護理總分類別二排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                        }
                                    }
                                    #endregion 處理學業健康與護理

                                    // 處理學業國防通識
                                    #region 處理學業國防通識
                                    string id1_3 = studRec.StudentID + "學業國防通識";

                                    if (selectScore.ContainsKey(id1_3))
                                    {
                                        if (selectScore[id1_3].gsScore11.HasValue)
                                            row["一上學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore11.Value);

                                        if (selectScore[id1_3].gsScore12.HasValue)
                                            row["一下學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore12.Value);

                                        if (selectScore[id1_3].gsScore21.HasValue)
                                            row["二上學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore21.Value);

                                        if (selectScore[id1_3].gsScore22.HasValue)
                                            row["二下學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore22.Value);

                                        if (selectScore[id1_3].gsScore31.HasValue)
                                            row["三上學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore31.Value);

                                        if (selectScore[id1_3].gsScore32.HasValue)
                                            row["三下學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore32.Value);

                                        if (selectScore[id1_3].gsScore41.HasValue)
                                            row["四上學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore41.Value);

                                        if (selectScore[id1_3].gsScore42.HasValue)
                                            row["四下學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore42.Value);


                                        row["學業國防通識平均"] = Utility.NoRound(selectScore[id1_3].avgScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key2 = "學業國防通識成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_3].avgScore) + 1;
                                                row["學業國防通識平均班排名"] = rr;
                                                row["學業國防通識平均班排名母數"] = ranks[key2].Count;
                                                row["學業國防通識平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業國防通識成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_3].avgScore) + 1;
                                                row["學業國防通識平均科排名"] = rr;
                                                row["學業國防通識平均科排名母數"] = ranks[key2].Count;
                                                row["學業國防通識平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業國防通識成績平均全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_3].avgScore) + 1;
                                                row["學業國防通識平均校排名"] = rr;
                                                row["學業國防通識平均校排名母數"] = ranks[key2].Count;
                                                row["學業國防通識平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key2 = "學業國防通識成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1_3].avgScoreC1) + 1;
                                                    row["學業國防通識平均類別一排名"] = rr;
                                                    row["學業國防通識平均類別一排名母數"] = ranks[key2].Count;
                                                    row["學業國防通識平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key2 = "學業國防通識成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1_3].avgScoreC2) + 1;
                                                    row["學業國防通識平均類別二排名"] = rr;
                                                    row["學業國防通識平均類別二排名母數"] = ranks[key2].Count;
                                                    row["學業國防通識平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                }
                                            }
                                        }
                                        row["學業國防通識總分"] = Utility.NoRound(selectScore[id1_3].sumScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key1 = "學業國防通識成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業國防通識總分班排名"] = ranks[key1].IndexOf(selectScore[id1_3].sumScore) + 1;
                                                row["學業國防通識總分班排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業國防通識成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業國防通識總分科排名"] = ranks[key1].IndexOf(selectScore[id1_3].sumScore) + 1;
                                                row["學業國防通識總分科排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業國防通識成績總分全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業國防通識總分校排名"] = ranks[key1].IndexOf(selectScore[id1_3].sumScore) + 1;
                                                row["學業國防通識總分校排名母數"] = ranks[key1].Count;
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key1 = "學業國防通識成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業國防通識總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1_3].sumScoreC1) + 1;
                                                    row["學業國防通識總分類別一排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key1 = "學業國防通識成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業國防通識總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1_3].sumScoreC2) + 1;
                                                    row["學業國防通識總分類別二排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                        }
                                    }
                                    #endregion 處理學業國防通識

                                    // 處理學業實習科目
                                    #region 處理學業實習科目
                                    string id1_4 = studRec.StudentID + "學業實習科目";

                                    if (selectScore.ContainsKey(id1_4))
                                    {
                                        if (selectScore[id1_4].gsScore11.HasValue)
                                            row["一上學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore11.Value);

                                        if (selectScore[id1_4].gsScore12.HasValue)
                                            row["一下學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore12.Value);

                                        if (selectScore[id1_4].gsScore21.HasValue)
                                            row["二上學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore21.Value);

                                        if (selectScore[id1_4].gsScore22.HasValue)
                                            row["二下學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore22.Value);

                                        if (selectScore[id1_4].gsScore31.HasValue)
                                            row["三上學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore31.Value);

                                        if (selectScore[id1_4].gsScore32.HasValue)
                                            row["三下學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore32.Value);

                                        if (selectScore[id1_4].gsScore41.HasValue)
                                            row["四上學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore41.Value);

                                        if (selectScore[id1_4].gsScore42.HasValue)
                                            row["四下學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore42.Value);


                                        row["學業實習科目平均"] = Utility.NoRound(selectScore[id1_4].avgScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key2 = "學業實習科目成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_4].avgScore) + 1;
                                                row["學業實習科目平均班排名"] = rr;
                                                row["學業實習科目平均班排名母數"] = ranks[key2].Count;
                                                row["學業實習科目平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業實習科目成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_4].avgScore) + 1;
                                                row["學業實習科目平均科排名"] = rr;
                                                row["學業實習科目平均科排名母數"] = ranks[key2].Count;
                                                row["學業實習科目平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            key2 = "學業實習科目成績平均全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key2))
                                            {
                                                int rr = ranks[key2].IndexOf(selectScore[id1_4].avgScore) + 1;
                                                row["學業實習科目平均校排名"] = rr;
                                                row["學業實習科目平均校排名母數"] = ranks[key2].Count;
                                                row["學業實習科目平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key2 = "學業實習科目成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1_4].avgScoreC1) + 1;
                                                    row["學業實習科目平均類別一排名"] = rr;
                                                    row["學業實習科目平均類別一排名母數"] = ranks[key2].Count;
                                                    row["學業實習科目平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key2 = "學業實習科目成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    int rr = ranks[key2].IndexOf(selectScore[id1_4].avgScoreC2) + 1;
                                                    row["學業實習科目平均類別二排名"] = rr;
                                                    row["學業實習科目平均類別二排名母數"] = ranks[key2].Count;
                                                    row["學業實習科目平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                }
                                            }
                                        }
                                        row["學業實習科目總分"] = Utility.NoRound(selectScore[id1_4].sumScore);
                                        if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                        {
                                            string key1 = "學業實習科目成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業實習科目總分班排名"] = ranks[key1].IndexOf(selectScore[id1_4].sumScore) + 1;
                                                row["學業實習科目總分班排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業實習科目成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業實習科目總分科排名"] = ranks[key1].IndexOf(selectScore[id1_4].sumScore) + 1;
                                                row["學業實習科目總分科排名母數"] = ranks[key1].Count;
                                            }

                                            key1 = "學業實習科目成績總分全校排名" + gradeyear + "^^^";
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["學業實習科目總分校排名"] = ranks[key1].IndexOf(selectScore[id1_4].sumScore) + 1;
                                                row["學業實習科目總分校排名母數"] = ranks[key1].Count;
                                            }

                                            if (studRec.Fields.ContainsKey("tag1"))
                                            {
                                                key1 = "學業實習科目成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業實習科目總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1_4].sumScoreC1) + 1;
                                                    row["學業實習科目總分類別一排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                            if (studRec.Fields.ContainsKey("tag2"))
                                            {
                                                key1 = "學業實習科目成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["學業實習科目總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1_4].sumScoreC2) + 1;
                                                    row["學業實習科目總分類別二排名母數"] = ranks[key1].Count;
                                                }
                                            }
                                        }
                                    }
                                    #endregion 處理學業實習科目

                                }
                                // 處理總計
                                #region 處理總計
                                string id2 = studRec.StudentID + "總計成績";
                                if (selectScore.ContainsKey(id2))
                                {

                                    row["總計加權平均"] = Utility.NoRound(selectScore[id2].avgScoreA);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key4 = "總計加權平均班排名" + studRec.RefClass.ClassID;
                                        if (ranks.ContainsKey(key4))
                                        {
                                            row["總計加權平均班排名"] = ranks[key4].IndexOf(selectScore[id2].avgScoreA) + 1;
                                            row["總計加權平均班排名母數"] = ranks[key4].Count;
                                        }
                                        if (ranks.ContainsKey(key4))
                                        {
                                            row["總計加權平均科排名"] = ranks[key4].IndexOf(selectScore[id2].avgScoreA) + 1;
                                            row["總計加權平均科排名母數"] = ranks[key4].Count;
                                        }
                                        if (ranks.ContainsKey(key4))
                                        {
                                            row["總計加權平均校排名"] = ranks[key4].IndexOf(selectScore[id2].avgScoreA) + 1;
                                            row["總計加權平均校排名母數"] = ranks[key4].Count;
                                        }
                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            if (ranks.ContainsKey(key4))
                                            {
                                                row["總計加權平均類別一"] = Utility.NoRound(selectScore[id2].avgScoreAC1);
                                                row["總計加權平均類別一排名"] = ranks[key4].IndexOf(selectScore[id2].avgScoreAC1) + 1;
                                                row["總計加權平均類別一排名母數"] = ranks[key4].Count;
                                            }
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            if (ranks.ContainsKey(key4))
                                            {
                                                row["總計加權平均類別二"] = selectScore[id2].avgScoreAC2;
                                                row["總計加權平均類別二排名"] = ranks[key4].IndexOf(selectScore[id2].avgScoreAC2) + 1;
                                                row["總計加權平均類別二排名母數"] = ranks[key4].Count;
                                            }
                                        }
                                    }

                                    row["總計加權總分"] = Utility.NoRound(selectScore[id2].sumScoreA);

                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key2 = "總計加權總分班排名" + studRec.RefClass.ClassID;
                                        if (ranks.ContainsKey(key2))
                                        {
                                            row["總計加權總分班排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreA) + 1; ;
                                            row["總計加權總分班排名母數"] = ranks[key2].Count;
                                        }

                                        key2 = "總計加權總分科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key2))
                                        {
                                            row["總計加權總分科排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreA) + 1; ;
                                            row["總計加權總分科排名母數"] = ranks[key2].Count;
                                        }

                                        key2 = "總計加權總分全校排名" + gradeyear;
                                        if (ranks.ContainsKey(key2))
                                        {
                                            row["總計加權總分校排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreA) + 1; ;
                                            row["總計加權總分校排名母數"] = ranks[key2].Count;
                                        }

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key2 = "總計加權總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                row["總計加權總分類別一"] = Utility.NoRound(selectScore[id2].sumScoreAC1);
                                                row["總計加權總分類別一排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreAC1) + 1; ;
                                                row["總計加權總分類別一排名母數"] = ranks[key2].Count;
                                            }
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key2 = "總計加權總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key2))
                                            {
                                                row["總計加權總分類別二"] = Utility.NoRound(selectScore[id2].sumScoreAC2);
                                                row["總計加權總分類別二排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreAC2) + 1; ;
                                                row["總計加權總分類別二排名母數"] = ranks[key2].Count;
                                            }
                                        }
                                    }

                                    row["總計平均"] = Utility.NoRound(selectScore[id2].avgScore);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key3 = "總計平均班排名" + studRec.RefClass.ClassID;
                                        if (ranks.ContainsKey(key3))
                                        {
                                            row["總計平均班排名"] = ranks[key3].IndexOf(selectScore[id2].avgScore) + 1;
                                            row["總計平均班排名母數"] = ranks[key3].Count;
                                        }

                                        key3 = "總計平均科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key3))
                                        {
                                            row["總計平均科排名"] = ranks[key3].IndexOf(selectScore[id2].avgScore) + 1;
                                            row["總計平均科排名母數"] = ranks[key3].Count;
                                        }

                                        key3 = "總計平均全校排名" + gradeyear;
                                        if (ranks.ContainsKey(key3))
                                        {
                                            row["總計平均校排名"] = ranks[key3].IndexOf(selectScore[id2].avgScore) + 1;
                                            row["總計平均校排名母數"] = ranks[key3].Count;
                                        }

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key3 = "總計平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key3))
                                            {
                                                row["總計平均類別一"] = Utility.NoRound(selectScore[id2].avgScoreC1);
                                                row["總計平均類別一排名"] = ranks[key3].IndexOf(selectScore[id2].avgScoreC1) + 1;
                                                row["總計平均類別一排名母數"] = ranks[key3].Count;
                                            }
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key3 = "總計平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key3))
                                            {
                                                row["總計平均類別二"] = Utility.NoRound(selectScore[id2].avgScoreC2);
                                                row["總計平均類別二排名"] = ranks[key3].IndexOf(selectScore[id2].avgScoreC2) + 1;
                                                row["總計平均類別二排名母數"] = ranks[key3].Count;
                                            }
                                        }
                                    }
                                    row["總計總分"] = Utility.NoRound(selectScore[id2].sumScore);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key1 = "總計總分班排名" + studRec.RefClass.ClassID;
                                        if (ranks.ContainsKey(key1))
                                        {
                                            row["總計總分班排名"] = ranks[key1].IndexOf(selectScore[id2].sumScore) + 1;
                                            row["總計總分班排名母數"] = ranks[key1].Count;
                                        }

                                        key1 = "總計總分科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                        {
                                            row["總計總分科排名"] = ranks[key1].IndexOf(selectScore[id2].sumScore) + 1;
                                            row["總計總分科排名母數"] = ranks[key1].Count;
                                        }

                                        key1 = "總計總分全校排名" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                        {
                                            row["總計總分校排名"] = ranks[key1].IndexOf(selectScore[id2].sumScore) + 1;
                                            row["總計總分校排名母數"] = ranks[key1].Count;
                                        }

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key1 = "總計總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["總計總分類別一"] = Utility.NoRound(selectScore[id2].sumScoreC1);
                                                row["總計總分類別一排名"] = ranks[key1].IndexOf(selectScore[id2].sumScoreC1) + 1;
                                                row["總計總分類別一排名母數"] = ranks[key1].Count;
                                            }
                                        }
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key1 = "總計總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                            {
                                                row["總計總分類別二"] = Utility.NoRound(selectScore[id2].sumScoreC2);
                                                row["總計總分類別二排名"] = ranks[key1].IndexOf(selectScore[id2].sumScoreC2) + 1;
                                                row["總計總分類別二排名母數"] = ranks[key1].Count;
                                            }
                                        }
                                    }
                                }
                                #endregion 處理總計

                                // 處理科目原始成績加權平均
                                #region 處理科目原始成績加權平均
                                string key5_name = "科目原始成績加權平均";
                                string key7_name = "篩選科目原始成績加權平均";
                                for (int g = 1; g <= 4; g++)
                                {
                                    string key5 = "", key7 = "";
                                    string gsS = "";
                                    for (int s = 1; s <= 2; s++)
                                    {
                                        if (g == 1 && s == 1) gsS = "一上";
                                        if (g == 1 && s == 2) gsS = "一下";
                                        if (g == 2 && s == 1) gsS = "二上";
                                        if (g == 2 && s == 2) gsS = "二下";
                                        if (g == 3 && s == 1) gsS = "三上";
                                        if (g == 3 && s == 2) gsS = "三下";
                                        if (g == 4 && s == 1) gsS = "四上";
                                        if (g == 4 && s == 2) gsS = "四下";
                                        key5 = studRec.StudentID + gsS + key5_name;
                                        key7 = studRec.StudentID + gsS + key7_name;

                                        string id5 = studRec.StudentID + gsS + key5_name;
                                        string id7 = studRec.StudentID + gsS + key7_name;

                                        #region 處理科目原始成績
                                        if (selectScore.ContainsKey(id5))
                                        {
                                            row[gsS + "科目原始成績加權平均"] = Utility.NoRound(selectScore[id5].avgScoreA);
                                            if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                            {
                                                string key5a = gsS + "科目原始成績加權平均班排名" + studRec.RefClass.ClassID;
                                                if (ranks.ContainsKey(key5a))
                                                {
                                                    int rr = ranks[key5a].IndexOf(selectScore[id5].avgScoreA) + 1;
                                                    row[gsS + "科目原始成績加權平均班排名"] = rr;
                                                    row[gsS + "科目原始成績加權平均班排名母數"] = ranks[key5a].Count;
                                                    row[gsS + "科目原始成績加權平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                }

                                                key5a = gsS + "科目原始成績加權平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key5a))
                                                {
                                                    int rr = ranks[key5a].IndexOf(selectScore[id5].avgScoreA) + 1;
                                                    row[gsS + "科目原始成績加權平均科排名"] = rr;
                                                    row[gsS + "科目原始成績加權平均科排名母數"] = ranks[key5a].Count;
                                                    row[gsS + "科目原始成績加權平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                }

                                                key5a = gsS + "科目原始成績加權平均校排名" + gradeyear;
                                                if (ranks.ContainsKey(key5a))
                                                {
                                                    int rr = ranks[key5a].IndexOf(selectScore[id5].avgScoreA) + 1;
                                                    row[gsS + "科目原始成績加權平均校排名"] = rr;
                                                    row[gsS + "科目原始成績加權平均校排名母數"] = ranks[key5a].Count;
                                                    row[gsS + "科目原始成績加權平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                }
                                                if (studRec.Fields.ContainsKey("tag1"))
                                                {
                                                    key5a = gsS + "科目原始成績加權平均類別一" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key5a))
                                                    {
                                                        int rr = ranks[key5a].IndexOf(selectScore[id5].avgScoreAC1) + 1;
                                                        row[gsS + "科目原始成績加權平均類別一"] = Utility.NoRound(selectScore[id5].avgScoreAC1);
                                                        row[gsS + "科目原始成績加權平均類別一排名"] = rr;
                                                        row[gsS + "科目原始成績加權平均類別一排名母數"] = ranks[key5a].Count;
                                                        row[gsS + "科目原始成績加權平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                    }
                                                }

                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    key5a = gsS + "科目原始成績加權平均類別二" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key5a))
                                                    {
                                                        int rr = ranks[key5a].IndexOf(selectScore[id5].avgScoreAC2) + 1;
                                                        row[gsS + "科目原始成績加權平均類別二"] = Utility.NoRound(selectScore[id5].avgScoreAC2);
                                                        row[gsS + "科目原始成績加權平均類別二排名"] = rr;
                                                        row[gsS + "科目原始成績加權平均類別二排名母數"] = ranks[key5a].Count;
                                                        row[gsS + "科目原始成績加權平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                    }
                                                }
                                            }
                                        }
                                        #endregion 處理科目原始成績

                                        #region 處理篩選科目原始成績
                                        if (selectScore.ContainsKey(id7))
                                        {
                                            row[gsS + "篩選科目原始成績加權平均"] = Utility.NoRound(selectScore[id7].avgScoreA);

                                            if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                            {
                                                string key7a = gsS + "篩選科目原始成績加權平均班排名" + studRec.RefClass.ClassID;
                                                if (ranks.ContainsKey(key7a))
                                                {
                                                    int rr = ranks[key7a].IndexOf(selectScore[id7].avgScoreA) + 1;
                                                    row[gsS + "篩選科目原始成績加權平均班排名"] = rr;
                                                    row[gsS + "篩選科目原始成績加權平均班排名母數"] = ranks[key7a].Count;
                                                    row[gsS + "篩選科目原始成績加權平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
                                                }

                                                key7a = gsS + "篩選科目原始成績加權平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key7a))
                                                {
                                                    int rr = ranks[key7a].IndexOf(selectScore[id7].avgScoreA) + 1;
                                                    row[gsS + "篩選科目原始成績加權平均科排名"] = rr;
                                                    row[gsS + "篩選科目原始成績加權平均科排名母數"] = ranks[key7a].Count;
                                                    row[gsS + "篩選科目原始成績加權平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
                                                }

                                                key7a = gsS + "篩選科目原始成績加權平均校排名" + gradeyear;
                                                if (ranks.ContainsKey(key7a))
                                                {
                                                    int rr = ranks[key7a].IndexOf(selectScore[id7].avgScoreA) + 1;
                                                    row[gsS + "篩選科目原始成績加權平均校排名"] = rr;
                                                    row[gsS + "篩選科目原始成績加權平均校排名母數"] = ranks[key7a].Count;
                                                    row[gsS + "篩選科目原始成績加權平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
                                                }

                                                if (studRec.Fields.ContainsKey("tag1"))
                                                {
                                                    key7a = gsS + "篩選科目原始成績加權平均類別一" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key7a))
                                                    {
                                                        int rr = ranks[key7a].IndexOf(selectScore[id7].avgScoreAC1) + 1;
                                                        row[gsS + "篩選科目原始成績加權平均類別一"] = Utility.NoRound(selectScore[id7].avgScoreAC1);
                                                        row[gsS + "篩選科目原始成績加權平均類別一排名"] = rr;
                                                        row[gsS + "篩選科目原始成績加權平均類別一排名母數"] = ranks[key7a].Count;
                                                        row[gsS + "篩選科目原始成績加權平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
                                                    }
                                                }

                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    key7a = gsS + "篩選科目原始成績加權平均類別二" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key7a))
                                                    {
                                                        int rr = ranks[key7a].IndexOf(selectScore[id7].avgScoreAC2) + 1;
                                                        row[gsS + "篩選科目原始成績加權平均類別二"] = Utility.NoRound(selectScore[id7].avgScoreAC2);
                                                        row[gsS + "篩選科目原始成績加權平均類別二排名"] = rr;
                                                        row[gsS + "篩選科目原始成績加權平均類別二排名母數"] = ranks[key7a].Count;
                                                        row[gsS + "篩選科目原始成績加權平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
                                                    }
                                                }
                                            }
                                        }
                                        #endregion 處理篩選科目原始成績
                                    }
                                }
                                #endregion 處理科目原始成績加權平均

                                // 處理科目原始成績加權平均平均
                                #region 處理科目原始成績加權平均平均
                                string key6_name = "科目原始成績加權平均平均";
                                string key8_name = "篩選科目原始成績加權平均平均";
                                string key6 = studRec.StudentID + key6_name;
                                string key8 = studRec.StudentID + key8_name;

                                string id6 = studRec.StudentID + key6_name;
                                string id8 = studRec.StudentID + key8_name;

                                #region 處理科目原始成績
                                if (selectScore.ContainsKey(id6))
                                {

                                    row["科目原始成績加權平均平均"] = Utility.NoRound(selectScore[id6].avgScoreA);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key6a = "科目原始成績加權平均平均班排名" + studRec.RefClass.ClassID;
                                        if (ranks.ContainsKey(key6a))
                                        {
                                            int rr = ranks[key6a].IndexOf(selectScore[id6].avgScoreA) + 1;
                                            row["科目原始成績加權平均平均班排名"] = rr;
                                            row["科目原始成績加權平均平均班排名母數"] = ranks[key6a].Count;
                                            row["科目原始成績加權平均平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                        }

                                        key6a = "科目原始成績加權平均平均科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key6a))
                                        {
                                            int rr = ranks[key6a].IndexOf(selectScore[id6].avgScoreA) + 1;
                                            row["科目原始成績加權平均平均科排名"] = rr;
                                            row["科目原始成績加權平均平均科排名母數"] = ranks[key6a].Count;
                                            row["科目原始成績加權平均平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                        }

                                        key6a = "科目原始成績加權平均平均校排名" + gradeyear;
                                        if (ranks.ContainsKey(key6a))
                                        {
                                            int rr = ranks[key6a].IndexOf(selectScore[id6].avgScoreA) + 1;
                                            row["科目原始成績加權平均平均校排名"] = rr;
                                            row["科目原始成績加權平均平均校排名母數"] = ranks[key6a].Count;
                                            row["科目原始成績加權平均平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                        }
                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key6a = "科目原始成績加權平均平均類別一" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key6a))
                                            {
                                                int rr = ranks[key6a].IndexOf(selectScore[id6].avgScoreAC1) + 1;
                                                row["科目原始成績加權平均平均類別一"] = Utility.NoRound(selectScore[id6].avgScoreAC1);
                                                row["科目原始成績加權平均平均類別一排名"] = rr;
                                                row["科目原始成績加權平均平均類別一排名母數"] = ranks[key6a].Count;
                                                row["科目原始成績加權平均平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                            }
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key6a = "科目原始成績加權平均平均類別二" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key6a))
                                            {
                                                int rr = ranks[key6a].IndexOf(selectScore[id6].avgScoreAC2) + 1;
                                                row["科目原始成績加權平均平均類別二"] = Utility.NoRound(selectScore[id6].avgScoreAC2);
                                                row["科目原始成績加權平均平均類別二排名"] = rr;
                                                row["科目原始成績加權平均平均類別二排名母數"] = ranks[key6a].Count;
                                                row["科目原始成績加權平均平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                            }
                                        }
                                    }
                                }
                                #endregion 處理科目原始成績

                                #region 處理篩選科目原始成績
                                if (selectScore.ContainsKey(id8))
                                {
                                    row["篩選科目原始成績加權平均平均"] = Utility.NoRound(selectScore[id8].avgScoreA);
                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                    {
                                        string key8a = "篩選科目原始成績加權平均平均班排名" + studRec.RefClass.ClassID;
                                        if (ranks.ContainsKey(key8a))
                                        {
                                            int rr = ranks[key8a].IndexOf(selectScore[id8].avgScoreA) + 1;
                                            row["篩選科目原始成績加權平均平均班排名"] = rr;
                                            row["篩選科目原始成績加權平均平均班排名母數"] = ranks[key8a].Count;
                                            row["篩選科目原始成績加權平均平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                        }

                                        key8a = "篩選科目原始成績加權平均平均科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key8a))
                                        {
                                            int rr = ranks[key8a].IndexOf(selectScore[id8].avgScoreA) + 1;
                                            row["篩選科目原始成績加權平均平均科排名"] = rr;
                                            row["篩選科目原始成績加權平均平均科排名母數"] = ranks[key8a].Count;
                                            row["篩選科目原始成績加權平均平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                        }

                                        key8a = "篩選科目原始成績加權平均平均校排名" + gradeyear;
                                        if (ranks.ContainsKey(key8a))
                                        {
                                            int rr = ranks[key8a].IndexOf(selectScore[id8].avgScoreA) + 1;
                                            row["篩選科目原始成績加權平均平均校排名"] = rr;
                                            row["篩選科目原始成績加權平均平均校排名母數"] = ranks[key8a].Count;
                                            row["篩選科目原始成績加權平均平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                        }
                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key8a = "篩選科目原始成績加權平均平均類別一" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key8a))
                                            {
                                                int rr = ranks[key8a].IndexOf(selectScore[id8].avgScoreAC1) + 1;
                                                row["篩選科目原始成績加權平均平均類別一"] = Utility.NoRound(selectScore[id8].avgScoreAC1);
                                                row["篩選科目原始成績加權平均平均類別一排名"] = rr;
                                                row["篩選科目原始成績加權平均平均類別一排名母數"] = ranks[key8a].Count;
                                                row["篩選科目原始成績加權平均平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                            }
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key8a = "篩選科目原始成績加權平均平均類別二" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key8a))
                                            {
                                                int rr = ranks[key8a].IndexOf(selectScore[id8].avgScoreAC2) + 1;
                                                row["篩選科目原始成績加權平均平均類別二"] = Utility.NoRound(selectScore[id8].avgScoreAC2);
                                                row["篩選科目原始成績加權平均平均類別二排名"] = rr;
                                                row["篩選科目原始成績加權平均平均類別二排名母數"] = ranks[key8a].Count;
                                                row["篩選科目原始成績加權平均平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                            }
                                        }
                                    }
                                }
                                #endregion 處理篩選科目原始成績
                                #endregion 處理科目原始成績加權平均平均

                                // 不排名不放入
                                if (!noRankList.Contains(studRec.StudentID))
                                {
                                    _table.Rows.Add(row);

                                    if (OneClassCompleted != null)
                                        OneClassCompleted();
                                    //List<string> fields = new List<string>(docTemplate.MailMerge.GetFieldNames());
                                    //List<string> rmColumns = new List<string>();

                                    Aspose.Words.Document document = new Aspose.Words.Document();
                                    document = docTemplate;
                                    doc.Sections.Add(doc.ImportNode(document.Sections[0], true));

                                    doc.MailMerge.Execute(_table);
                                    doc.MailMerge.RemoveEmptyParagraphs = true;
                                    doc.MailMerge.DeleteFields();

                                    _table.Rows.Clear();

                                    #region PDF 存檔
                                    string reportNameW = FileKey;
                                    string pathW = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports", FolderName);
                                    if (!Directory.Exists(pathW))
                                        Directory.CreateDirectory(pathW);
                                    pathW = Path.Combine(pathW, reportNameW + ".pdf");

                                    if (File.Exists(pathW))
                                    {
                                        int i = 1;
                                        while (true)
                                        {
                                            string newPathW = Path.GetDirectoryName(pathW) + "\\" + Path.GetFileNameWithoutExtension(pathW) + (i++) + Path.GetExtension(pathW);
                                            if (!File.Exists(newPathW))
                                            {
                                                pathW = newPathW;
                                                break;
                                            }
                                        }
                                    }


                                    try
                                    {
                                        doc.Save(pathW, Aspose.Words.SaveFormat.Pdf);

                                        // 計算進度
                                        int xx = (int)(100d / StudSumCount * StudCountStart);
                                        FISCA.RTContext.Invoke(new Action<string, int>(PDF_Msg), new object[] { "產生學生個人PDF檔中...", xx });
                                        StudCountStart++;
                                    }
                                    catch (OutOfMemoryException exow)
                                    {
                                        exc = exow;
                                    }
                                    doc = null;
                                    GC.Collect();
                                }
                                #endregion

                            }

                            #endregion
                            try
                            {
                                if (sbPDFErr.Length > 0)
                                {
                                    string pathF = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports", FolderName);
                                    StreamWriter sw = new StreamWriter(pathF + "\\產生PDF發生錯誤.txt", true);
                                    sw.Write(sbPDFErr.ToString());
                                    sw.Close();
                                }
                            }
                            catch (Exception ex)
                            {

                            }
                        }
                        else
                        {
                            #region 產生 Word 檔案
                            int ClassCountStart = 0, ClassSumCount = classNameList.Count;
                            foreach (string className in classNameList)
                            {

                                // 產生 DataRow 資料
                                List<int> g1List = new List<int>();
                                List<int> g2List = new List<int>();
                                List<int> g3List = new List<int>();
                                List<int> g4List = new List<int>();

                                // Word 套表
                                // 樣板
                                Aspose.Words.Document doc = new Aspose.Words.Document();
                                Aspose.Words.Document docTemplate = setting.Template;
                                if (docTemplate == null)
                                    docTemplate = new Aspose.Words.Document(new MemoryStream(Properties.Resources.高中多學期學生成績證明書));


                                doc.Sections.Clear();


                                foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                                {
                                    //2018/3/31 穎驊註解，先在每一輪loop 開頭 就先濾掉不排名學生，可以避免資源浪費，以及
                                    // 可能會造成#9853 DataRow row = _table.NewRow(); 記憶體洩漏的問題，
                                    // 因為在一個DataRow row 產生後，其來源table 會預期它會被加入，
                                    //但若後來造印製表才判斷此份資料不需被加入就會產生記憶體洩漏，資料比數一大，就會造成系統當機。
                                    if (noRankList.Contains(studRec.StudentID))
                                        continue;


                                    if (studRec.RefClass.ClassName == className)
                                    {
                                        DataRow row = _table.NewRow();
                                        row["學校名稱"] = SchoolName;
                                        row["班級"] = studRec.RefClass.ClassName;
                                        row["座號"] = studRec.SeatNo;
                                        row["學號"] = studRec.StudentNumber;
                                        row["學生系統編號"] = studRec.StudentID;
                                        row["教師系統編號"] = "";
                                        row["教師姓名"] = "";
                                        if (studRec.RefClass.RefTeacher != null)
                                        {
                                            row["教師系統編號"] = studRec.RefClass.RefTeacher.TeacherID;
                                            row["教師姓名"] = studRec.RefClass.RefTeacher.TeacherName;
                                        }
                                        row["姓名"] = studRec.StudentName;
                                        row["科別"] = studRec.Department;
                                        row["類別一分類"] = (cat1Dict.ContainsKey(studRec.StudentID)) ? cat1Dict[studRec.StudentID] : "";
                                        row["類別二分類"] = (cat2Dict.ContainsKey(studRec.StudentID)) ? cat2Dict[studRec.StudentID] : "";
                                        row["一年級學年度"] = "";
                                        row["二年級學年度"] = "";
                                        row["三年級學年度"] = "";
                                        row["四年級學年度"] = "";
                                        g1List.Clear();
                                        g2List.Clear();
                                        g3List.Clear();
                                        g4List.Clear();

                                        int subjIndex = 1;
                                        // 處理科目
                                        #region 處理科目
                                        foreach (string subjName in setting.useSubjectPrintList)
                                        {
                                            string id = studRec.StudentID + "^^^" + subjName;

                                            if (selectScore.ContainsKey(id))
                                            {
                                                row["科目名稱" + subjIndex] = subjName;

                                                if (selectScore[id].gsSchoolYear11.HasValue)
                                                {
                                                    if (!g1List.Contains(selectScore[id].gsSchoolYear11.Value))
                                                        g1List.Add(selectScore[id].gsSchoolYear11.Value);
                                                }
                                                if (selectScore[id].gsSchoolYear12.HasValue)
                                                {
                                                    if (!g1List.Contains(selectScore[id].gsSchoolYear12.Value))
                                                        g1List.Add(selectScore[id].gsSchoolYear12.Value);
                                                }

                                                if (selectScore[id].gsSchoolYear21.HasValue)
                                                {
                                                    if (!g2List.Contains(selectScore[id].gsSchoolYear21.Value))
                                                        g2List.Add(selectScore[id].gsSchoolYear21.Value);
                                                }

                                                if (selectScore[id].gsSchoolYear22.HasValue)
                                                {
                                                    if (!g2List.Contains(selectScore[id].gsSchoolYear22.Value))
                                                        g2List.Add(selectScore[id].gsSchoolYear22.Value);
                                                }

                                                if (selectScore[id].gsSchoolYear31.HasValue)
                                                {
                                                    if (!g3List.Contains(selectScore[id].gsSchoolYear31.Value))
                                                        g3List.Add(selectScore[id].gsSchoolYear31.Value);
                                                }
                                                if (selectScore[id].gsSchoolYear32.HasValue)
                                                {
                                                    if (!g3List.Contains(selectScore[id].gsSchoolYear32.Value))
                                                        g3List.Add(selectScore[id].gsSchoolYear32.Value);
                                                }

                                                if (selectScore[id].gsSchoolYear41.HasValue)
                                                {
                                                    if (!g4List.Contains(selectScore[id].gsSchoolYear41.Value))
                                                        g4List.Add(selectScore[id].gsSchoolYear41.Value);
                                                }

                                                if (selectScore[id].gsSchoolYear42.HasValue)
                                                {
                                                    if (!g4List.Contains(selectScore[id].gsSchoolYear42.Value))
                                                        g4List.Add(selectScore[id].gsSchoolYear42.Value);
                                                }

                                                if (selectScore[id].gsScore11.HasValue)
                                                    row["一上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore11.Value);
                                                if (selectScore[id].gsCredit11.HasValue)
                                                    row["一上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit11.Value);
                                                if (selectScore[id].gsScore12.HasValue)
                                                    row["一下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore12.Value);

                                                if (selectScore[id].gsCredit12.HasValue)
                                                    row["一下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit12.Value);
                                                if (selectScore[id].gsScore21.HasValue)
                                                    row["二上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore21.Value);

                                                if (selectScore[id].gsCredit21.HasValue)
                                                    row["二上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit21.Value);

                                                if (selectScore[id].gsScore22.HasValue)
                                                    row["二下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore22.Value);

                                                if (selectScore[id].gsCredit22.HasValue)
                                                    row["二下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit22.Value);

                                                if (selectScore[id].gsScore31.HasValue)
                                                    row["三上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore31.Value);

                                                if (selectScore[id].gsCredit31.HasValue)
                                                    row["三上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit31.Value);

                                                if (selectScore[id].gsScore32.HasValue)
                                                    row["三下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore32.Value);

                                                if (selectScore[id].gsCredit32.HasValue)
                                                    row["三下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit32.Value);

                                                if (selectScore[id].gsScore41.HasValue)
                                                    row["四上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore41.Value);

                                                if (selectScore[id].gsCredit41.HasValue)
                                                    row["四上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit41.Value);

                                                if (selectScore[id].gsScore42.HasValue)
                                                    row["四下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore42.Value);

                                                if (selectScore[id].gsCredit42.HasValue)
                                                    row["四下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit42.Value);

                                                row["科目平均" + subjIndex] = Utility.NoRound(selectScore[id].avgScore);
                                                row["科目總分" + subjIndex] = Utility.NoRound(selectScore[id].sumScore);
                                                row["科目加權平均" + subjIndex] = Utility.NoRound(selectScore[id].avgScoreA);
                                                row["科目加權總分" + subjIndex] = Utility.NoRound(selectScore[id].sumScoreA);

                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key4 = "加權平均班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key4))
                                                    {
                                                        int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                        row["科目加權平均班排名" + subjIndex] = rr;
                                                        row["科目加權平均班排名母數" + subjIndex] = ranks[key4].Count;
                                                        row["科目加權平均班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                    }

                                                    key4 = "加權平均科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key4))
                                                    {
                                                        int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                        row["科目加權平均科排名" + subjIndex] = rr;
                                                        row["科目加權平均科排名母數" + subjIndex] = ranks[key4].Count;
                                                        row["科目加權平均科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);

                                                    }
                                                    key4 = "加權平均全校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key4))
                                                    {
                                                        int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                        row["科目加權平均校排名" + subjIndex] = rr;
                                                        row["科目加權平均校排名母數" + subjIndex] = ranks[key4].Count;
                                                        row["科目加權平均校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key4 = "加權平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key4))
                                                        {
                                                            if (ranks[key4].IndexOf(selectScore[id].avgScoreAC1) >= 0)
                                                            {
                                                                int rr = ranks[key4].IndexOf(selectScore[id].avgScoreAC1) + 1;
                                                                row["科目加權平均類別一排名" + subjIndex] = rr;
                                                                row["科目加權平均類別一排名母數" + subjIndex] = ranks[key4].Count;
                                                                row["科目加權平均類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                            }
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        if (ranks.ContainsKey(key4))
                                                        {
                                                            if (ranks[key4].IndexOf(selectScore[id].avgScoreAC2) >= 0)
                                                            {
                                                                int rr = ranks[key4].IndexOf(selectScore[id].avgScoreAC2) + 1;
                                                                row["科目加權平均類別二排名" + subjIndex] = rr;
                                                                row["科目加權平均類別二排名母數" + subjIndex] = ranks[key4].Count;
                                                                row["科目加權平均類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                            }
                                                        }
                                                    }

                                                    string key2 = "加權總分班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                        row["科目加權總分班排名" + subjIndex] = rr;
                                                        row["科目加權總分班排名母數" + subjIndex] = ranks[key2].Count;
                                                        row["科目加權總分班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }
                                                    key2 = "加權總分科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                        row["科目加權總分科排名" + subjIndex] = rr;
                                                        row["科目加權總分科排名母數" + subjIndex] = ranks[key2].Count;
                                                        row["科目加權總分科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                    }

                                                    key2 = "加權總分全校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                        row["科目加權總分校排名" + subjIndex] = rr;
                                                        row["科目加權總分校排名母數" + subjIndex] = ranks[key2].Count;
                                                        row["科目加權總分校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key2 = "加權總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            if (ranks[key2].IndexOf(selectScore[id].sumScoreAC1) >= 0)
                                                            {
                                                                int rr = ranks[key2].IndexOf(selectScore[id].sumScoreAC1) + 1;
                                                                row["科目加權總分類別一排名" + subjIndex] = rr;
                                                                row["科目加權總分類別一排名母數" + subjIndex] = ranks[key2].Count;
                                                                row["科目加權總分類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                            }
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key2 = "加權總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            if (ranks[key2].IndexOf(selectScore[id].sumScoreAC2) >= 0)
                                                            {
                                                                int rr = ranks[key2].IndexOf(selectScore[id].sumScoreAC2) + 1;
                                                                row["科目加權總分類別二排名" + subjIndex] = rr;
                                                                row["科目加權總分類別二排名母數" + subjIndex] = ranks[key2].Count;
                                                                row["科目加權總分類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                            }
                                                        }
                                                    }

                                                    string key3 = "平均班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key3))
                                                    {
                                                        int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                        row["科目平均班排名" + subjIndex] = rr;
                                                        row["科目平均班排名母數" + subjIndex] = ranks[key3].Count;
                                                        row["科目平均班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                    }

                                                    key3 = "平均科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key3))
                                                    {
                                                        int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                        row["科目平均科排名" + subjIndex] = rr;
                                                        row["科目平均科排名母數" + subjIndex] = ranks[key3].Count;
                                                        row["科目平均科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                    }
                                                    key3 = "平均全校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key3))
                                                    {
                                                        int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                        row["科目平均校排名" + subjIndex] = rr;
                                                        row["科目平均校排名母數" + subjIndex] = ranks[key3].Count;
                                                        row["科目平均校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key3 = "平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key3))
                                                        {
                                                            if (ranks[key3].IndexOf(selectScore[id].avgScoreC1) >= 0)
                                                            {
                                                                int rr = ranks[key3].IndexOf(selectScore[id].avgScoreC1) + 1;
                                                                row["科目平均類別一排名" + subjIndex] = rr;
                                                                row["科目平均類別一排名母數" + subjIndex] = ranks[key3].Count;
                                                                row["科目平均類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                            }
                                                        }
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key3 = "平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key3))
                                                        {
                                                            if (ranks[key3].IndexOf(selectScore[id].avgScoreC2) >= 0)
                                                            {
                                                                int rr = ranks[key3].IndexOf(selectScore[id].avgScoreC2) + 1;
                                                                row["科目平均類別二排名" + subjIndex] = rr;
                                                                row["科目平均類別二排名母數" + subjIndex] = ranks[key3].Count;
                                                                row["科目平均類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                            }
                                                        }
                                                    }
                                                    string key1 = "總分班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                        row["科目總分班排名" + subjIndex] = rr;
                                                        row["科目總分班排名母數" + subjIndex] = ranks[key1].Count;
                                                        row["科目總分班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                    }
                                                    key1 = "總分科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                        row["科目總分科排名" + subjIndex] = rr;
                                                        row["科目總分科排名母數" + subjIndex] = ranks[key1].Count;
                                                        row["科目總分科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                    }
                                                    key1 = "總分全校排名" + gradeyear + "^^^" + subjName;

                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                        row["科目總分校排名" + subjIndex] = rr;
                                                        row["科目總分校排名母數" + subjIndex] = ranks[key1].Count;
                                                        row["科目總分校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key1 = "總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            if (ranks[key1].IndexOf(selectScore[id].sumScoreC1) >= 0)
                                                            {
                                                                int rr = ranks[key1].IndexOf(selectScore[id].sumScoreC1) + 1;
                                                                row["科目總分類別一排名" + subjIndex] = rr;
                                                                row["科目總分類別一排名母數" + subjIndex] = ranks[key1].Count;
                                                                row["科目總分類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                            }
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key1 = "總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            if (ranks[key1].IndexOf(selectScore[id].sumScoreC2) >= 0)
                                                            {
                                                                int rr = ranks[key1].IndexOf(selectScore[id].sumScoreC2) + 1;
                                                                row["科目總分類別二排名" + subjIndex] = rr;
                                                                row["科目總分類別二排名母數" + subjIndex] = ranks[key1].Count;
                                                                row["科目總分類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                            }
                                                        }
                                                    }
                                                }

                                                subjIndex++;
                                            }
                                        }
                                        #endregion 處理科目


                                        subjIndex = 1;
                                        // 處理科目回歸
                                        #region 處理科目回歸
                                        foreach (string subjName in SubjMappingDict.Keys)
                                        {
                                            string id = studRec.StudentID + "^^^回歸" + subjName;

                                            if (selectScore.ContainsKey(id))
                                            {
                                                row["回歸科目名稱" + subjIndex] = subjName;

                                                if (selectScore[id].gsSchoolYear11.HasValue)
                                                {
                                                    if (!g1List.Contains(selectScore[id].gsSchoolYear11.Value))
                                                        g1List.Add(selectScore[id].gsSchoolYear11.Value);
                                                }
                                                if (selectScore[id].gsSchoolYear12.HasValue)
                                                {
                                                    if (!g1List.Contains(selectScore[id].gsSchoolYear12.Value))
                                                        g1List.Add(selectScore[id].gsSchoolYear12.Value);
                                                }

                                                if (selectScore[id].gsSchoolYear21.HasValue)
                                                {
                                                    if (!g2List.Contains(selectScore[id].gsSchoolYear21.Value))
                                                        g2List.Add(selectScore[id].gsSchoolYear21.Value);
                                                }

                                                if (selectScore[id].gsSchoolYear22.HasValue)
                                                {
                                                    if (!g2List.Contains(selectScore[id].gsSchoolYear22.Value))
                                                        g2List.Add(selectScore[id].gsSchoolYear22.Value);
                                                }

                                                if (selectScore[id].gsSchoolYear31.HasValue)
                                                {
                                                    if (!g3List.Contains(selectScore[id].gsSchoolYear31.Value))
                                                        g3List.Add(selectScore[id].gsSchoolYear31.Value);
                                                }
                                                if (selectScore[id].gsSchoolYear32.HasValue)
                                                {
                                                    if (!g3List.Contains(selectScore[id].gsSchoolYear32.Value))
                                                        g3List.Add(selectScore[id].gsSchoolYear32.Value);
                                                }

                                                if (selectScore[id].gsSchoolYear41.HasValue)
                                                {
                                                    if (!g4List.Contains(selectScore[id].gsSchoolYear41.Value))
                                                        g4List.Add(selectScore[id].gsSchoolYear41.Value);
                                                }

                                                if (selectScore[id].gsSchoolYear42.HasValue)
                                                {
                                                    if (!g4List.Contains(selectScore[id].gsSchoolYear42.Value))
                                                        g4List.Add(selectScore[id].gsSchoolYear42.Value);
                                                }

                                                if (selectScore[id].gsScore11.HasValue)
                                                    row["回歸一上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore11.Value);
                                                if (selectScore[id].gsCredit11.HasValue)
                                                    row["回歸一上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit11.Value);
                                                if (selectScore[id].gsScore12.HasValue)
                                                    row["回歸一下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore12.Value);

                                                if (selectScore[id].gsCredit12.HasValue)
                                                    row["回歸一下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit12.Value);
                                                if (selectScore[id].gsScore21.HasValue)
                                                    row["回歸二上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore21.Value);

                                                if (selectScore[id].gsCredit21.HasValue)
                                                    row["回歸二上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit21.Value);

                                                if (selectScore[id].gsScore22.HasValue)
                                                    row["回歸二下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore22.Value);

                                                if (selectScore[id].gsCredit22.HasValue)
                                                    row["回歸二下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit22.Value);

                                                if (selectScore[id].gsScore31.HasValue)
                                                    row["回歸三上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore31.Value);

                                                if (selectScore[id].gsCredit31.HasValue)
                                                    row["回歸三上科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit31.Value);

                                                if (selectScore[id].gsScore32.HasValue)
                                                    row["回歸三下科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore32.Value);

                                                if (selectScore[id].gsCredit32.HasValue)
                                                    row["回歸三下科目學分數" + subjIndex] = Utility.NoRound(selectScore[id].gsCredit32.Value);

                                                if (selectScore[id].gsScore41.HasValue)
                                                    row["回歸四上科目成績" + subjIndex] = Utility.NoRound(selectScore[id].gsScore41.Value);

                                                row["回歸科目平均" + subjIndex] = Utility.NoRound(selectScore[id].avgScore);
                                                row["回歸科目總分" + subjIndex] = Utility.NoRound(selectScore[id].sumScore);
                                                row["回歸科目加權平均" + subjIndex] = Utility.NoRound(selectScore[id].avgScoreA);
                                                row["回歸科目加權總分" + subjIndex] = Utility.NoRound(selectScore[id].sumScoreA);

                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {

                                                    #region 寫入各科目成績班排名
                                                    string key11 = "回歸一上科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key11))
                                                    {
                                                        if (selectScore[id].gsScore11.HasValue)
                                                        {
                                                            int rr = ranks[key11].IndexOf(selectScore[id].gsScore11.Value) + 1;
                                                            row["回歸一上科目班排名" + subjIndex] = rr;
                                                            row["回歸一上科目班排名母數" + subjIndex] = ranks[key11].Count;
                                                            row["回歸一上科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key11].Count);
                                                        }
                                                    }

                                                    string key12 = "回歸一下科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key12))
                                                    {
                                                        if (selectScore[id].gsScore12.HasValue)
                                                        {
                                                            int rr = ranks[key12].IndexOf(selectScore[id].gsScore12.Value) + 1;
                                                            row["回歸一下科目班排名" + subjIndex] = rr;
                                                            row["回歸一下科目班排名母數" + subjIndex] = ranks[key12].Count;
                                                            row["回歸一下科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key12].Count);
                                                        }
                                                    }

                                                    string key21 = "回歸二上科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key21))
                                                    {
                                                        if (selectScore[id].gsScore21.HasValue)
                                                        {
                                                            int rr = ranks[key21].IndexOf(selectScore[id].gsScore21.Value) + 1;
                                                            row["回歸二上科目班排名" + subjIndex] = rr;
                                                            row["回歸二上科目班排名母數" + subjIndex] = ranks[key21].Count;
                                                            row["回歸二上科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key21].Count);
                                                        }
                                                    }

                                                    string key22 = "回歸二下科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key22))
                                                    {
                                                        if (selectScore[id].gsScore22.HasValue)
                                                        {
                                                            int rr = ranks[key22].IndexOf(selectScore[id].gsScore22.Value) + 1;
                                                            row["回歸二下科目班排名" + subjIndex] = rr;
                                                            row["回歸二下科目班排名母數" + subjIndex] = ranks[key22].Count;
                                                            row["回歸二下科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key22].Count);
                                                        }
                                                    }

                                                    string key31 = "回歸三上科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key31))
                                                    {
                                                        if (selectScore[id].gsScore31.HasValue)
                                                        {
                                                            int rr = ranks[key31].IndexOf(selectScore[id].gsScore31.Value) + 1;
                                                            row["回歸三上科目班排名" + subjIndex] = rr;
                                                            row["回歸三上科目班排名母數" + subjIndex] = ranks[key31].Count;
                                                            row["回歸三上科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key31].Count);
                                                        }
                                                    }

                                                    string key32 = "回歸三下科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key32))
                                                    {
                                                        if (selectScore[id].gsScore32.HasValue)
                                                        {
                                                            int rr = ranks[key32].IndexOf(selectScore[id].gsScore32.Value) + 1;
                                                            row["回歸三下科目班排名" + subjIndex] = rr;
                                                            row["回歸三下科目班排名母數" + subjIndex] = ranks[key32].Count;
                                                            row["回歸三下科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key32].Count);
                                                        }
                                                    }

                                                    string key41 = "回歸四上科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key41))
                                                    {
                                                        if (selectScore[id].gsScore41.HasValue)
                                                        {
                                                            int rr = ranks[key41].IndexOf(selectScore[id].gsScore41.Value) + 1;
                                                            row["回歸四上科目班排名" + subjIndex] = rr;
                                                            row["回歸四上科目班排名母數" + subjIndex] = ranks[key41].Count;
                                                            row["回歸四上科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key41].Count);
                                                        }
                                                    }

                                                    string key42 = "回歸四下科目班排名" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key42))
                                                    {
                                                        if (selectScore[id].gsScore42.HasValue)
                                                        {
                                                            int rr = ranks[key42].IndexOf(selectScore[id].gsScore42.Value) + 1;
                                                            row["回歸四下科目班排名" + subjIndex] = rr;
                                                            row["回歸四下科目班排名母數" + subjIndex] = ranks[key42].Count;
                                                            row["回歸四下科目班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key42].Count);
                                                        }
                                                    }

                                                    #endregion

                                                    #region 寫入各科目成績科排名
                                                    key11 = "回歸一上科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key11))
                                                    {
                                                        if (selectScore[id].gsScore11.HasValue)
                                                        {
                                                            int rr = ranks[key11].IndexOf(selectScore[id].gsScore11.Value) + 1;
                                                            row["回歸一上科目科排名" + subjIndex] = rr;
                                                            row["回歸一上科目科排名母數" + subjIndex] = ranks[key11].Count;
                                                            row["回歸一上科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key11].Count);
                                                        }
                                                    }

                                                    key12 = "回歸一下科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key12))
                                                    {
                                                        if (selectScore[id].gsScore12.HasValue)
                                                        {
                                                            int rr = ranks[key12].IndexOf(selectScore[id].gsScore12.Value) + 1;
                                                            row["回歸一下科目科排名" + subjIndex] = rr;
                                                            row["回歸一下科目科排名母數" + subjIndex] = ranks[key12].Count;
                                                            row["回歸一下科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key12].Count);
                                                        }
                                                    }

                                                    key21 = "回歸二上科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key21))
                                                    {
                                                        if (selectScore[id].gsScore21.HasValue)
                                                        {
                                                            int rr = ranks[key21].IndexOf(selectScore[id].gsScore21.Value) + 1;
                                                            row["回歸二上科目科排名" + subjIndex] = rr;
                                                            row["回歸二上科目科排名母數" + subjIndex] = ranks[key21].Count;
                                                            row["回歸二上科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key21].Count);
                                                        }
                                                    }

                                                    key22 = "回歸二下科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key22))
                                                    {
                                                        if (selectScore[id].gsScore22.HasValue)
                                                        {
                                                            int rr = ranks[key22].IndexOf(selectScore[id].gsScore22.Value) + 1;
                                                            row["回歸二下科目科排名" + subjIndex] = rr;
                                                            row["回歸二下科目科排名母數" + subjIndex] = ranks[key22].Count;
                                                            row["回歸二下科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key22].Count);
                                                        }
                                                    }

                                                    key31 = "回歸三上科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key31))
                                                    {
                                                        if (selectScore[id].gsScore31.HasValue)
                                                        {
                                                            int rr = ranks[key31].IndexOf(selectScore[id].gsScore31.Value) + 1;
                                                            row["回歸三上科目科排名" + subjIndex] = rr;
                                                            row["回歸三上科目科排名母數" + subjIndex] = ranks[key31].Count;
                                                            row["回歸三上科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key31].Count);
                                                        }
                                                    }

                                                    key32 = "回歸三下科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key32))
                                                    {
                                                        if (selectScore[id].gsScore32.HasValue)
                                                        {
                                                            int rr = ranks[key32].IndexOf(selectScore[id].gsScore32.Value) + 1;
                                                            row["回歸三下科目科排名" + subjIndex] = rr;
                                                            row["回歸三下科目科排名母數" + subjIndex] = ranks[key32].Count;
                                                            row["回歸三下科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key32].Count);
                                                        }
                                                    }

                                                    key41 = "回歸四上科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key41))
                                                    {
                                                        if (selectScore[id].gsScore41.HasValue)
                                                        {
                                                            int rr = ranks[key41].IndexOf(selectScore[id].gsScore41.Value) + 1;
                                                            row["回歸四上科目科排名" + subjIndex] = rr;
                                                            row["回歸四上科目科排名母數" + subjIndex] = ranks[key41].Count;
                                                            row["回歸四上科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key41].Count);
                                                        }
                                                    }

                                                    key42 = "回歸四下科目科排名" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key42))
                                                    {
                                                        if (selectScore[id].gsScore42.HasValue)
                                                        {
                                                            int rr = ranks[key42].IndexOf(selectScore[id].gsScore42.Value) + 1;
                                                            row["回歸四下科目科排名" + subjIndex] = rr;
                                                            row["回歸四下科目科排名母數" + subjIndex] = ranks[key42].Count;
                                                            row["回歸四下科目科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key42].Count);
                                                        }
                                                    }

                                                    #endregion

                                                    #region 寫入各科目成績校排名
                                                    key11 = "回歸一上科目校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key11))
                                                    {
                                                        if (selectScore[id].gsScore11.HasValue)
                                                        {
                                                            int rr = ranks[key11].IndexOf(selectScore[id].gsScore11.Value) + 1;
                                                            row["回歸一上科目校排名" + subjIndex] = rr;
                                                            row["回歸一上科目校排名母數" + subjIndex] = ranks[key11].Count;
                                                            row["回歸一上科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key11].Count);
                                                        }
                                                    }

                                                    key12 = "回歸一下科目校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key12))
                                                    {
                                                        if (selectScore[id].gsScore12.HasValue)
                                                        {
                                                            int rr = ranks[key12].IndexOf(selectScore[id].gsScore12.Value) + 1;
                                                            row["回歸一下科目校排名" + subjIndex] = rr;
                                                            row["回歸一下科目校排名母數" + subjIndex] = ranks[key12].Count;
                                                            row["回歸一下科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key12].Count);
                                                        }
                                                    }

                                                    key21 = "回歸二上科目校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key21))
                                                    {
                                                        if (selectScore[id].gsScore21.HasValue)
                                                        {
                                                            int rr = ranks[key21].IndexOf(selectScore[id].gsScore21.Value) + 1;
                                                            row["回歸二上科目校排名" + subjIndex] = rr;
                                                            row["回歸二上科目校排名母數" + subjIndex] = ranks[key21].Count;
                                                            row["回歸二上科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key21].Count);
                                                        }
                                                    }

                                                    key22 = "回歸二下科目校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key22))
                                                    {
                                                        if (selectScore[id].gsScore22.HasValue)
                                                        {
                                                            int rr = ranks[key22].IndexOf(selectScore[id].gsScore22.Value) + 1;
                                                            row["回歸二下科目校排名" + subjIndex] = rr;
                                                            row["回歸二下科目校排名母數" + subjIndex] = ranks[key22].Count;
                                                            row["回歸二下科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key22].Count);
                                                        }
                                                    }

                                                    key31 = "回歸三上科目校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key31))
                                                    {
                                                        if (selectScore[id].gsScore31.HasValue)
                                                        {
                                                            int rr = ranks[key31].IndexOf(selectScore[id].gsScore31.Value) + 1;
                                                            row["回歸三上科目校排名" + subjIndex] = rr;
                                                            row["回歸三上科目校排名母數" + subjIndex] = ranks[key31].Count;
                                                            row["回歸三上科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key31].Count);
                                                        }
                                                    }

                                                    key32 = "回歸三下科目校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key32))
                                                    {
                                                        if (selectScore[id].gsScore32.HasValue)
                                                        {
                                                            int rr = ranks[key32].IndexOf(selectScore[id].gsScore32.Value) + 1;
                                                            row["回歸三下科目校排名" + subjIndex] = rr;
                                                            row["回歸三下科目校排名母數" + subjIndex] = ranks[key32].Count;
                                                            row["回歸三下科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key32].Count);
                                                        }
                                                    }

                                                    key41 = "回歸四上科目校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key41))
                                                    {
                                                        if (selectScore[id].gsScore41.HasValue)
                                                        {
                                                            int rr = ranks[key41].IndexOf(selectScore[id].gsScore41.Value) + 1;
                                                            row["回歸四上科目校排名" + subjIndex] = rr;
                                                            row["回歸四上科目校排名母數" + subjIndex] = ranks[key41].Count;
                                                            row["回歸四上科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key41].Count);
                                                        }
                                                    }

                                                    key42 = "回歸四下科目校排名" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key42))
                                                    {
                                                        if (selectScore[id].gsScore42.HasValue)
                                                        {
                                                            int rr = ranks[key42].IndexOf(selectScore[id].gsScore42.Value) + 1;
                                                            row["回歸四下科目校排名" + subjIndex] = rr;
                                                            row["回歸四下科目校排名母數" + subjIndex] = ranks[key42].Count;
                                                            row["回歸四下科目校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key42].Count);
                                                        }
                                                    }

                                                    #endregion

                                                    #region 寫入各科目成績類1排名
                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key11 = "回歸一上科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key11))
                                                        {
                                                            if (selectScore[id].gsScore11.HasValue)
                                                            {
                                                                int rr = ranks[key11].IndexOf(selectScore[id].gsScore11.Value) + 1;
                                                                row["回歸一上科目類1排名" + subjIndex] = rr;
                                                                row["回歸一上科目類1排名母數" + subjIndex] = ranks[key11].Count;
                                                                row["回歸一上科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key11].Count);
                                                            }
                                                        }

                                                        key12 = "回歸一下科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key12))
                                                        {
                                                            if (selectScore[id].gsScore12.HasValue)
                                                            {
                                                                int rr = ranks[key12].IndexOf(selectScore[id].gsScore12.Value) + 1;
                                                                row["回歸一下科目類1排名" + subjIndex] = rr;
                                                                row["回歸一下科目類1排名母數" + subjIndex] = ranks[key12].Count;
                                                                row["回歸一下科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key12].Count);
                                                            }
                                                        }

                                                        key21 = "回歸二上科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key21))
                                                        {
                                                            if (selectScore[id].gsScore21.HasValue)
                                                            {
                                                                int rr = ranks[key21].IndexOf(selectScore[id].gsScore21.Value) + 1;
                                                                row["回歸二上科目類1排名" + subjIndex] = rr;
                                                                row["回歸二上科目類1排名母數" + subjIndex] = ranks[key21].Count;
                                                                row["回歸二上科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key21].Count);
                                                            }
                                                        }

                                                        key22 = "回歸二下科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key22))
                                                        {
                                                            if (selectScore[id].gsScore22.HasValue)
                                                            {
                                                                int rr = ranks[key22].IndexOf(selectScore[id].gsScore22.Value) + 1;
                                                                row["回歸二下科目類1排名" + subjIndex] = rr;
                                                                row["回歸二下科目類1排名母數" + subjIndex] = ranks[key22].Count;
                                                                row["回歸二下科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key22].Count);
                                                            }
                                                        }

                                                        key31 = "回歸三上科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key31))
                                                        {
                                                            if (selectScore[id].gsScore31.HasValue)
                                                            {
                                                                int rr = ranks[key31].IndexOf(selectScore[id].gsScore31.Value) + 1;
                                                                row["回歸三上科目類1排名" + subjIndex] = rr;
                                                                row["回歸三上科目類1排名母數" + subjIndex] = ranks[key31].Count;
                                                                row["回歸三上科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key31].Count);
                                                            }
                                                        }

                                                        key32 = "回歸三下科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key32))
                                                        {
                                                            if (selectScore[id].gsScore32.HasValue)
                                                            {
                                                                int rr = ranks[key32].IndexOf(selectScore[id].gsScore32.Value) + 1;
                                                                row["回歸三下科目類1排名" + subjIndex] = rr;
                                                                row["回歸三下科目類1排名母數" + subjIndex] = ranks[key32].Count;
                                                                row["回歸三下科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key32].Count);
                                                            }
                                                        }

                                                        key41 = "回歸四上科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key41))
                                                        {
                                                            if (selectScore[id].gsScore41.HasValue)
                                                            {
                                                                int rr = ranks[key41].IndexOf(selectScore[id].gsScore41.Value) + 1;
                                                                row["回歸四上科目類1排名" + subjIndex] = rr;
                                                                row["回歸四上科目類1排名母數" + subjIndex] = ranks[key41].Count;
                                                                row["回歸四上科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key41].Count);
                                                            }
                                                        }

                                                        key42 = "回歸四下科目類1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key42))
                                                        {
                                                            if (selectScore[id].gsScore42.HasValue)
                                                            {
                                                                int rr = ranks[key42].IndexOf(selectScore[id].gsScore42.Value) + 1;
                                                                row["回歸四下科目類1排名" + subjIndex] = rr;
                                                                row["回歸四下科目類1排名母數" + subjIndex] = ranks[key42].Count;
                                                                row["回歸四下科目類1排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key42].Count);
                                                            }
                                                        }
                                                    }

                                                    #endregion

                                                    #region 寫入各科目成績類2排名

                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key11 = "回歸一上科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key11))
                                                        {
                                                            if (selectScore[id].gsScore11.HasValue)
                                                            {
                                                                int rr = ranks[key11].IndexOf(selectScore[id].gsScore11.Value) + 1;
                                                                row["回歸一上科目類2排名" + subjIndex] = rr;
                                                                row["回歸一上科目類2排名母數" + subjIndex] = ranks[key11].Count;
                                                                row["回歸一上科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key11].Count);
                                                            }
                                                        }

                                                        key12 = "回歸一下科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key12))
                                                        {
                                                            if (selectScore[id].gsScore12.HasValue)
                                                            {
                                                                int rr = ranks[key12].IndexOf(selectScore[id].gsScore12.Value) + 1;
                                                                row["回歸一下科目類2排名" + subjIndex] = rr;
                                                                row["回歸一下科目類2排名母數" + subjIndex] = ranks[key12].Count;
                                                                row["回歸一下科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key12].Count);
                                                            }
                                                        }

                                                        key21 = "回歸二上科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key21))
                                                        {
                                                            if (selectScore[id].gsScore21.HasValue)
                                                            {
                                                                int rr = ranks[key21].IndexOf(selectScore[id].gsScore21.Value) + 1;
                                                                row["回歸二上科目類2排名" + subjIndex] = rr;
                                                                row["回歸二上科目類2排名母數" + subjIndex] = ranks[key21].Count;
                                                                row["回歸二上科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key21].Count);
                                                            }
                                                        }

                                                        key22 = "回歸二下科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key22))
                                                        {
                                                            if (selectScore[id].gsScore22.HasValue)
                                                            {
                                                                int rr = ranks[key22].IndexOf(selectScore[id].gsScore22.Value) + 1;
                                                                row["回歸二下科目類2排名" + subjIndex] = rr;
                                                                row["回歸二下科目類2排名母數" + subjIndex] = ranks[key22].Count;
                                                                row["回歸二下科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key22].Count);
                                                            }
                                                        }

                                                        key31 = "回歸三上科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key31))
                                                        {
                                                            if (selectScore[id].gsScore31.HasValue)
                                                            {
                                                                int rr = ranks[key31].IndexOf(selectScore[id].gsScore31.Value) + 1;
                                                                row["回歸三上科目類2排名" + subjIndex] = rr;
                                                                row["回歸三上科目類2排名母數" + subjIndex] = ranks[key31].Count;
                                                                row["回歸三上科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key31].Count);
                                                            }
                                                        }

                                                        key32 = "回歸三下科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key32))
                                                        {
                                                            if (selectScore[id].gsScore32.HasValue)
                                                            {
                                                                int rr = ranks[key32].IndexOf(selectScore[id].gsScore32.Value) + 1;
                                                                row["回歸三下科目類2排名" + subjIndex] = rr;
                                                                row["回歸三下科目類2排名母數" + subjIndex] = ranks[key32].Count;
                                                                row["回歸三下科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key32].Count);
                                                            }
                                                        }

                                                        key41 = "回歸四上科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key41))
                                                        {
                                                            if (selectScore[id].gsScore41.HasValue)
                                                            {
                                                                int rr = ranks[key41].IndexOf(selectScore[id].gsScore41.Value) + 1;
                                                                row["回歸四上科目類2排名" + subjIndex] = rr;
                                                                row["回歸四上科目類2排名母數" + subjIndex] = ranks[key41].Count;
                                                                row["回歸四上科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key41].Count);
                                                            }
                                                        }

                                                        key42 = "回歸四下科目類2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key42))
                                                        {
                                                            if (selectScore[id].gsScore42.HasValue)
                                                            {
                                                                int rr = ranks[key42].IndexOf(selectScore[id].gsScore42.Value) + 1;
                                                                row["回歸四下科目類2排名" + subjIndex] = rr;
                                                                row["回歸四下科目類2排名母數" + subjIndex] = ranks[key42].Count;
                                                                row["回歸四下科目類2排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key42].Count);
                                                            }
                                                        }
                                                    }

                                                    #endregion


                                                    string key4 = "加權平均班排名_回歸_" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key4))
                                                    {
                                                        int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                        row["回歸科目加權平均班排名" + subjIndex] = rr;
                                                        row["回歸科目加權平均班排名母數" + subjIndex] = ranks[key4].Count;
                                                        row["回歸科目加權平均班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                    }

                                                    key4 = "加權平均科排名_回歸_" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key4))
                                                    {
                                                        int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                        row["回歸科目加權平均科排名" + subjIndex] = rr;
                                                        row["回歸科目加權平均科排名母數" + subjIndex] = ranks[key4].Count;
                                                        row["回歸科目加權平均科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);

                                                    }
                                                    key4 = "加權平均全校排名_回歸_" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key4))
                                                    {
                                                        int rr = ranks[key4].IndexOf(selectScore[id].avgScoreA) + 1;
                                                        row["回歸科目加權平均校排名" + subjIndex] = rr;
                                                        row["回歸科目加權平均校排名母數" + subjIndex] = ranks[key4].Count;
                                                        row["回歸科目加權平均校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key4 = "加權平均類別1排名_回歸_" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key4))
                                                        {
                                                            if (ranks[key4].IndexOf(selectScore[id].avgScoreAC1) >= 0)
                                                            {
                                                                int rr = ranks[key4].IndexOf(selectScore[id].avgScoreAC1) + 1;
                                                                row["回歸科目加權平均類別一排名" + subjIndex] = rr;
                                                                row["回歸科目加權平均類別一排名母數" + subjIndex] = ranks[key4].Count;
                                                                row["回歸科目加權平均類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                            }
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        if (ranks.ContainsKey(key4))
                                                        {
                                                            if (ranks[key4].IndexOf(selectScore[id].avgScoreAC2) >= 0)
                                                            {
                                                                int rr = ranks[key4].IndexOf(selectScore[id].avgScoreAC2) + 1;
                                                                row["回歸科目加權平均類別二排名" + subjIndex] = rr;
                                                                row["回歸科目加權平均類別二排名母數" + subjIndex] = ranks[key4].Count;
                                                                row["回歸科目加權平均類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key4].Count);
                                                            }
                                                        }
                                                    }

                                                    string key2 = "加權總分班排名_回歸_" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                        row["回歸科目加權總分班排名" + subjIndex] = rr;
                                                        row["回歸科目加權總分班排名母數" + subjIndex] = ranks[key2].Count;
                                                        row["回歸科目加權總分班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }
                                                    key2 = "加權總分科排名_回歸_" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                        row["回歸科目加權總分科排名" + subjIndex] = rr;
                                                        row["回歸科目加權總分科排名母數" + subjIndex] = ranks[key2].Count;
                                                        row["回歸科目加權總分科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                    }

                                                    key2 = "加權總分全校排名_回歸_" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1;
                                                        row["回歸科目加權總分校排名" + subjIndex] = rr;
                                                        row["回歸科目加權總分校排名母數" + subjIndex] = ranks[key2].Count;
                                                        row["回歸科目加權總分校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key2 = "加權總分類別1排名_回歸_" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            if (ranks[key2].IndexOf(selectScore[id].sumScoreAC1) >= 0)
                                                            {
                                                                int rr = ranks[key2].IndexOf(selectScore[id].sumScoreAC1) + 1;
                                                                row["回歸科目加權總分類別一排名" + subjIndex] = rr;
                                                                row["回歸科目加權總分類別一排名母數" + subjIndex] = ranks[key2].Count;
                                                                row["回歸科目加權總分類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                            }
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key2 = "加權總分類別2排名_回歸_" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            if (ranks[key2].IndexOf(selectScore[id].sumScoreAC2) >= 0)
                                                            {
                                                                int rr = ranks[key2].IndexOf(selectScore[id].sumScoreAC2) + 1;
                                                                row["回歸科目加權總分類別二排名" + subjIndex] = rr;
                                                                row["回歸科目加權總分類別二排名母數" + subjIndex] = ranks[key2].Count;
                                                                row["回歸科目加權總分類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                            }
                                                        }
                                                    }

                                                    string key3 = "平均班排名_回歸_" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key3))
                                                    {
                                                        int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                        row["回歸科目平均班排名" + subjIndex] = rr;
                                                        row["回歸科目平均班排名母數" + subjIndex] = ranks[key3].Count;
                                                        row["回歸科目平均班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                    }

                                                    key3 = "平均科排名_回歸_" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key3))
                                                    {
                                                        int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                        row["回歸科目平均科排名" + subjIndex] = rr;
                                                        row["回歸科目平均科排名母數" + subjIndex] = ranks[key3].Count;
                                                        row["回歸科目平均科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                    }
                                                    key3 = "平均全校排名_回歸_" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key3))
                                                    {
                                                        int rr = ranks[key3].IndexOf(selectScore[id].avgScore) + 1;
                                                        row["回歸科目平均校排名" + subjIndex] = rr;
                                                        row["回歸科目平均校排名母數" + subjIndex] = ranks[key3].Count;
                                                        row["回歸科目平均校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key3 = "平均類別1排名_回歸_" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key3))
                                                        {
                                                            if (ranks[key3].IndexOf(selectScore[id].avgScoreC1) >= 0)
                                                            {
                                                                int rr = ranks[key3].IndexOf(selectScore[id].avgScoreC1) + 1;
                                                                row["回歸科目平均類別一排名" + subjIndex] = rr;
                                                                row["回歸科目平均類別一排名母數" + subjIndex] = ranks[key3].Count;
                                                                row["回歸科目平均類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                            }
                                                        }
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key3 = "平均類別2排名_回歸_" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key3))
                                                        {
                                                            if (ranks[key3].IndexOf(selectScore[id].avgScoreC2) >= 0)
                                                            {
                                                                int rr = ranks[key3].IndexOf(selectScore[id].avgScoreC2) + 1;
                                                                row["回歸科目平均類別二排名" + subjIndex] = rr;
                                                                row["回歸科目平均類別二排名母數" + subjIndex] = ranks[key3].Count;
                                                                row["回歸科目平均類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key3].Count);
                                                            }
                                                        }
                                                    }
                                                    string key1 = "總分班排名_回歸_" + studRec.RefClass.ClassID + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                        row["回歸科目總分班排名" + subjIndex] = rr;
                                                        row["回歸科目總分班排名母數" + subjIndex] = ranks[key1].Count;
                                                        row["回歸科目總分班排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                    }
                                                    key1 = "總分科排名_回歸_" + studRec.Department + "^^^" + gradeyear + "^^^" + subjName;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                        row["回歸科目總分科排名" + subjIndex] = rr;
                                                        row["回歸科目總分科排名母數" + subjIndex] = ranks[key1].Count;
                                                        row["回歸科目總分科排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                    }
                                                    key1 = "總分全校排名_回歸_" + gradeyear + "^^^" + subjName;

                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        int rr = ranks[key1].IndexOf(selectScore[id].sumScore) + 1;
                                                        row["回歸科目總分校排名" + subjIndex] = rr;
                                                        row["回歸科目總分校排名母數" + subjIndex] = ranks[key1].Count;
                                                        row["回歸科目總分校排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key1 = "總分類別1排名_回歸_" + studRec.Fields["tag1"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            if (ranks[key1].IndexOf(selectScore[id].sumScoreC1) >= 0)
                                                            {
                                                                int rr = ranks[key1].IndexOf(selectScore[id].sumScoreC1) + 1;
                                                                row["回歸科目總分類別一排名" + subjIndex] = rr;
                                                                row["回歸科目總分類別一排名母數" + subjIndex] = ranks[key1].Count;
                                                                row["回歸科目總分類別一排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                            }
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key1 = "總分類別2排名_回歸_" + studRec.Fields["tag2"] + "^^^" + gradeyear + "^^^" + subjName;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            if (ranks[key1].IndexOf(selectScore[id].sumScoreC2) >= 0)
                                                            {
                                                                int rr = ranks[key1].IndexOf(selectScore[id].sumScoreC2) + 1;
                                                                row["回歸科目總分類別二排名" + subjIndex] = rr;
                                                                row["回歸科目總分類別二排名母數" + subjIndex] = ranks[key1].Count;
                                                                row["回歸科目總分類別二排名百分比" + subjIndex] = Utility.ParseRankPercent(rr, ranks[key1].Count);
                                                            }
                                                        }
                                                    }
                                                }

                                                subjIndex++;
                                            }
                                        }
                                        #endregion 處理科目回歸


                                        // 處理科目年級學年度
                                        row["一年級學年度"] = string.Join("/", g1List.ToArray());
                                        row["二年級學年度"] = string.Join("/", g2List.ToArray());
                                        row["三年級學年度"] = string.Join("/", g3List.ToArray());
                                        row["四年級學年度"] = string.Join("/", g4List.ToArray());

                                        // 處理各學期各科目排名
                                        if (StudSemsSubjRatingDict.ContainsKey(studRec.StudentID))
                                        {
                                            foreach (StudSemsSubjRating ssr in StudSemsSubjRatingDict[studRec.StudentID])
                                            {

                                                #region 舊
                                                if (ssr.GradeYear == "1" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["一上科目排名" + i] = ssr.GetClassRank(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "1" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["一下科目排名" + i] = ssr.GetClassRank(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "2" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["二上科目排名" + i] = ssr.GetClassRank(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "2" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["二下科目排名" + i] = ssr.GetClassRank(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "3" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["三上科目排名" + i] = ssr.GetClassRank(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "3" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["三下科目排名" + i] = ssr.GetClassRank(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "4" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["四上科目排名" + i] = ssr.GetClassRank(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "4" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["四下科目排名" + i] = ssr.GetClassRank(subjName);
                                                    }
                                                }
                                                #endregion

                                                #region 班排
                                                if (ssr.GradeYear == "1" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["一上科目班排名" + i] = ssr.GetClassRank(subjName);
                                                        row["一上科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                        row["一上科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "1" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["一下科目班排名" + i] = ssr.GetClassRank(subjName);
                                                        row["一下科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                        row["一下科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "2" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["二上科目班排名" + i] = ssr.GetClassRank(subjName);
                                                        row["二上科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                        row["二上科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "2" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["二下科目班排名" + i] = ssr.GetClassRank(subjName);
                                                        row["二下科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                        row["二下科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "3" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["三上科目班排名" + i] = ssr.GetClassRank(subjName);
                                                        row["三上科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                        row["三上科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "3" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["三下科目班排名" + i] = ssr.GetClassRank(subjName);
                                                        row["三下科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                        row["三下科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "4" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["四上科目班排名" + i] = ssr.GetClassRank(subjName);
                                                        row["四上科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                        row["四上科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "4" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["四下科目班排名" + i] = ssr.GetClassRank(subjName);
                                                        row["四下科目班排名母數" + i] = ssr.GetClassCount(subjName);
                                                        row["四下科目班排名百分比" + i] = ssr.GetClassRankP(subjName);
                                                    }
                                                }
                                                #endregion

                                                #region 科排
                                                if (ssr.GradeYear == "1" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["一上科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                        row["一上科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                        row["一上科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "1" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["一下科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                        row["一下科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                        row["一下科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "2" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["二上科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                        row["二上科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                        row["二上科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "2" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["二下科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                        row["二下科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                        row["二下科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "3" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["三上科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                        row["三上科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                        row["三上科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "3" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["三下科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                        row["三下科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                        row["三下科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "4" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["四上科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                        row["四上科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                        row["四上科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "4" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["四下科目科排名" + i] = ssr.GetDeptRank(subjName);
                                                        row["四下科目科排名母數" + i] = ssr.GetDeptCount(subjName);
                                                        row["四下科目科排名百分比" + i] = ssr.GetDeptRankP(subjName);
                                                    }
                                                }
                                                #endregion

                                                #region 校排
                                                if (ssr.GradeYear == "1" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["一上科目校排名" + i] = ssr.GetYearRank(subjName);
                                                        row["一上科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                        row["一上科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "1" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["一下科目校排名" + i] = ssr.GetYearRank(subjName);
                                                        row["一下科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                        row["一下科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "2" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["二上科目校排名" + i] = ssr.GetYearRank(subjName);
                                                        row["二上科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                        row["二上科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "2" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["二下科目校排名" + i] = ssr.GetYearRank(subjName);
                                                        row["二下科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                        row["二下科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "3" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["三上科目校排名" + i] = ssr.GetYearRank(subjName);
                                                        row["三上科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                        row["三上科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "3" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["三下科目校排名" + i] = ssr.GetYearRank(subjName);
                                                        row["三下科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                        row["三下科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "4" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["四上科目校排名" + i] = ssr.GetYearRank(subjName);
                                                        row["四上科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                        row["四上科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "4" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["四下科目校排名" + i] = ssr.GetYearRank(subjName);
                                                        row["四下科目校排名母數" + i] = ssr.GetYearCount(subjName);
                                                        row["四下科目校排名百分比" + i] = ssr.GetYearRankP(subjName);
                                                    }
                                                }
                                                #endregion

                                                #region 類1
                                                if (ssr.GradeYear == "1" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["一上科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                        row["一上科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                        row["一上科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "1" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["一下科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                        row["一下科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                        row["一下科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "2" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["二上科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                        row["二上科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                        row["二上科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "2" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["二下科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                        row["二下科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                        row["二下科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "3" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["三上科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                        row["三上科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                        row["三上科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "3" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["三下科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                        row["三下科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                        row["三下科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                                    }
                                                }

                                                if (ssr.GradeYear == "4" && ssr.Semester == "1")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["四上科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                        row["四上科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                        row["四上科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                                    }
                                                }
                                                if (ssr.GradeYear == "4" && ssr.Semester == "2")
                                                {
                                                    for (int i = 1; i <= subjCot; i++)
                                                    {
                                                        string subjName = row["科目名稱" + i].ToString();
                                                        row["四下科目類1排名" + i] = ssr.GetGroup1Rank(subjName);
                                                        row["四下科目類1排名母數" + i] = ssr.GetGroup1Count(subjName);
                                                        row["四下科目類1排名百分比" + i] = ssr.GetGroup1RankP(subjName);
                                                    }
                                                }
                                                #endregion

                                            }
                                        }

                                        if (setting.計算學業成績排名)
                                        {
                                            // 處理學業
                                            #region 處理學業
                                            string id1 = studRec.StudentID + "學業";

                                            #region 讀取各學期學業排名
                                            if (StudSemsEntryRatingDict.ContainsKey(studRec.StudentID))
                                            {
                                                foreach (StudSemsEntryRating sser in StudSemsEntryRatingDict[studRec.StudentID])
                                                {
                                                    if (sser.GradeYear == "1" && sser.Semester == "1")
                                                    {
                                                        if (sser.ClassRank.HasValue)
                                                        {
                                                            row["一上學業班排名"] = sser.ClassRank.Value;
                                                            if (sser.ClassCount.HasValue)
                                                            {
                                                                row["一上學業班排名母數"] = sser.ClassCount.Value;
                                                                row["一上學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                            }
                                                        }

                                                        if (sser.DeptRank.HasValue)
                                                        {
                                                            row["一上學業科排名"] = sser.DeptRank.Value;
                                                            if (sser.DeptCount.HasValue)
                                                            {
                                                                row["一上學業科排名母數"] = sser.DeptCount.Value;
                                                                row["一上學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                            }
                                                        }

                                                        if (sser.YearRank.HasValue)
                                                        {
                                                            row["一上學業校排名"] = sser.YearRank.Value;
                                                            if (sser.YearCount.HasValue)
                                                            {
                                                                row["一上學業校排名母數"] = sser.YearCount.Value;
                                                                row["一上學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                            }
                                                        }

                                                        if (sser.Group1Rank.HasValue)
                                                        {
                                                            row["一上學業類1排名"] = sser.Group1Rank.Value;
                                                            if (sser.Group1Count.HasValue)
                                                            {
                                                                row["一上學業類1排名母數"] = sser.Group1Count.Value;
                                                                row["一上學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                            }
                                                        }
                                                    }
                                                    if (sser.GradeYear == "1" && sser.Semester == "2")
                                                    {
                                                        if (sser.ClassRank.HasValue)
                                                        {
                                                            row["一下學業班排名"] = sser.ClassRank.Value;
                                                            if (sser.ClassCount.HasValue)
                                                            {
                                                                row["一下學業班排名母數"] = sser.ClassCount.Value;
                                                                row["一下學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                            }
                                                        }

                                                        if (sser.DeptRank.HasValue)
                                                        {
                                                            row["一下學業科排名"] = sser.DeptRank.Value;
                                                            if (sser.DeptCount.HasValue)
                                                            {
                                                                row["一下學業科排名母數"] = sser.DeptCount.Value;
                                                                row["一下學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                            }
                                                        }

                                                        if (sser.YearRank.HasValue)
                                                        {
                                                            row["一下學業校排名"] = sser.YearRank.Value;
                                                            if (sser.YearCount.HasValue)
                                                            {
                                                                row["一下學業校排名母數"] = sser.YearCount.Value;
                                                                row["一下學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                            }
                                                        }

                                                        if (sser.Group1Rank.HasValue)
                                                        {
                                                            row["一下學業類1排名"] = sser.Group1Rank.Value;
                                                            if (sser.Group1Count.HasValue)
                                                            {
                                                                row["一下學業類1排名母數"] = sser.Group1Count.Value;
                                                                row["一下學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                            }
                                                        }
                                                    }
                                                    if (sser.GradeYear == "2" && sser.Semester == "1")
                                                    {
                                                        if (sser.ClassRank.HasValue)
                                                        {
                                                            row["二上學業班排名"] = sser.ClassRank.Value;
                                                            if (sser.ClassCount.HasValue)
                                                            {
                                                                row["二上學業班排名母數"] = sser.ClassCount.Value;
                                                                row["二上學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                            }
                                                        }

                                                        if (sser.DeptRank.HasValue)
                                                        {
                                                            row["二上學業科排名"] = sser.DeptRank.Value;
                                                            if (sser.DeptCount.HasValue)
                                                            {
                                                                row["二上學業科排名母數"] = sser.DeptCount.Value;
                                                                row["二上學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                            }
                                                        }

                                                        if (sser.YearRank.HasValue)
                                                        {
                                                            row["二上學業校排名"] = sser.YearRank.Value;
                                                            if (sser.YearCount.HasValue)
                                                            {
                                                                row["二上學業校排名母數"] = sser.YearCount.Value;
                                                                row["二上學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                            }
                                                        }

                                                        if (sser.Group1Rank.HasValue)
                                                        {
                                                            row["二上學業類1排名"] = sser.Group1Rank.Value;
                                                            if (sser.Group1Count.HasValue)
                                                            {
                                                                row["二上學業類1排名母數"] = sser.Group1Count.Value;
                                                                row["二上學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                            }
                                                        }
                                                    }
                                                    if (sser.GradeYear == "2" && sser.Semester == "2")
                                                    {
                                                        if (sser.ClassRank.HasValue)
                                                        {
                                                            row["二下學業班排名"] = sser.ClassRank.Value;
                                                            if (sser.ClassCount.HasValue)
                                                            {
                                                                row["二下學業班排名母數"] = sser.ClassCount.Value;
                                                                row["二下學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                            }
                                                        }

                                                        if (sser.DeptRank.HasValue)
                                                        {
                                                            row["二下學業科排名"] = sser.DeptRank.Value;
                                                            if (sser.DeptCount.HasValue)
                                                            {
                                                                row["二下學業科排名母數"] = sser.DeptCount.Value;
                                                                row["二下學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                            }
                                                        }

                                                        if (sser.YearRank.HasValue)
                                                        {
                                                            row["二下學業校排名"] = sser.YearRank.Value;
                                                            if (sser.YearCount.HasValue)
                                                            {
                                                                row["二下學業校排名母數"] = sser.YearCount.Value;
                                                                row["二下學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                            }
                                                        }

                                                        if (sser.Group1Rank.HasValue)
                                                        {
                                                            row["二下學業類1排名"] = sser.Group1Rank.Value;
                                                            if (sser.Group1Count.HasValue)
                                                            {
                                                                row["二下學業類1排名母數"] = sser.Group1Count.Value;
                                                                row["二下學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                            }
                                                        }
                                                    }
                                                    if (sser.GradeYear == "3" && sser.Semester == "1")
                                                    {
                                                        if (sser.ClassRank.HasValue)
                                                        {
                                                            row["三上學業班排名"] = sser.ClassRank.Value;
                                                            if (sser.ClassCount.HasValue)
                                                            {
                                                                row["三上學業班排名母數"] = sser.ClassCount.Value;
                                                                row["三上學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                            }
                                                        }

                                                        if (sser.DeptRank.HasValue)
                                                        {
                                                            row["三上學業科排名"] = sser.DeptRank.Value;
                                                            if (sser.DeptCount.HasValue)
                                                            {
                                                                row["三上學業科排名母數"] = sser.DeptCount.Value;
                                                                row["三上學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                            }
                                                        }

                                                        if (sser.YearRank.HasValue)
                                                        {
                                                            row["三上學業校排名"] = sser.YearRank.Value;
                                                            if (sser.YearCount.HasValue)
                                                            {
                                                                row["三上學業校排名母數"] = sser.YearCount.Value;
                                                                row["三上學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                            }
                                                        }

                                                        if (sser.Group1Rank.HasValue)
                                                        {
                                                            row["三上學業類1排名"] = sser.Group1Rank.Value;
                                                            if (sser.Group1Count.HasValue)
                                                            {
                                                                row["三上學業類1排名母數"] = sser.Group1Count.Value;
                                                                row["三上學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                            }
                                                        }
                                                    }
                                                    if (sser.GradeYear == "3" && sser.Semester == "2")
                                                    {
                                                        if (sser.ClassRank.HasValue)
                                                        {
                                                            row["三下學業班排名"] = sser.ClassRank.Value;
                                                            if (sser.ClassCount.HasValue)
                                                            {
                                                                row["三下學業班排名母數"] = sser.ClassCount.Value;
                                                                row["三下學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                            }
                                                        }

                                                        if (sser.DeptRank.HasValue)
                                                        {
                                                            row["三下學業科排名"] = sser.DeptRank.Value;
                                                            if (sser.DeptCount.HasValue)
                                                            {
                                                                row["三下學業科排名母數"] = sser.DeptCount.Value;
                                                                row["三下學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                            }
                                                        }

                                                        if (sser.YearRank.HasValue)
                                                        {
                                                            row["三下學業校排名"] = sser.YearRank.Value;
                                                            if (sser.YearCount.HasValue)
                                                            {
                                                                row["三下學業校排名母數"] = sser.YearCount.Value;
                                                                row["三下學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                            }
                                                        }

                                                        if (sser.Group1Rank.HasValue)
                                                        {
                                                            row["三下學業類1排名"] = sser.Group1Rank.Value;
                                                            if (sser.Group1Count.HasValue)
                                                            {
                                                                row["三下學業類1排名母數"] = sser.Group1Count.Value;
                                                                row["三下學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                            }
                                                        }
                                                    }
                                                    if (sser.GradeYear == "4" && sser.Semester == "1")
                                                    {
                                                        if (sser.ClassRank.HasValue)
                                                        {
                                                            row["四上學業班排名"] = sser.ClassRank.Value;
                                                            if (sser.ClassCount.HasValue)
                                                            {
                                                                row["四上學業班排名母數"] = sser.ClassCount.Value;
                                                                row["四上學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                            }
                                                        }

                                                        if (sser.DeptRank.HasValue)
                                                        {
                                                            row["四上學業科排名"] = sser.DeptRank.Value;
                                                            if (sser.DeptCount.HasValue)
                                                            {
                                                                row["四上學業科排名母數"] = sser.DeptCount.Value;
                                                                row["四上學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                            }
                                                        }

                                                        if (sser.YearRank.HasValue)
                                                        {
                                                            row["四上學業校排名"] = sser.YearRank.Value;
                                                            if (sser.YearCount.HasValue)
                                                            {
                                                                row["四上學業校排名母數"] = sser.YearCount.Value;
                                                                row["四上學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                            }
                                                        }

                                                        if (sser.Group1Rank.HasValue)
                                                        {
                                                            row["四上學業類1排名"] = sser.Group1Rank.Value;
                                                            if (sser.Group1Count.HasValue)
                                                            {
                                                                row["四上學業類1排名母數"] = sser.Group1Count.Value;
                                                                row["四上學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                            }
                                                        }
                                                    }
                                                    if (sser.GradeYear == "4" && sser.Semester == "2")
                                                    {
                                                        if (sser.ClassRank.HasValue)
                                                        {
                                                            row["四下學業班排名"] = sser.ClassRank.Value;
                                                            if (sser.ClassCount.HasValue)
                                                            {
                                                                row["四下學業班排名母數"] = sser.ClassCount.Value;
                                                                row["四下學業班排名百分比"] = Utility.ParseRankPercent(sser.ClassRank.Value, sser.ClassCount.Value);
                                                            }
                                                        }

                                                        if (sser.DeptRank.HasValue)
                                                        {
                                                            row["四下學業科排名"] = sser.DeptRank.Value;
                                                            if (sser.DeptCount.HasValue)
                                                            {
                                                                row["四下學業科排名母數"] = sser.DeptCount.Value;
                                                                row["四下學業科排名百分比"] = Utility.ParseRankPercent(sser.DeptRank.Value, sser.DeptCount.Value);
                                                            }
                                                        }

                                                        if (sser.YearRank.HasValue)
                                                        {
                                                            row["四下學業校排名"] = sser.YearRank.Value;
                                                            if (sser.YearCount.HasValue)
                                                            {
                                                                row["四下學業校排名母數"] = sser.YearCount.Value;
                                                                row["四下學業校排名百分比"] = Utility.ParseRankPercent(sser.YearRank.Value, sser.YearCount.Value);
                                                            }
                                                        }

                                                        if (sser.Group1Rank.HasValue)
                                                        {
                                                            row["四下學業類1排名"] = sser.Group1Rank.Value;
                                                            if (sser.Group1Count.HasValue)
                                                            {
                                                                row["四下學業類1排名母數"] = sser.Group1Count.Value;
                                                                row["四下學業類1排名百分比"] = Utility.ParseRankPercent(sser.Group1Rank.Value, sser.Group1Count.Value);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion

                                            if (selectScore.ContainsKey(id1))
                                            {
                                                if (selectScore[id1].gsScore11.HasValue)
                                                    row["一上學業成績"] = Utility.NoRound(selectScore[id1].gsScore11.Value);

                                                if (selectScore[id1].gsScore12.HasValue)
                                                    row["一下學業成績"] = Utility.NoRound(selectScore[id1].gsScore12.Value);

                                                if (selectScore[id1].gsScore21.HasValue)
                                                    row["二上學業成績"] = Utility.NoRound(selectScore[id1].gsScore21.Value);

                                                if (selectScore[id1].gsScore22.HasValue)
                                                    row["二下學業成績"] = Utility.NoRound(selectScore[id1].gsScore22.Value);

                                                if (selectScore[id1].gsScore31.HasValue)
                                                    row["三上學業成績"] = Utility.NoRound(selectScore[id1].gsScore31.Value);

                                                if (selectScore[id1].gsScore32.HasValue)
                                                    row["三下學業成績"] = Utility.NoRound(selectScore[id1].gsScore32.Value);

                                                if (selectScore[id1].gsScore41.HasValue)
                                                    row["四上學業成績"] = Utility.NoRound(selectScore[id1].gsScore41.Value);

                                                if (selectScore[id1].gsScore42.HasValue)
                                                    row["四下學業成績"] = Utility.NoRound(selectScore[id1].gsScore42.Value);


                                                row["學業平均"] = Utility.NoRound(selectScore[id1].avgScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key2 = "學業成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1].avgScore) + 1;
                                                        row["學業平均班排名"] = rr;
                                                        row["學業平均班排名母數"] = ranks[key2].Count;
                                                        row["學業平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1].avgScore) + 1;
                                                        row["學業平均科排名"] = rr;
                                                        row["學業平均科排名母數"] = ranks[key2].Count;
                                                        row["學業平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業成績平均全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1].avgScore) + 1;
                                                        row["學業平均校排名"] = rr;
                                                        row["學業平均校排名母數"] = ranks[key2].Count;
                                                        row["學業平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key2 = "學業成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1].avgScoreC1) + 1;
                                                            row["學業平均類別一排名"] = rr;
                                                            row["學業平均類別一排名母數"] = ranks[key2].Count;
                                                            row["學業平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key2 = "學業成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1].avgScoreC2) + 1;
                                                            row["學業平均類別二排名"] = rr;
                                                            row["學業平均類別二排名母數"] = ranks[key2].Count;
                                                            row["學業平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                        }
                                                    }
                                                }
                                                row["學業總分"] = Utility.NoRound(selectScore[id1].sumScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key1 = "學業成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業總分班排名"] = ranks[key1].IndexOf(selectScore[id1].sumScore) + 1;
                                                        row["學業總分班排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業總分科排名"] = ranks[key1].IndexOf(selectScore[id1].sumScore) + 1;
                                                        row["學業總分科排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業成績總分全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業總分校排名"] = ranks[key1].IndexOf(selectScore[id1].sumScore) + 1;
                                                        row["學業總分校排名母數"] = ranks[key1].Count;
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key1 = "學業成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1].sumScoreC1) + 1;
                                                            row["學業總分類別一排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key1 = "學業成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1].sumScoreC2) + 1;
                                                            row["學業總分類別二排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion 處理學業

                                            // 處理學業原始
                                            #region 處理學業原始
                                            string id1_5 = studRec.StudentID + "學業原始";

                                            if (selectScore.ContainsKey(id1_5))
                                            {
                                                if (selectScore[id1_5].gsScore11.HasValue)
                                                    row["一上學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore11.Value);

                                                if (selectScore[id1_5].gsScore12.HasValue)
                                                    row["一下學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore12.Value);

                                                if (selectScore[id1_5].gsScore21.HasValue)
                                                    row["二上學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore21.Value);

                                                if (selectScore[id1_5].gsScore22.HasValue)
                                                    row["二下學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore22.Value);

                                                if (selectScore[id1_5].gsScore31.HasValue)
                                                    row["三上學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore31.Value);

                                                if (selectScore[id1_5].gsScore32.HasValue)
                                                    row["三下學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore32.Value);

                                                if (selectScore[id1_5].gsScore41.HasValue)
                                                    row["四上學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore41.Value);

                                                if (selectScore[id1_5].gsScore42.HasValue)
                                                    row["四下學業原始成績"] = Utility.NoRound(selectScore[id1_5].gsScore42.Value);


                                                row["學業原始平均"] = Utility.NoRound(selectScore[id1_5].avgScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key2 = "學業原始成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_5].avgScore) + 1;
                                                        row["學業原始平均班排名"] = rr;
                                                        row["學業原始平均班排名母數"] = ranks[key2].Count;
                                                        row["學業原始平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業原始成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_5].avgScore) + 1;
                                                        row["學業原始平均科排名"] = rr;
                                                        row["學業原始平均科排名母數"] = ranks[key2].Count;
                                                        row["學業原始平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業原始成績平均全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_5].avgScore) + 1;
                                                        row["學業原始平均校排名"] = rr;
                                                        row["學業原始平均校排名母數"] = ranks[key2].Count;
                                                        row["學業原始平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key2 = "學業原始成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1_5].avgScoreC1) + 1;
                                                            row["學業原始平均類別一排名"] = rr;
                                                            row["學業原始平均類別一排名母數"] = ranks[key2].Count;
                                                            row["學業原始平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key2 = "學業原始成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1_5].avgScoreC2) + 1;
                                                            row["學業原始平均類別二排名"] = rr;
                                                            row["學業原始平均類別二排名母數"] = ranks[key2].Count;
                                                            row["學業原始平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                        }
                                                    }
                                                }
                                                row["學業原始總分"] = Utility.NoRound(selectScore[id1_5].sumScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key1 = "學業原始成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業原始總分班排名"] = ranks[key1].IndexOf(selectScore[id1_5].sumScore) + 1;
                                                        row["學業原始總分班排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業原始成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業原始總分科排名"] = ranks[key1].IndexOf(selectScore[id1_5].sumScore) + 1;
                                                        row["學業原始總分科排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業原始成績總分全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業原始總分校排名"] = ranks[key1].IndexOf(selectScore[id1_5].sumScore) + 1;
                                                        row["學業原始總分校排名母數"] = ranks[key1].Count;
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key1 = "學業原始成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業原始總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1_5].sumScoreC1) + 1;
                                                            row["學業原始總分類別一排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key1 = "學業原始成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業原始總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1_5].sumScoreC2) + 1;
                                                            row["學業原始總分類別二排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion 處理學業原始

                                            // 處理學業體育
                                            #region 處理學業體育
                                            string id1_1 = studRec.StudentID + "學業體育";

                                            if (selectScore.ContainsKey(id1_1))
                                            {
                                                if (selectScore[id1_1].gsScore11.HasValue)
                                                    row["一上學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore11.Value);

                                                if (selectScore[id1_1].gsScore12.HasValue)
                                                    row["一下學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore12.Value);

                                                if (selectScore[id1_1].gsScore21.HasValue)
                                                    row["二上學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore21.Value);

                                                if (selectScore[id1_1].gsScore22.HasValue)
                                                    row["二下學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore22.Value);

                                                if (selectScore[id1_1].gsScore31.HasValue)
                                                    row["三上學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore31.Value);

                                                if (selectScore[id1_1].gsScore32.HasValue)
                                                    row["三下學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore32.Value);

                                                if (selectScore[id1_1].gsScore41.HasValue)
                                                    row["四上學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore41.Value);

                                                if (selectScore[id1_1].gsScore42.HasValue)
                                                    row["四下學業體育成績"] = Utility.NoRound(selectScore[id1_1].gsScore42.Value);


                                                row["學業體育平均"] = Utility.NoRound(selectScore[id1_1].avgScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key2 = "學業體育成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_1].avgScore) + 1;
                                                        row["學業體育平均班排名"] = rr;
                                                        row["學業體育平均班排名母數"] = ranks[key2].Count;
                                                        row["學業體育平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業體育成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_1].avgScore) + 1;
                                                        row["學業體育平均科排名"] = rr;
                                                        row["學業體育平均科排名母數"] = ranks[key2].Count;
                                                        row["學業體育平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業體育成績平均全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_1].avgScore) + 1;
                                                        row["學業體育平均校排名"] = rr;
                                                        row["學業體育平均校排名母數"] = ranks[key2].Count;
                                                        row["學業體育平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key2 = "學業體育成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1_1].avgScoreC1) + 1;
                                                            row["學業體育平均類別一排名"] = rr;
                                                            row["學業體育平均類別一排名母數"] = ranks[key2].Count;
                                                            row["學業體育平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key2 = "學業體育成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1_1].avgScoreC2) + 1;
                                                            row["學業體育平均類別二排名"] = rr;
                                                            row["學業體育平均類別二排名母數"] = ranks[key2].Count;
                                                            row["學業體育平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                        }
                                                    }
                                                }
                                                row["學業體育總分"] = Utility.NoRound(selectScore[id1_1].sumScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key1 = "學業體育成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業體育總分班排名"] = ranks[key1].IndexOf(selectScore[id1_1].sumScore) + 1;
                                                        row["學業體育總分班排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業體育成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業體育總分科排名"] = ranks[key1].IndexOf(selectScore[id1_1].sumScore) + 1;
                                                        row["學業體育總分科排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業體育成績總分全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業體育總分校排名"] = ranks[key1].IndexOf(selectScore[id1_1].sumScore) + 1;
                                                        row["學業體育總分校排名母數"] = ranks[key1].Count;
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key1 = "學業體育成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業體育總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1_1].sumScoreC1) + 1;
                                                            row["學業體育總分類別一排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key1 = "學業體育成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業體育總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1_1].sumScoreC2) + 1;
                                                            row["學業體育總分類別二排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion 處理學業體育

                                            // 處理學業健康與護理
                                            #region 處理學業健康與護理
                                            string id1_2 = studRec.StudentID + "學業健康與護理";

                                            if (selectScore.ContainsKey(id1_2))
                                            {
                                                if (selectScore[id1_2].gsScore11.HasValue)
                                                    row["一上學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore11.Value);

                                                if (selectScore[id1_2].gsScore12.HasValue)
                                                    row["一下學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore12.Value);

                                                if (selectScore[id1_2].gsScore21.HasValue)
                                                    row["二上學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore21.Value);

                                                if (selectScore[id1_2].gsScore22.HasValue)
                                                    row["二下學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore22.Value);

                                                if (selectScore[id1_2].gsScore31.HasValue)
                                                    row["三上學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore31.Value);

                                                if (selectScore[id1_2].gsScore32.HasValue)
                                                    row["三下學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore32.Value);

                                                if (selectScore[id1_2].gsScore41.HasValue)
                                                    row["四上學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore41.Value);

                                                if (selectScore[id1_2].gsScore42.HasValue)
                                                    row["四下學業健康與護理成績"] = Utility.NoRound(selectScore[id1_2].gsScore42.Value);


                                                row["學業健康與護理平均"] = Utility.NoRound(selectScore[id1_2].avgScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key2 = "學業健康與護理成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_2].avgScore) + 1;
                                                        row["學業健康與護理平均班排名"] = rr;
                                                        row["學業健康與護理平均班排名母數"] = ranks[key2].Count;
                                                        row["學業健康與護理平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業健康與護理成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_2].avgScore) + 1;
                                                        row["學業健康與護理平均科排名"] = rr;
                                                        row["學業健康與護理平均科排名母數"] = ranks[key2].Count;
                                                        row["學業健康與護理平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業健康與護理成績平均全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_2].avgScore) + 1;
                                                        row["學業健康與護理平均校排名"] = rr;
                                                        row["學業健康與護理平均校排名母數"] = ranks[key2].Count;
                                                        row["學業健康與護理平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key2 = "學業健康與護理成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1_2].avgScoreC1) + 1;
                                                            row["學業健康與護理平均類別一排名"] = rr;
                                                            row["學業健康與護理平均類別一排名母數"] = ranks[key2].Count;
                                                            row["學業健康與護理平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key2 = "學業健康與護理成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1_2].avgScoreC2) + 1;
                                                            row["學業健康與護理平均類別二排名"] = rr;
                                                            row["學業健康與護理平均類別二排名母數"] = ranks[key2].Count;
                                                            row["學業健康與護理平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                        }
                                                    }
                                                }
                                                row["學業健康與護理總分"] = Utility.NoRound(selectScore[id1_2].sumScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key1 = "學業健康與護理成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業健康與護理總分班排名"] = ranks[key1].IndexOf(selectScore[id1_2].sumScore) + 1;
                                                        row["學業健康與護理總分班排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業健康與護理成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業健康與護理總分科排名"] = ranks[key1].IndexOf(selectScore[id1_2].sumScore) + 1;
                                                        row["學業健康與護理總分科排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業健康與護理成績總分全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業健康與護理總分校排名"] = ranks[key1].IndexOf(selectScore[id1_2].sumScore) + 1;
                                                        row["學業健康與護理總分校排名母數"] = ranks[key1].Count;
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key1 = "學業健康與護理成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業健康與護理總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1_2].sumScoreC1) + 1;
                                                            row["學業健康與護理總分類別一排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key1 = "學業健康與護理成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業健康與護理總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1_2].sumScoreC2) + 1;
                                                            row["學業健康與護理總分類別二排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion 處理學業健康與護理

                                            // 處理學業國防通識
                                            #region 處理學業國防通識
                                            string id1_3 = studRec.StudentID + "學業國防通識";

                                            if (selectScore.ContainsKey(id1_3))
                                            {
                                                if (selectScore[id1_3].gsScore11.HasValue)
                                                    row["一上學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore11.Value);

                                                if (selectScore[id1_3].gsScore12.HasValue)
                                                    row["一下學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore12.Value);

                                                if (selectScore[id1_3].gsScore21.HasValue)
                                                    row["二上學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore21.Value);

                                                if (selectScore[id1_3].gsScore22.HasValue)
                                                    row["二下學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore22.Value);

                                                if (selectScore[id1_3].gsScore31.HasValue)
                                                    row["三上學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore31.Value);

                                                if (selectScore[id1_3].gsScore32.HasValue)
                                                    row["三下學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore32.Value);

                                                if (selectScore[id1_3].gsScore41.HasValue)
                                                    row["四上學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore41.Value);

                                                if (selectScore[id1_3].gsScore42.HasValue)
                                                    row["四下學業國防通識成績"] = Utility.NoRound(selectScore[id1_3].gsScore42.Value);


                                                row["學業國防通識平均"] = Utility.NoRound(selectScore[id1_3].avgScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key2 = "學業國防通識成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_3].avgScore) + 1;
                                                        row["學業國防通識平均班排名"] = rr;
                                                        row["學業國防通識平均班排名母數"] = ranks[key2].Count;
                                                        row["學業國防通識平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業國防通識成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_3].avgScore) + 1;
                                                        row["學業國防通識平均科排名"] = rr;
                                                        row["學業國防通識平均科排名母數"] = ranks[key2].Count;
                                                        row["學業國防通識平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業國防通識成績平均全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_3].avgScore) + 1;
                                                        row["學業國防通識平均校排名"] = rr;
                                                        row["學業國防通識平均校排名母數"] = ranks[key2].Count;
                                                        row["學業國防通識平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key2 = "學業國防通識成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1_3].avgScoreC1) + 1;
                                                            row["學業國防通識平均類別一排名"] = rr;
                                                            row["學業國防通識平均類別一排名母數"] = ranks[key2].Count;
                                                            row["學業國防通識平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key2 = "學業國防通識成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1_3].avgScoreC2) + 1;
                                                            row["學業國防通識平均類別二排名"] = rr;
                                                            row["學業國防通識平均類別二排名母數"] = ranks[key2].Count;
                                                            row["學業國防通識平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                        }
                                                    }
                                                }
                                                row["學業國防通識總分"] = Utility.NoRound(selectScore[id1_3].sumScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key1 = "學業國防通識成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業國防通識總分班排名"] = ranks[key1].IndexOf(selectScore[id1_3].sumScore) + 1;
                                                        row["學業國防通識總分班排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業國防通識成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業國防通識總分科排名"] = ranks[key1].IndexOf(selectScore[id1_3].sumScore) + 1;
                                                        row["學業國防通識總分科排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業國防通識成績總分全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業國防通識總分校排名"] = ranks[key1].IndexOf(selectScore[id1_3].sumScore) + 1;
                                                        row["學業國防通識總分校排名母數"] = ranks[key1].Count;
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key1 = "學業國防通識成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業國防通識總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1_3].sumScoreC1) + 1;
                                                            row["學業國防通識總分類別一排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key1 = "學業國防通識成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業國防通識總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1_3].sumScoreC2) + 1;
                                                            row["學業國防通識總分類別二排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion 處理學業國防通識

                                            // 處理學業實習科目
                                            #region 處理學業實習科目
                                            string id1_4 = studRec.StudentID + "學業實習科目";

                                            if (selectScore.ContainsKey(id1_4))
                                            {
                                                if (selectScore[id1_4].gsScore11.HasValue)
                                                    row["一上學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore11.Value);

                                                if (selectScore[id1_4].gsScore12.HasValue)
                                                    row["一下學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore12.Value);

                                                if (selectScore[id1_4].gsScore21.HasValue)
                                                    row["二上學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore21.Value);

                                                if (selectScore[id1_4].gsScore22.HasValue)
                                                    row["二下學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore22.Value);

                                                if (selectScore[id1_4].gsScore31.HasValue)
                                                    row["三上學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore31.Value);

                                                if (selectScore[id1_4].gsScore32.HasValue)
                                                    row["三下學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore32.Value);

                                                if (selectScore[id1_4].gsScore41.HasValue)
                                                    row["四上學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore41.Value);

                                                if (selectScore[id1_4].gsScore42.HasValue)
                                                    row["四下學業實習科目成績"] = Utility.NoRound(selectScore[id1_4].gsScore42.Value);


                                                row["學業實習科目平均"] = Utility.NoRound(selectScore[id1_4].avgScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key2 = "學業實習科目成績平均班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_4].avgScore) + 1;
                                                        row["學業實習科目平均班排名"] = rr;
                                                        row["學業實習科目平均班排名母數"] = ranks[key2].Count;
                                                        row["學業實習科目平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業實習科目成績平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_4].avgScore) + 1;
                                                        row["學業實習科目平均科排名"] = rr;
                                                        row["學業實習科目平均科排名母數"] = ranks[key2].Count;
                                                        row["學業實習科目平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    key2 = "學業實習科目成績平均全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        int rr = ranks[key2].IndexOf(selectScore[id1_4].avgScore) + 1;
                                                        row["學業實習科目平均校排名"] = rr;
                                                        row["學業實習科目平均校排名母數"] = ranks[key2].Count;
                                                        row["學業實習科目平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key2 = "學業實習科目成績平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1_4].avgScoreC1) + 1;
                                                            row["學業實習科目平均類別一排名"] = rr;
                                                            row["學業實習科目平均類別一排名母數"] = ranks[key2].Count;
                                                            row["學業實習科目平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);

                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key2 = "學業實習科目成績平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key2))
                                                        {
                                                            int rr = ranks[key2].IndexOf(selectScore[id1_4].avgScoreC2) + 1;
                                                            row["學業實習科目平均類別二排名"] = rr;
                                                            row["學業實習科目平均類別二排名母數"] = ranks[key2].Count;
                                                            row["學業實習科目平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key2].Count);
                                                        }
                                                    }
                                                }
                                                row["學業實習科目總分"] = Utility.NoRound(selectScore[id1_4].sumScore);
                                                if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                {
                                                    string key1 = "學業實習科目成績總分班排名" + studRec.RefClass.ClassID + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業實習科目總分班排名"] = ranks[key1].IndexOf(selectScore[id1_4].sumScore) + 1;
                                                        row["學業實習科目總分班排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業實習科目成績總分科排名" + studRec.Department + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業實習科目總分科排名"] = ranks[key1].IndexOf(selectScore[id1_4].sumScore) + 1;
                                                        row["學業實習科目總分科排名母數"] = ranks[key1].Count;
                                                    }

                                                    key1 = "學業實習科目成績總分全校排名" + gradeyear + "^^^";
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["學業實習科目總分校排名"] = ranks[key1].IndexOf(selectScore[id1_4].sumScore) + 1;
                                                        row["學業實習科目總分校排名母數"] = ranks[key1].Count;
                                                    }

                                                    if (studRec.Fields.ContainsKey("tag1"))
                                                    {
                                                        key1 = "學業實習科目成績總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業實習科目總分類別一排名"] = ranks[key1].IndexOf(selectScore[id1_4].sumScoreC1) + 1;
                                                            row["學業實習科目總分類別一排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                    if (studRec.Fields.ContainsKey("tag2"))
                                                    {
                                                        key1 = "學業實習科目成績總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key1))
                                                        {
                                                            row["學業實習科目總分類別二排名"] = ranks[key1].IndexOf(selectScore[id1_4].sumScoreC2) + 1;
                                                            row["學業實習科目總分類別二排名母數"] = ranks[key1].Count;
                                                        }
                                                    }
                                                }
                                            }
                                            #endregion 處理學業實習科目

                                        }
                                        // 處理總計
                                        #region 處理總計
                                        string id2 = studRec.StudentID + "總計成績";
                                        if (selectScore.ContainsKey(id2))
                                        {

                                            row["總計加權平均"] = Utility.NoRound(selectScore[id2].avgScoreA);
                                            if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                            {
                                                string key4 = "總計加權平均班排名" + studRec.RefClass.ClassID;
                                                if (ranks.ContainsKey(key4))
                                                {
                                                    row["總計加權平均班排名"] = ranks[key4].IndexOf(selectScore[id2].avgScoreA) + 1;
                                                    row["總計加權平均班排名母數"] = ranks[key4].Count;
                                                }
                                                if (ranks.ContainsKey(key4))
                                                {
                                                    row["總計加權平均科排名"] = ranks[key4].IndexOf(selectScore[id2].avgScoreA) + 1;
                                                    row["總計加權平均科排名母數"] = ranks[key4].Count;
                                                }
                                                if (ranks.ContainsKey(key4))
                                                {
                                                    row["總計加權平均校排名"] = ranks[key4].IndexOf(selectScore[id2].avgScoreA) + 1;
                                                    row["總計加權平均校排名母數"] = ranks[key4].Count;
                                                }
                                                if (studRec.Fields.ContainsKey("tag1"))
                                                {
                                                    if (ranks.ContainsKey(key4))
                                                    {
                                                        row["總計加權平均類別一"] = Utility.NoRound(selectScore[id2].avgScoreAC1);
                                                        row["總計加權平均類別一排名"] = ranks[key4].IndexOf(selectScore[id2].avgScoreAC1) + 1;
                                                        row["總計加權平均類別一排名母數"] = ranks[key4].Count;
                                                    }
                                                }

                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    if (ranks.ContainsKey(key4))
                                                    {
                                                        row["總計加權平均類別二"] = selectScore[id2].avgScoreAC2;
                                                        row["總計加權平均類別二排名"] = ranks[key4].IndexOf(selectScore[id2].avgScoreAC2) + 1;
                                                        row["總計加權平均類別二排名母數"] = ranks[key4].Count;
                                                    }
                                                }
                                            }

                                            row["總計加權總分"] = Utility.NoRound(selectScore[id2].sumScoreA);

                                            if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                            {
                                                string key2 = "總計加權總分班排名" + studRec.RefClass.ClassID;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    row["總計加權總分班排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreA) + 1; ;
                                                    row["總計加權總分班排名母數"] = ranks[key2].Count;
                                                }

                                                key2 = "總計加權總分科排名" + studRec.Department + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    row["總計加權總分科排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreA) + 1; ;
                                                    row["總計加權總分科排名母數"] = ranks[key2].Count;
                                                }

                                                key2 = "總計加權總分全校排名" + gradeyear;
                                                if (ranks.ContainsKey(key2))
                                                {
                                                    row["總計加權總分校排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreA) + 1; ;
                                                    row["總計加權總分校排名母數"] = ranks[key2].Count;
                                                }

                                                if (studRec.Fields.ContainsKey("tag1"))
                                                {
                                                    key2 = "總計加權總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        row["總計加權總分類別一"] = Utility.NoRound(selectScore[id2].sumScoreAC1);
                                                        row["總計加權總分類別一排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreAC1) + 1; ;
                                                        row["總計加權總分類別一排名母數"] = ranks[key2].Count;
                                                    }
                                                }

                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    key2 = "總計加權總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        row["總計加權總分類別二"] = Utility.NoRound(selectScore[id2].sumScoreAC2);
                                                        row["總計加權總分類別二排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreAC2) + 1; ;
                                                        row["總計加權總分類別二排名母數"] = ranks[key2].Count;
                                                    }
                                                }
                                            }

                                            row["總計平均"] = Utility.NoRound(selectScore[id2].avgScore);
                                            if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                            {
                                                string key3 = "總計平均班排名" + studRec.RefClass.ClassID;
                                                if (ranks.ContainsKey(key3))
                                                {
                                                    row["總計平均班排名"] = ranks[key3].IndexOf(selectScore[id2].avgScore) + 1;
                                                    row["總計平均班排名母數"] = ranks[key3].Count;
                                                }

                                                key3 = "總計平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key3))
                                                {
                                                    row["總計平均科排名"] = ranks[key3].IndexOf(selectScore[id2].avgScore) + 1;
                                                    row["總計平均科排名母數"] = ranks[key3].Count;
                                                }

                                                key3 = "總計平均全校排名" + gradeyear;
                                                if (ranks.ContainsKey(key3))
                                                {
                                                    row["總計平均校排名"] = ranks[key3].IndexOf(selectScore[id2].avgScore) + 1;
                                                    row["總計平均校排名母數"] = ranks[key3].Count;
                                                }

                                                if (studRec.Fields.ContainsKey("tag1"))
                                                {
                                                    key3 = "總計平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key3))
                                                    {
                                                        row["總計平均類別一"] = Utility.NoRound(selectScore[id2].avgScoreC1);
                                                        row["總計平均類別一排名"] = ranks[key3].IndexOf(selectScore[id2].avgScoreC1) + 1;
                                                        row["總計平均類別一排名母數"] = ranks[key3].Count;
                                                    }
                                                }

                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    key3 = "總計平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key3))
                                                    {
                                                        row["總計平均類別二"] = Utility.NoRound(selectScore[id2].avgScoreC2);
                                                        row["總計平均類別二排名"] = ranks[key3].IndexOf(selectScore[id2].avgScoreC2) + 1;
                                                        row["總計平均類別二排名母數"] = ranks[key3].Count;
                                                    }
                                                }
                                            }
                                            row["總計總分"] = Utility.NoRound(selectScore[id2].sumScore);
                                            if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                            {
                                                string key1 = "總計總分班排名" + studRec.RefClass.ClassID;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["總計總分班排名"] = ranks[key1].IndexOf(selectScore[id2].sumScore) + 1;
                                                    row["總計總分班排名母數"] = ranks[key1].Count;
                                                }

                                                key1 = "總計總分科排名" + studRec.Department + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["總計總分科排名"] = ranks[key1].IndexOf(selectScore[id2].sumScore) + 1;
                                                    row["總計總分科排名母數"] = ranks[key1].Count;
                                                }

                                                key1 = "總計總分全校排名" + gradeyear;
                                                if (ranks.ContainsKey(key1))
                                                {
                                                    row["總計總分校排名"] = ranks[key1].IndexOf(selectScore[id2].sumScore) + 1;
                                                    row["總計總分校排名母數"] = ranks[key1].Count;
                                                }

                                                if (studRec.Fields.ContainsKey("tag1"))
                                                {
                                                    key1 = "總計總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["總計總分類別一"] = Utility.NoRound(selectScore[id2].sumScoreC1);
                                                        row["總計總分類別一排名"] = ranks[key1].IndexOf(selectScore[id2].sumScoreC1) + 1;
                                                        row["總計總分類別一排名母數"] = ranks[key1].Count;
                                                    }
                                                }
                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    key1 = "總計總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["總計總分類別二"] = Utility.NoRound(selectScore[id2].sumScoreC2);
                                                        row["總計總分類別二排名"] = ranks[key1].IndexOf(selectScore[id2].sumScoreC2) + 1;
                                                        row["總計總分類別二排名母數"] = ranks[key1].Count;
                                                    }
                                                }
                                            }
                                        }
                                        #endregion 處理總計

                                        // 處理科目原始成績加權平均
                                        #region 處理科目原始成績加權平均
                                        string key5_name = "科目原始成績加權平均";
                                        string key7_name = "篩選科目原始成績加權平均";
                                        for (int g = 1; g <= 4; g++)
                                        {
                                            string key5 = "", key7 = "";
                                            string gsS = "";
                                            for (int s = 1; s <= 2; s++)
                                            {
                                                if (g == 1 && s == 1) gsS = "一上";
                                                if (g == 1 && s == 2) gsS = "一下";
                                                if (g == 2 && s == 1) gsS = "二上";
                                                if (g == 2 && s == 2) gsS = "二下";
                                                if (g == 3 && s == 1) gsS = "三上";
                                                if (g == 3 && s == 2) gsS = "三下";
                                                if (g == 4 && s == 1) gsS = "四上";
                                                if (g == 4 && s == 2) gsS = "四下";
                                                key5 = studRec.StudentID + gsS + key5_name;
                                                key7 = studRec.StudentID + gsS + key7_name;

                                                string id5 = studRec.StudentID + gsS + key5_name;
                                                string id7 = studRec.StudentID + gsS + key7_name;

                                                #region 處理科目原始成績
                                                if (selectScore.ContainsKey(id5))
                                                {
                                                    row[gsS + "科目原始成績加權平均"] = Utility.NoRound(selectScore[id5].avgScoreA);
                                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                    {
                                                        string key5a = gsS + "科目原始成績加權平均班排名" + studRec.RefClass.ClassID;
                                                        if (ranks.ContainsKey(key5a))
                                                        {
                                                            int rr = ranks[key5a].IndexOf(selectScore[id5].avgScoreA) + 1;
                                                            row[gsS + "科目原始成績加權平均班排名"] = rr;
                                                            row[gsS + "科目原始成績加權平均班排名母數"] = ranks[key5a].Count;
                                                            row[gsS + "科目原始成績加權平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                        }

                                                        key5a = gsS + "科目原始成績加權平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key5a))
                                                        {
                                                            int rr = ranks[key5a].IndexOf(selectScore[id5].avgScoreA) + 1;
                                                            row[gsS + "科目原始成績加權平均科排名"] = rr;
                                                            row[gsS + "科目原始成績加權平均科排名母數"] = ranks[key5a].Count;
                                                            row[gsS + "科目原始成績加權平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                        }

                                                        key5a = gsS + "科目原始成績加權平均校排名" + gradeyear;
                                                        if (ranks.ContainsKey(key5a))
                                                        {
                                                            int rr = ranks[key5a].IndexOf(selectScore[id5].avgScoreA) + 1;
                                                            row[gsS + "科目原始成績加權平均校排名"] = rr;
                                                            row[gsS + "科目原始成績加權平均校排名母數"] = ranks[key5a].Count;
                                                            row[gsS + "科目原始成績加權平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                        }
                                                        if (studRec.Fields.ContainsKey("tag1"))
                                                        {
                                                            key5a = gsS + "科目原始成績加權平均類別一" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                            if (ranks.ContainsKey(key5a))
                                                            {
                                                                int rr = ranks[key5a].IndexOf(selectScore[id5].avgScoreAC1) + 1;
                                                                row[gsS + "科目原始成績加權平均類別一"] = Utility.NoRound(selectScore[id5].avgScoreAC1);
                                                                row[gsS + "科目原始成績加權平均類別一排名"] = rr;
                                                                row[gsS + "科目原始成績加權平均類別一排名母數"] = ranks[key5a].Count;
                                                                row[gsS + "科目原始成績加權平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                            }
                                                        }

                                                        if (studRec.Fields.ContainsKey("tag2"))
                                                        {
                                                            key5a = gsS + "科目原始成績加權平均類別二" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                            if (ranks.ContainsKey(key5a))
                                                            {
                                                                int rr = ranks[key5a].IndexOf(selectScore[id5].avgScoreAC2) + 1;
                                                                row[gsS + "科目原始成績加權平均類別二"] = Utility.NoRound(selectScore[id5].avgScoreAC2);
                                                                row[gsS + "科目原始成績加權平均類別二排名"] = rr;
                                                                row[gsS + "科目原始成績加權平均類別二排名母數"] = ranks[key5a].Count;
                                                                row[gsS + "科目原始成績加權平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion 處理科目原始成績

                                                #region 處理篩選科目原始成績
                                                if (selectScore.ContainsKey(id7))
                                                {
                                                    row[gsS + "篩選科目原始成績加權平均"] = Utility.NoRound(selectScore[id7].avgScoreA);

                                                    if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                                    {
                                                        string key7a = gsS + "篩選科目原始成績加權平均班排名" + studRec.RefClass.ClassID;
                                                        if (ranks.ContainsKey(key7a))
                                                        {
                                                            int rr = ranks[key7a].IndexOf(selectScore[id7].avgScoreA) + 1;
                                                            row[gsS + "篩選科目原始成績加權平均班排名"] = rr;
                                                            row[gsS + "篩選科目原始成績加權平均班排名母數"] = ranks[key7a].Count;
                                                            row[gsS + "篩選科目原始成績加權平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
                                                        }

                                                        key7a = gsS + "篩選科目原始成績加權平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                        if (ranks.ContainsKey(key7a))
                                                        {
                                                            int rr = ranks[key7a].IndexOf(selectScore[id7].avgScoreA) + 1;
                                                            row[gsS + "篩選科目原始成績加權平均科排名"] = rr;
                                                            row[gsS + "篩選科目原始成績加權平均科排名母數"] = ranks[key7a].Count;
                                                            row[gsS + "篩選科目原始成績加權平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
                                                        }

                                                        key7a = gsS + "篩選科目原始成績加權平均校排名" + gradeyear;
                                                        if (ranks.ContainsKey(key7a))
                                                        {
                                                            int rr = ranks[key7a].IndexOf(selectScore[id7].avgScoreA) + 1;
                                                            row[gsS + "篩選科目原始成績加權平均校排名"] = rr;
                                                            row[gsS + "篩選科目原始成績加權平均校排名母數"] = ranks[key7a].Count;
                                                            row[gsS + "篩選科目原始成績加權平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
                                                        }

                                                        if (studRec.Fields.ContainsKey("tag1"))
                                                        {
                                                            key7a = gsS + "篩選科目原始成績加權平均類別一" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                            if (ranks.ContainsKey(key7a))
                                                            {
                                                                int rr = ranks[key7a].IndexOf(selectScore[id7].avgScoreAC1) + 1;
                                                                row[gsS + "篩選科目原始成績加權平均類別一"] = Utility.NoRound(selectScore[id7].avgScoreAC1);
                                                                row[gsS + "篩選科目原始成績加權平均類別一排名"] = rr;
                                                                row[gsS + "篩選科目原始成績加權平均類別一排名母數"] = ranks[key7a].Count;
                                                                row[gsS + "篩選科目原始成績加權平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
                                                            }
                                                        }

                                                        if (studRec.Fields.ContainsKey("tag2"))
                                                        {
                                                            key7a = gsS + "篩選科目原始成績加權平均類別二" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                            if (ranks.ContainsKey(key7a))
                                                            {
                                                                int rr = ranks[key7a].IndexOf(selectScore[id7].avgScoreAC2) + 1;
                                                                row[gsS + "篩選科目原始成績加權平均類別二"] = Utility.NoRound(selectScore[id7].avgScoreAC2);
                                                                row[gsS + "篩選科目原始成績加權平均類別二排名"] = rr;
                                                                row[gsS + "篩選科目原始成績加權平均類別二排名母數"] = ranks[key7a].Count;
                                                                row[gsS + "篩選科目原始成績加權平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion 處理篩選科目原始成績
                                            }
                                        }
                                        #endregion 處理科目原始成績加權平均

                                        // 處理科目原始成績加權平均平均
                                        #region 處理科目原始成績加權平均平均
                                        string key6_name = "科目原始成績加權平均平均";
                                        string key8_name = "篩選科目原始成績加權平均平均";
                                        string key6 = studRec.StudentID + key6_name;
                                        string key8 = studRec.StudentID + key8_name;

                                        string id6 = studRec.StudentID + key6_name;
                                        string id8 = studRec.StudentID + key8_name;

                                        #region 處理科目原始成績
                                        if (selectScore.ContainsKey(id6))
                                        {

                                            row["科目原始成績加權平均平均"] = Utility.NoRound(selectScore[id6].avgScoreA);
                                            if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                            {
                                                string key6a = "科目原始成績加權平均平均班排名" + studRec.RefClass.ClassID;
                                                if (ranks.ContainsKey(key6a))
                                                {
                                                    int rr = ranks[key6a].IndexOf(selectScore[id6].avgScoreA) + 1;
                                                    row["科目原始成績加權平均平均班排名"] = rr;
                                                    row["科目原始成績加權平均平均班排名母數"] = ranks[key6a].Count;
                                                    row["科目原始成績加權平均平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                                }

                                                key6a = "科目原始成績加權平均平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key6a))
                                                {
                                                    int rr = ranks[key6a].IndexOf(selectScore[id6].avgScoreA) + 1;
                                                    row["科目原始成績加權平均平均科排名"] = rr;
                                                    row["科目原始成績加權平均平均科排名母數"] = ranks[key6a].Count;
                                                    row["科目原始成績加權平均平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                                }

                                                key6a = "科目原始成績加權平均平均校排名" + gradeyear;
                                                if (ranks.ContainsKey(key6a))
                                                {
                                                    int rr = ranks[key6a].IndexOf(selectScore[id6].avgScoreA) + 1;
                                                    row["科目原始成績加權平均平均校排名"] = rr;
                                                    row["科目原始成績加權平均平均校排名母數"] = ranks[key6a].Count;
                                                    row["科目原始成績加權平均平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                                }
                                                if (studRec.Fields.ContainsKey("tag1"))
                                                {
                                                    key6a = "科目原始成績加權平均平均類別一" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key6a))
                                                    {
                                                        int rr = ranks[key6a].IndexOf(selectScore[id6].avgScoreAC1) + 1;
                                                        row["科目原始成績加權平均平均類別一"] = Utility.NoRound(selectScore[id6].avgScoreAC1);
                                                        row["科目原始成績加權平均平均類別一排名"] = rr;
                                                        row["科目原始成績加權平均平均類別一排名母數"] = ranks[key6a].Count;
                                                        row["科目原始成績加權平均平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                                    }
                                                }

                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    key6a = "科目原始成績加權平均平均類別二" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key6a))
                                                    {
                                                        int rr = ranks[key6a].IndexOf(selectScore[id6].avgScoreAC2) + 1;
                                                        row["科目原始成績加權平均平均類別二"] = Utility.NoRound(selectScore[id6].avgScoreAC2);
                                                        row["科目原始成績加權平均平均類別二排名"] = rr;
                                                        row["科目原始成績加權平均平均類別二排名母數"] = ranks[key6a].Count;
                                                        row["科目原始成績加權平均平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                                    }
                                                }
                                            }
                                        }
                                        #endregion 處理科目原始成績

                                        #region 處理篩選科目原始成績
                                        if (selectScore.ContainsKey(id8))
                                        {
                                            row["篩選科目原始成績加權平均平均"] = Utility.NoRound(selectScore[id8].avgScoreA);
                                            if (!noRankList.Contains(studRec.StudentID))//不是不排名學生
                                            {
                                                string key8a = "篩選科目原始成績加權平均平均班排名" + studRec.RefClass.ClassID;
                                                if (ranks.ContainsKey(key8a))
                                                {
                                                    int rr = ranks[key8a].IndexOf(selectScore[id8].avgScoreA) + 1;
                                                    row["篩選科目原始成績加權平均平均班排名"] = rr;
                                                    row["篩選科目原始成績加權平均平均班排名母數"] = ranks[key8a].Count;
                                                    row["篩選科目原始成績加權平均平均班排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                                }

                                                key8a = "篩選科目原始成績加權平均平均科排名" + studRec.Department + "^^^" + gradeyear;
                                                if (ranks.ContainsKey(key8a))
                                                {
                                                    int rr = ranks[key8a].IndexOf(selectScore[id8].avgScoreA) + 1;
                                                    row["篩選科目原始成績加權平均平均科排名"] = rr;
                                                    row["篩選科目原始成績加權平均平均科排名母數"] = ranks[key8a].Count;
                                                    row["篩選科目原始成績加權平均平均科排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                                }

                                                key8a = "篩選科目原始成績加權平均平均校排名" + gradeyear;
                                                if (ranks.ContainsKey(key8a))
                                                {
                                                    int rr = ranks[key8a].IndexOf(selectScore[id8].avgScoreA) + 1;
                                                    row["篩選科目原始成績加權平均平均校排名"] = rr;
                                                    row["篩選科目原始成績加權平均平均校排名母數"] = ranks[key8a].Count;
                                                    row["篩選科目原始成績加權平均平均校排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                                }
                                                if (studRec.Fields.ContainsKey("tag1"))
                                                {
                                                    key8a = "篩選科目原始成績加權平均平均類別一" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key8a))
                                                    {
                                                        int rr = ranks[key8a].IndexOf(selectScore[id8].avgScoreAC1) + 1;
                                                        row["篩選科目原始成績加權平均平均類別一"] = Utility.NoRound(selectScore[id8].avgScoreAC1);
                                                        row["篩選科目原始成績加權平均平均類別一排名"] = rr;
                                                        row["篩選科目原始成績加權平均平均類別一排名母數"] = ranks[key8a].Count;
                                                        row["篩選科目原始成績加權平均平均類別一排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                                    }
                                                }

                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    key8a = "篩選科目原始成績加權平均平均類別二" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key8a))
                                                    {
                                                        int rr = ranks[key8a].IndexOf(selectScore[id8].avgScoreAC2) + 1;
                                                        row["篩選科目原始成績加權平均平均類別二"] = Utility.NoRound(selectScore[id8].avgScoreAC2);
                                                        row["篩選科目原始成績加權平均平均類別二排名"] = rr;
                                                        row["篩選科目原始成績加權平均平均類別二排名母數"] = ranks[key8a].Count;
                                                        row["篩選科目原始成績加權平均平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                                    }
                                                }
                                            }
                                        }
                                        #endregion 處理篩選科目原始成績
                                        #endregion 處理科目原始成績加權平均平均

                                        // 不排名不放入
                                        if (!noRankList.Contains(studRec.StudentID))
                                            _table.Rows.Add(row);

                                    }
                                } // data row

                                if (OneClassCompleted != null)
                                    OneClassCompleted();

                                //List<string> fields = new List<string>(docTemplate.MailMerge.GetFieldNames());
                                //List<string> rmColumns = new List<string>();

                                //foreach (DataColumn dc in _table.Columns)
                                //{
                                //    if (!fields.Contains(dc.ColumnName))
                                //    {
                                //        rmColumns.Add(dc.ColumnName);
                                //    }
                                //}
                                //// 檢查有資料才Merge
                                //foreach (string str in rmColumns)
                                //    _table.Columns.Remove(str);

                                //GC.Collect();

                                //// debug 用
                                //_table.TableName = "debug";
                                //string pathAA = System.Windows.Forms.Application.StartupPath + "\\多學期成績單debug.xml";
                                //_table.WriteXml(pathAA);

                                // 當 table 有資料再合併
                                if (_table.Rows.Count > 0)
                                {

                                    Aspose.Words.Document document = new Aspose.Words.Document();
                                    document = docTemplate;
                                    doc.Sections.Add(doc.ImportNode(document.Sections[0], true));

                                    doc.MailMerge.Execute(_table);
                                    doc.MailMerge.RemoveEmptyParagraphs = true;
                                    doc.MailMerge.DeleteFields();

                                    _table.Rows.Clear();

                                    #region Word 存檔
                                    string reportNameW = "W_" + className + "-多學期科目成績固定排名成績單";
                                    string pathW = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports", FolderName);
                                    if (!Directory.Exists(pathW))
                                        Directory.CreateDirectory(pathW);
                                    pathW = Path.Combine(pathW, reportNameW + ".doc");

                                    if (File.Exists(pathW))
                                    {
                                        int i = 1;
                                        while (true)
                                        {
                                            string newPathW = Path.GetDirectoryName(pathW) + "\\" + Path.GetFileNameWithoutExtension(pathW) + (i++) + Path.GetExtension(pathW);
                                            if (!File.Exists(newPathW))
                                            {
                                                pathW = newPathW;
                                                break;
                                            }
                                        }
                                    }


                                    try
                                    {
                                        if (setting.CheckExportStudent)
                                        {
                                            doc.Save(pathW, Aspose.Words.SaveFormat.Doc);
                                        }

                                        int xx = (int)(100d / ClassSumCount * ClassCountStart);
                                        FISCA.RTContext.Invoke(new Action<string, int>(Word_Msg), new object[] { "產生班級Word檔中...", xx });
                                        ClassCountStart++;

                                    }
                                    catch (OutOfMemoryException exow)
                                    {
                                        exc = exow;
                                    }
                                    doc = null;
                                    GC.Collect();
                                    #endregion
                                }

                            }// doc
                            #endregion
                        }


                        bkw.ReportProgress(99);

                    }

                }
                catch (OutOfMemoryException exo)
                {
                    exc = exo;
                }
                catch (Exception ex)
                {
                    exc = ex;
                }
            };
            bkw.RunWorkerAsync();
        }


        private static void PDF_Msg(string msg, int p)
        {
            MotherForm.SetStatusBarMessage(msg, p);
        }

        private static void Word_Msg(string msg, int p)
        {
            MotherForm.SetStatusBarMessage(msg, p);
        }

        //static void MailMerge_MergeImageField(object sender, Aspose.Words.Reporting.MergeImageFieldEventArgs e)
        //{

        //}

        public static Document getMergeTable()
        {
            return new Document(new MemoryStream(Properties.Resources.合併欄位總表));
        }
    }
}
