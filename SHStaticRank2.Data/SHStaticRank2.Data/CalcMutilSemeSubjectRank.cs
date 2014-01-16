using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.Customization.Data;
using System.IO;
using FISCA.DSAUtil;
using System.Data;

namespace SHStaticRank2.Data
{
    class CalcMutilSemeSubjectRank
    {
        static DataTable _table = new DataTable();

        // 計算學生科目數
        static Dictionary<string, int> _studSubjCountDict = new Dictionary<string, int>();
        static Dictionary<string, string> _studNameDict = new Dictionary<string, string>();
        static Dictionary<string, Aspose.Words.Document> _WordDocDict = new Dictionary<string, Aspose.Words.Document>();
        static Dictionary<string, Aspose.Cells.Workbook> _ExcelCellDict = new Dictionary<string, Aspose.Cells.Workbook>();
        static List<memoText> _memoText= new List<memoText> ();
        static List<string> _PPSubjNameList = new List<string> ();
        // 用來紀錄那個ListView有沒有勾選"部訂必修專業科目"或"部訂必修實習科目"
        // Key: ListView's name, Value: ListViewItem's text
        static Dictionary<_ListViewName, string> _SpecialListViewItem = new Dictionary<_ListViewName, string>();
        enum _ListViewName {lvwSubjectPri, lvwSubjectOrd1, lvwSubjectOrd2};

        static string FolderName = "";

        public static void Setup()
        {
            FISCA.Permission.Catalog cat = FISCA.Permission.RoleAclSource.Instance["教務作業"]["功能按鈕"];
            cat.Add(new FISCA.Permission.RibbonFeature("SHSchool.SHStaticRank2.Data","計算固定排名(測試版)"));

            var button = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["計算固定排名(測試版)"]["計算多學期成績固定排名"];
            button.Enable = FISCA.Permission.UserAcl.Current["SHSchool.SHStaticRank2.Data"].Executable;//MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["學期成績處理"].Enable = CurrentUser.Acl["Button0670"].Executable;
            button.Click += delegate
            {
                var conf = new CalcMutilSemeSubjectStatucRank();
                
                if (conf.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {                    
                    //Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
                    //wb.ConvertNumericData = false;
                    //Aspose.Words.Document doc= new Aspose.Words.Document ();
                    //wb.Worksheets.Clear();
                    AccessHelper accessHelper = new AccessHelper();
                    UpdateHelper updateHelper = new UpdateHelper();
                    var setting = conf.Setting;
                    Dictionary<string, List<StudentRecord>> gradeyearStudents = new Dictionary<string, List<StudentRecord>>();
                    System.ComponentModel.BackgroundWorker bkw = new System.ComponentModel.BackgroundWorker();
                    bkw.WorkerReportsProgress = true;
                    bkw.ProgressChanged += delegate(object sender, System.ComponentModel.ProgressChangedEventArgs e)
                    {
                        FISCA.Presentation.MotherForm.SetStatusBarMessage("多學期科目成績固定排名計算中...", e.ProgressPercentage);
                    };
                    Exception exc = null;
                    bkw.RunWorkerCompleted += delegate
                    {

                        // 檢查學生科目數是否超過樣板可容納數
                        List<string> ErrorStudIDList = new List<string>();

                        foreach (KeyValuePair<string, int> data in _studSubjCountDict)
                        {
                            if (data.Value > setting.SubjectLimit)
                                ErrorStudIDList.Add(data.Key);
                        }

                        ////if (ErrorStudIDList.Count > 0)
                        ////{
                        ////    StringBuilder sb = new StringBuilder();
                        ////    sb.AppendLine("==學生科目數超出範本科目數("+setting.SubjectLimit+")清單==");
                        ////    foreach (string sid in ErrorStudIDList)
                        ////    {
                        ////        if (_studNameDict.ContainsKey(sid))
                        ////            sb.AppendLine(_studNameDict[sid]);
                        ////    }

                        ////    FISCA.Presentation.Controls.MsgBox.Show(sb.ToString());
                        ////}

                        if (ErrorStudIDList.Count > 0)
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("學生科目數超出範本科目數，清單放在工作表:[學生科目數超出範本科目數清單]");


                            //var ErrroSheet = wb.Worksheets[wb.Worksheets.Add()];
                            //ErrroSheet.Name = "學生科目數超出範本科目數清單";


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

                        


                        #region 儲存檔案
//                        string FolderName = "多學期科目成績固定排名" + DateTime.Now.Year + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day) + string.Format("{0:00}", DateTime.Now.Hour) + string.Format("{0:00}", DateTime.Now.Minute);
                        //#region Save Excl
                        
                        //foreach(string key in _ExcelCellDict.Keys)
                        //{
                        //    Aspose.Cells.Workbook data=_ExcelCellDict[key];

                        //        string inputReportName = "多學期科目成績固定排名";
                        //        string reportName = "E_" + key + inputReportName;


                        //        string path = Path.Combine(System.Windows.Forms.Application.StartupPath + "\\Reports", FolderName);
                        //        if (!Directory.Exists(path))
                        //            Directory.CreateDirectory(path);
                        //        path = Path.Combine(path, reportName + ".xls");

                        //        if (File.Exists(path))
                        //        {
                        //            int i = 1;
                        //            while (true)
                        //            {
                        //                string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                        //                if (!File.Exists(newPath))
                        //                {
                        //                    path = newPath;
                        //                    break;
                        //                }
                        //            }
                        //        }

                        //        try
                        //        {
                        //            foreach (memoText mt in _memoText)
                        //            {
                        //                if (mt.WorkSheetName == key)
                        //                    mt.FileAddress = path;
                        //            }

                        //            data.Save(path, Aspose.Cells.FileFormatType.Excel2003);


                        //            // System.Diagnostics.Process.Start(path);
                        //        }
                        //        catch (OutOfMemoryException exo1)
                        //        {
                        //            FISCA.Presentation.Controls.MsgBox.Show("記憶體不足無法產生 Excel檔案.");
                        //            return;
                        //        }
                        //        catch (Exception ex1)
                        //        {
                        //            FISCA.Presentation.Controls.MsgBox.Show(ex1.Message);

                        //            System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                        //            sd.Title = "另存新檔";
                        //            sd.FileName = reportName + ".xls";
                        //            sd.Filter = "Excel檔案 (*.xls)|*.xls|所有檔案 (*.*)|*.*";
                        //            if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        //            {
                        //                try
                        //                {
                        //                    data.Save(path, Aspose.Cells.FileFormatType.Excel2003);

                        //                }
                        //                catch (Exception exx)
                        //                {
                        //                    FISCA.Presentation.Controls.MsgBox.Show(exx.Message + ",指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        //                    return;
                        //                }
                        //            }

                        //        }
                        //        foreach (Aspose.Cells.Worksheet item in data.Worksheets)
                        //        {
                        //            item.Cells.Clear();
                        //        }
                        //        data.Worksheets.Clear();
                        //        _ExcelCellDict.Remove(key);
                        //        //_ExcelCellDict[key] = null;
                        //        GC.Collect();
                            
                        //}
                        //_ExcelCellDict.Clear();
                        //#endregion
                        //GC.Collect();


                        //#region Excel檔案索引
                        //if (_memoText.Count > 0)
                        //{
                        //    Aspose.Cells.Workbook wbm = new Aspose.Cells.Workbook();
                        //    wbm.Worksheets[0].Cells[0, 0].PutValue("名稱");
                        //    wbm.Worksheets[0].Cells[0, 1].PutValue("連結");

                        //    int idx = 1;

                        //    foreach (memoText mt in _memoText)
                        //    {
                        //        if (!string.IsNullOrEmpty(mt.Memo))
                        //        {
                        //            wbm.Worksheets[0].Cells[idx, 0].PutValue(mt.Memo);
                        //            wbm.Worksheets[0].Hyperlinks.Add(idx, 1, idx, 1, mt.FileAddress);
                        //            idx++;
                        //        }
                        //    }
                        //    wbm.Worksheets[0].AutoFitColumns();
                        //    string pathaa = Path.Combine(System.Windows.Forms.Application.StartupPath + "\\Reports", FolderName);
                        //    wbm.Save(pathaa + "\\E_多學期科目成績固定排名_索引.xls", Aspose.Cells.FileFormatType.Excel2003);
                        //} 
                        //#endregion

                        //#region Save Word
                        
                        //foreach(string key in _WordDocDict.Keys)
                        //{
                        //    Aspose.Words.Document data =_WordDocDict[key];

                        
                        //    string reportNameW = "W_" + key + "-多學期科目成績固定排名成績單";
                        //    string pathW = Path.Combine(System.Windows.Forms.Application.StartupPath + "\\Reports", FolderName);
                        //    if (!Directory.Exists(pathW))
                        //        Directory.CreateDirectory(pathW);
                        //    pathW = Path.Combine(pathW, reportNameW + ".doc");

                        //    if (File.Exists(pathW))
                        //    {
                        //        int i = 1;
                        //        while (true)
                        //        {
                        //            string newPathW = Path.GetDirectoryName(pathW) + "\\" + Path.GetFileNameWithoutExtension(pathW) + (i++) + Path.GetExtension(pathW);
                        //            if (!File.Exists(newPathW))
                        //            {
                        //                pathW = newPathW;
                        //                break;
                        //            }
                        //        }
                        //    }

                        //    Aspose.Words.Document doc = data;

                        //    try
                        //    {
                        //        doc.Save(pathW, Aspose.Words.SaveFormat.Doc);
                        //        //  System.Diagnostics.Process.Start(pathW);


                        //    }
                        //    catch (OutOfMemoryException exow)
                        //    {
                        //        FISCA.Presentation.Controls.MsgBox.Show(exow + ",記憶體不足，無法產生Word檔案");
                        //        return;
                        //    }
                        //    catch
                        //    {

                        //        System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                        //        sd.Title = "另存新檔";
                        //        sd.FileName = reportNameW + ".doc";
                        //        sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                        //        if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                        //        {
                        //            try
                        //            {
                        //                doc.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                        //            }
                        //            catch
                        //            {
                        //                FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        //                return;
                        //            }
                        //        }
                        //    }
                        //    doc.Document.Remove();
                        //    _WordDocDict.Remove(key);
                        //    //_WordDocDict[key] = null;
                        //    GC.Collect();
                        //}
                        //_WordDocDict.Clear();
                        //#endregion

                        #endregion
                        
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

                        System.Diagnostics.Process.Start(System.Windows.Forms.Application.StartupPath + "\\Reports\\"+ FolderName);
                        
                    };
                    bkw.DoWork += delegate
                    {
                        try
                        {

                            FolderName = "多學期科目成績固定排名" + DateTime.Now.Year + string.Format("{0:00}", DateTime.Now.Month) + string.Format("{0:00}", DateTime.Now.Day) + string.Format("{0:00}", DateTime.Now.Hour) + string.Format("{0:00}", DateTime.Now.Minute);

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
                            List<string> ssList = new List<string>();
                            int scc = 1;

                            // 假如有勾選"部訂必修專業科目"或"部訂必修實習科目", 需要更動勾選的科目
                            if (CheckListViewItem(setting) == true)
                            {
                                List<StudentRecord> studentRankList = new List<StudentRecord>();// 排除不排名學生
                                List<string> studentTag1List = new List<string>();   // 類別1學生
                                List<string> studentTag2List = new List<string>();   // 類別2學生

                                // 取得需要的學生
                                FilterStudents(gradeyearStudents, setting, studentRankList, studentTag1List, studentTag2List);

                                // 取得學生學期科目成績
                                accessHelper.StudentHelper.FillSemesterSubjectScore(true, studentRankList);

                                // 更動勾選的科目
                                ReNewSelectedSubject(setting, studentRankList, studentTag1List, studentTag2List);

                                if (FISCA.RTContext.IsDiagMode)
                                {
                                    Console.WriteLine("列印科目 :" + string.Join(",", setting.useSubjectPrintList));
                                    Console.WriteLine("類別1科目:" + string.Join(",", setting.useSubjecOrder1List));
                                    Console.WriteLine("類別2科目:" + string.Join(",", setting.useSubjecOrder2List));
                                    Console.WriteLine("學生數量 :" + studentRankList.Count);
                                    Console.WriteLine("類別1學生:" + studentTag1List.Count);
                                    Console.WriteLine("類別2學生:" + studentTag2List.Count);
                                    Console.WriteLine("類別1學生:" + string.Join(",", studentTag1List));
                                    Console.WriteLine("類別2學生:" + string.Join(",", studentTag2List));
                                }
                            }

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


                            bkw.ReportProgress(20);
                            string errName = "學生科目數超出範本科目數清單";
                            if (!_ExcelCellDict.ContainsKey(errName))
                                _ExcelCellDict.Add(errName, new Aspose.Cells.Workbook());

                            _ExcelCellDict[errName].Worksheets[0].Name = errName;
                            //var ErrroSheet = wb.Worksheets[wb.Worksheets.Add()];
                            //ErrroSheet.Name = "學生科目數超出範本科目數清單";
                            Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();

                            #endregion
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

                                // 學生類別一、類別二名稱
                                Dictionary<string, string> cat1Dict = new Dictionary<string, string>();
                                Dictionary<string, string> cat2Dict = new Dictionary<string, string>();
                                var studentList = gradeyearStudents[gradeyear];
                                #region 分析學生所屬類別
                                List<string> cat1List = new List<string>(), cat2List = new List<string>();
                                //不排名的學生移出list
                                //List<StudentRecord> notRankList = new List<StudentRecord>();
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
                                            //if (setting.NotRankTag != "" && setting.NotRankTag == tag.Name)
                                            //{
                                            //    notRankList.Add(studentRec);
                                            //    break;
                                            //}
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
                                            //if (setting.NotRankTag != "" && setting.NotRankTag == "[" + tag.Name + "]")
                                            //{
                                            //    notRankList.Add(studentRec);
                                            //    break;
                                            //}
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
                                }
                                //不排名的學生直接移出list
                                //foreach (var r in notRankList)
                                //{
                                //    studentList.Remove(r);
                                //}

                                // 比對學生類別與類別一、二
                                foreach (StudentRecord studRec in studentList)
                                {
                                    string tag1 = "", tag2 = "";
                                    if (studRec.Fields.ContainsKey("tag1"))
                                        tag1 = studRec.Fields["tag1"].ToString();

                                    if (studRec.Fields.ContainsKey("tag2"))
                                        tag2 = studRec.Fields["tag2"].ToString();

                                    if (cat1List.Contains(tag1))
                                        cat1Dict[studRec.StudentID] = tag1;

                                    if (cat2List.Contains(tag2))
                                        cat2Dict[studRec.StudentID] = tag2;
                                }

                                #endregion
                                //取得學生學期科目成績
                                //accessHelper.StudentHelper.FillSemesterSubjectScore(true, studentList);
                                if (setting.計算學業成績排名)
                                {
                                    accessHelper.StudentHelper.FillSemesterEntryScore(true, studentList);
                                }
                                //bkw.ReportProgress(1 + 99 * yearCount / gradeyearStudents.Count / 2);
                                Dictionary<string, List<decimal>> ranks = new Dictionary<string, List<decimal>>();
                                Dictionary<string, List<string>> rankStudents = new Dictionary<string, List<string>>();
                                //Dictionary<SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo, decimal?> selectScore = new Dictionary<SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo, decimal?>();

                                // 儲存成績用
                                Dictionary<string, studScore> selectScore = new Dictionary<string, studScore>();
                                _studNameDict.Clear();
                                _studSubjCountDict.Clear();
                                foreach (var studentRec in studentList)
                                {
                                    string studentID = studentRec.StudentID;
                                    string name = "學號：" + studentRec.StudentNumber + ",班級:" + studentRec.RefClass.ClassName + ",座號：" + studentRec.SeatNo + ",姓名：" + studentRec.StudentName;
                                    if (!_studNameDict.ContainsKey(studentID))
                                        _studNameDict.Add(studentID, name);

                                    // 處理勾選科目
                                    #region 處理勾選科目
                                    foreach (string SubjName in setting.useSubjectPrintList)
                                    {
                                        // 總分,加權總分,平均,加權平均
                                        string subjKey = studentID + "^^^" + SubjName;
                                        bool chkHasSubjectName = false;
                                        bool chkHasTag1 = false, chkHasTag2 = false;

                                        #region 處理學期科目成績
                                        foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                        {

                                            // 當成績不在勾選年級學期，科目名稱不符跳過
                                            string gs = subjectScore.GradeYear + "" + subjectScore.Semester;
                                            if (!setting.useGradeSemesterList.Contains(gs))
                                                continue;
                                            // 判斷此科目是否為需要的
                                            if (IsNeededSubject(subjectScore, _ListViewName.lvwSubjectPri, SubjName))
                                            {
                                                // 計算學生科目數
                                                if (!_studSubjCountDict.ContainsKey(studentID))
                                                    _studSubjCountDict.Add(studentID, 0);

                                                _studSubjCountDict[studentID]++;

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
                                                    selectScore[subjKey].gsCredit11 = subjectScore.Credit;
                                                }
                                                if (subjectScore.GradeYear == 1 && subjectScore.Semester == 2)
                                                {
                                                    selectScore[subjKey].gsScore12 = score;
                                                    selectScore[subjKey].gsSchoolYear12 = subjectScore.SchoolYear;
                                                    selectScore[subjKey].gsCredit12 = subjectScore.Credit;
                                                }
                                                if (subjectScore.GradeYear == 2 && subjectScore.Semester == 1)
                                                {
                                                    selectScore[subjKey].gsScore21 = score;
                                                    selectScore[subjKey].gsSchoolYear21 = subjectScore.SchoolYear;
                                                    selectScore[subjKey].gsCredit21 = subjectScore.Credit;
                                                }

                                                if (subjectScore.GradeYear == 2 && subjectScore.Semester == 2)
                                                {
                                                    selectScore[subjKey].gsScore22 = score;
                                                    selectScore[subjKey].gsSchoolYear22 = subjectScore.SchoolYear;
                                                    selectScore[subjKey].gsCredit22 = subjectScore.Credit;
                                                }
                                                if (subjectScore.GradeYear == 3 && subjectScore.Semester == 1)
                                                {
                                                    selectScore[subjKey].gsScore31 = score;
                                                    selectScore[subjKey].gsSchoolYear31 = subjectScore.SchoolYear;
                                                    selectScore[subjKey].gsCredit31 = subjectScore.Credit;
                                                }
                                                if (subjectScore.GradeYear == 3 && subjectScore.Semester == 2)
                                                {
                                                    selectScore[subjKey].gsScore32 = score;
                                                    selectScore[subjKey].gsSchoolYear32 = subjectScore.SchoolYear;
                                                    selectScore[subjKey].gsCredit32 = subjectScore.Credit;
                                                }

                                                if (subjectScore.GradeYear == 4 && subjectScore.Semester == 1)
                                                {
                                                    selectScore[subjKey].gsScore41 = score;
                                                    selectScore[subjKey].gsSchoolYear41 = subjectScore.SchoolYear;
                                                    selectScore[subjKey].gsCredit41 = subjectScore.Credit;
                                                }
                                                if (subjectScore.GradeYear == 4 && subjectScore.Semester == 2)
                                                {
                                                    selectScore[subjKey].gsScore42 = score;
                                                    selectScore[subjKey].gsSchoolYear42 = subjectScore.SchoolYear;
                                                    selectScore[subjKey].gsCredit42 = subjectScore.Credit;
                                                }
                                                // 總分
                                                selectScore[subjKey].sumScore += score;
                                                // 總分加權
                                                selectScore[subjKey].sumScoreA += (score * subjectScore.Credit);
                                                // 筆數
                                                selectScore[subjKey].subjCount++;
                                                // 學分加總
                                                selectScore[subjKey].sumCredit += subjectScore.Credit;

                                                // 類別一處理, 判斷此科目是否為類別1需要的
                                                if ((setting.useSubjecOrder1List.Contains(SubjName)) &&
                                                    IsNeededSubject(subjectScore, _ListViewName.lvwSubjectOrd1, SubjName))
                                                {
                                                    chkHasTag1 = true;
                                                    // 總分
                                                    selectScore[subjKey].sumScoreC1 += score;
                                                    // 總分加權
                                                    selectScore[subjKey].sumScoreAC1 += (score * subjectScore.Credit);
                                                    // 筆數
                                                    selectScore[subjKey].subjCountC1++;
                                                    // 學分加總
                                                    selectScore[subjKey].sumCreditC1 += subjectScore.Credit;
                                                }

                                                // 類別二處理, 判斷此科目是否為類別1需要的
                                                if ((setting.useSubjecOrder2List.Contains(SubjName)) &&
                                                    IsNeededSubject(subjectScore, _ListViewName.lvwSubjectOrd2, SubjName))
                                                {
                                                    chkHasTag2 = true;
                                                    // 總分
                                                    selectScore[subjKey].sumScoreC2 += score;
                                                    // 總分加權
                                                    selectScore[subjKey].sumScoreAC2 += (score * subjectScore.Credit);
                                                    // 筆數
                                                    selectScore[subjKey].subjCountC2++;
                                                    // 學分加總
                                                    selectScore[subjKey].sumCreditC2 += subjectScore.Credit;
                                                }

                                            }
                                        }
                                        #endregion 處理學期科目成績

                                        if (chkHasSubjectName)
                                            if (selectScore.ContainsKey(subjKey))
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

                                    }   // 科目                                
                                    #endregion 處理勾選科目

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
                                        #endregion 處理實習科目

                                    }   // 學業

                                    bkw.ReportProgress(50);
                                    // 總計成績
                                    // 處理勾選科目
                                    string subjKeyAll = "";
                                    bool chkHasSubjecName = false;
                                    #region 處理總計成績
                                    foreach (string SubjName in setting.useSubjectPrintList)
                                    {
                                        // 總分,加權總分,平均,加權平均

                                        foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                        {
                                            // 當成績不在勾選年級學期，科目名稱不符跳過
                                            string gs = subjectScore.GradeYear + "" + subjectScore.Semester;
                                            if (!setting.useGradeSemesterList.Contains(gs))
                                                continue;
                                            // 判斷是否為需要的科目
                                            if (IsNeededSubject(subjectScore, _ListViewName.lvwSubjectPri, SubjName))
                                            {
                                                chkHasSubjecName = true;
                                                subjKeyAll = studentID + "總計成績";

                                                if (!selectScore.ContainsKey(subjKeyAll))
                                                    selectScore.Add(subjKeyAll, new studScore());

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
                                                    selectScore[subjKeyAll].gsScore11 = score;
                                                    selectScore[subjKeyAll].gsSchoolYear11 = subjectScore.SchoolYear;
                                                    selectScore[subjKeyAll].gsCredit11 = subjectScore.Credit;
                                                }
                                                if (subjectScore.GradeYear == 1 && subjectScore.Semester == 2)
                                                {
                                                    selectScore[subjKeyAll].gsScore12 = score;
                                                    selectScore[subjKeyAll].gsSchoolYear12 = subjectScore.SchoolYear;
                                                    selectScore[subjKeyAll].gsCredit12 = subjectScore.Credit;
                                                }
                                                if (subjectScore.GradeYear == 2 && subjectScore.Semester == 1)
                                                {
                                                    selectScore[subjKeyAll].gsScore21 = score;
                                                    selectScore[subjKeyAll].gsSchoolYear21 = subjectScore.SchoolYear;
                                                    selectScore[subjKeyAll].gsCredit21 = subjectScore.Credit;
                                                }
                                                if (subjectScore.GradeYear == 2 && subjectScore.Semester == 2)
                                                {
                                                    selectScore[subjKeyAll].gsScore22 = score;
                                                    selectScore[subjKeyAll].gsSchoolYear22 = subjectScore.SchoolYear;
                                                    selectScore[subjKeyAll].gsCredit22 = subjectScore.Credit;
                                                }

                                                if (subjectScore.GradeYear == 3 && subjectScore.Semester == 1)
                                                {
                                                    selectScore[subjKeyAll].gsScore31 = score;
                                                    selectScore[subjKeyAll].gsSchoolYear31 = subjectScore.SchoolYear;
                                                    selectScore[subjKeyAll].gsCredit31 = subjectScore.Credit;
                                                }

                                                if (subjectScore.GradeYear == 3 && subjectScore.Semester == 2)
                                                {
                                                    selectScore[subjKeyAll].gsScore32 = score;
                                                    selectScore[subjKeyAll].gsSchoolYear32 = subjectScore.SchoolYear;
                                                    selectScore[subjKeyAll].gsCredit32 = subjectScore.Credit;
                                                }
                                                if (subjectScore.GradeYear == 4 && subjectScore.Semester == 1)
                                                {
                                                    selectScore[subjKeyAll].gsScore41 = score;
                                                    selectScore[subjKeyAll].gsSchoolYear41 = subjectScore.SchoolYear;
                                                    selectScore[subjKeyAll].gsCredit41 = subjectScore.Credit;
                                                }
                                                if (subjectScore.GradeYear == 4 && subjectScore.Semester == 2)
                                                {
                                                    selectScore[subjKeyAll].gsScore42 = score;
                                                    selectScore[subjKeyAll].gsSchoolYear42 = subjectScore.SchoolYear;
                                                    selectScore[subjKeyAll].gsCredit42 = subjectScore.Credit;
                                                }
                                                // 總分
                                                selectScore[subjKeyAll].sumScore += score;
                                                // 加總
                                                selectScore[subjKeyAll].sumScoreA += (score * subjectScore.Credit);
                                                // 筆數
                                                selectScore[subjKeyAll].subjCount++;
                                                // 學分加總
                                                selectScore[subjKeyAll].sumCredit += subjectScore.Credit;

                                                // 類別一處理, 判斷此科目是否為類別1需要的
                                                if ((setting.useSubjecOrder1List.Contains(SubjName)) &&
                                                    IsNeededSubject(subjectScore,_ListViewName.lvwSubjectOrd1,SubjName))
                                                {
                                                    // 總分
                                                    selectScore[subjKeyAll].sumScoreC1 += score;
                                                    // 總分加權
                                                    selectScore[subjKeyAll].sumScoreAC1 += (score * subjectScore.Credit);
                                                    // 筆數
                                                    selectScore[subjKeyAll].subjCountC1++;
                                                    // 學分加總
                                                    selectScore[subjKeyAll].sumCreditC1 += subjectScore.Credit;
                                                }

                                                // 類別二處理, 判斷此科目是否為類別2需要的
                                                if ((setting.useSubjecOrder2List.Contains(SubjName)) &&
                                                    IsNeededSubject(subjectScore, _ListViewName.lvwSubjectOrd2, SubjName))
                                                {
                                                    // 總分
                                                    selectScore[subjKeyAll].sumScoreC2 += score;
                                                    // 總分加權
                                                    selectScore[subjKeyAll].sumScoreAC2 += (score * subjectScore.Credit);
                                                    // 筆數
                                                    selectScore[subjKeyAll].subjCountC2++;
                                                    // 學分加總
                                                    selectScore[subjKeyAll].sumCreditC2 += subjectScore.Credit;
                                                }
                                            }
                                        }
                                    }
                                    // 平均
                                    if (chkHasSubjecName)
                                        if (selectScore.ContainsKey(subjKeyAll))
                                        {
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
                                    #endregion 處理總計成績

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
                                            key5 = studentID + gsS + key5_name;
                                            key7 = studentID + gsS + key7_name;

                                            if (!selectScore.ContainsKey(key5))
                                                selectScore.Add(key5, new studScore());

                                            if (!selectScore.ContainsKey(key7))
                                                selectScore.Add(key7, new studScore());

                                            foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                            {
                                                if (g == subjectScore.GradeYear && s == subjectScore.Semester)
                                                {
                                                    decimal score = 0, Sscore;
                                                    if (decimal.TryParse(subjectScore.Detail.GetAttribute("原始成績"), out Sscore))
                                                        score = Sscore;
                                                    #region 處理科目原始成績
                                                    if (selectScore.ContainsKey(key5))
                                                    {
                                                        // 總分
                                                        selectScore[key5].sumScore += score;
                                                        // 加總
                                                        selectScore[key5].sumScoreA += (score * subjectScore.Credit);
                                                        // 筆數
                                                        selectScore[key5].subjCount++;
                                                        // 學分加總
                                                        selectScore[key5].sumCredit += subjectScore.Credit;

                                                        // 類別一處理
                                                        // 總分
                                                        selectScore[key5].sumScoreC1 += score;
                                                        // 總分加權
                                                        selectScore[key5].sumScoreAC1 += (score * subjectScore.Credit);
                                                        // 筆數
                                                        selectScore[key5].subjCountC1++;
                                                        // 學分加總
                                                        selectScore[key5].sumCreditC1 += subjectScore.Credit;

                                                        // 類別二處理
                                                        // 總分
                                                        selectScore[key5].sumScoreC2 += score;
                                                        // 總分加權
                                                        selectScore[key5].sumScoreAC2 += (score * subjectScore.Credit);
                                                        // 筆數
                                                        selectScore[key5].subjCountC2++;
                                                        // 學分加總
                                                        selectScore[key5].sumCreditC2 += subjectScore.Credit;
                                                    }
                                                    #endregion 處理科目原始成績

                                                    #region 處理篩選科目原始成績
                                                    if (selectScore.ContainsKey(key7))
                                                    {
                                                        // 假如是使用者勾選的科目
                                                        if (IsNeededSubject(subjectScore, _ListViewName.lvwSubjectPri, setting.useSubjectPrintList))
                                                        {
                                                            // 總分
                                                            selectScore[key7].sumScore += score;
                                                            // 加總
                                                            selectScore[key7].sumScoreA += (score * subjectScore.Credit);
                                                            // 筆數
                                                            selectScore[key7].subjCount++;
                                                            // 學分加總
                                                            selectScore[key7].sumCredit += subjectScore.Credit;

                                                            // 類別一處理
                                                            // 總分
                                                            selectScore[key7].sumScoreC1 += score;
                                                            // 總分加權
                                                            selectScore[key7].sumScoreAC1 += (score * subjectScore.Credit);
                                                            // 筆數
                                                            selectScore[key7].subjCountC1++;
                                                            // 學分加總
                                                            selectScore[key7].sumCreditC1 += subjectScore.Credit;

                                                            // 類別二處理
                                                            // 總分
                                                            selectScore[key7].sumScoreC2 += score;
                                                            // 總分加權
                                                            selectScore[key7].sumScoreAC2 += (score * subjectScore.Credit);
                                                            // 筆數
                                                            selectScore[key7].subjCountC2++;
                                                            // 學分加總
                                                            selectScore[key7].sumCreditC2 += subjectScore.Credit;
                                                        }
                                                    }
                                                    #endregion 處理篩選科目原始成績
                                                }
                                            }

                                            #region 處理科目原始成績
                                            if (selectScore.ContainsKey(key5))
                                            {
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
                                            }



                                            if (selectScore.ContainsKey(key5))
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
                                            if (selectScore.ContainsKey(key7))
                                            {
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
                                            }

                                            if (selectScore.ContainsKey(key7))
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
                                    #region 處理科目原始成績加權平均平均
                                    string key6_name = "科目原始成績加權平均平均";
                                    string key8_name = "篩選科目原始成績加權平均平均";

                                    string key6 = studentID + key6_name;
                                    string key8 = studentID + key8_name;

                                    if (!selectScore.ContainsKey(key6))
                                        selectScore.Add(key6, new studScore());

                                    if (!selectScore.ContainsKey(key8))
                                        selectScore.Add(key8, new studScore());

                                    foreach (var subjectScore in studentRec.SemesterSubjectScoreList)
                                    {
                                        decimal score = 0, Sscore;
                                        if (decimal.TryParse(subjectScore.Detail.GetAttribute("原始成績"), out Sscore))
                                            score = Sscore;

                                        #region 處理科目原始成績
                                        if (selectScore.ContainsKey(key6))
                                        {
                                            // 總分
                                            selectScore[key6].sumScore += score;
                                            // 加總
                                            selectScore[key6].sumScoreA += (score * subjectScore.Credit);
                                            // 筆數
                                            selectScore[key6].subjCount++;
                                            // 學分加總
                                            selectScore[key6].sumCredit += subjectScore.Credit;

                                            // 類別一處理
                                            // 總分
                                            selectScore[key6].sumScoreC1 += score;
                                            // 總分加權
                                            selectScore[key6].sumScoreAC1 += (score * subjectScore.Credit);
                                            // 筆數
                                            selectScore[key6].subjCountC1++;
                                            // 學分加總
                                            selectScore[key6].sumCreditC1 += subjectScore.Credit;

                                            // 類別二處理
                                            // 總分
                                            selectScore[key6].sumScoreC2 += score;
                                            // 總分加權
                                            selectScore[key6].sumScoreAC2 += (score * subjectScore.Credit);
                                            // 筆數
                                            selectScore[key6].subjCountC2++;
                                            // 學分加總
                                            selectScore[key6].sumCreditC2 += subjectScore.Credit;
                                        }
                                        #endregion 處理科目原始成績

                                        #region 處理篩選科目原始成績
                                        if (selectScore.ContainsKey(key8))
                                        {
                                            if (IsNeededSubject(subjectScore, _ListViewName.lvwSubjectPri, setting.useSubjectPrintList))
                                            {
                                                // 總分
                                                selectScore[key8].sumScore += score;
                                                // 加總
                                                selectScore[key8].sumScoreA += (score * subjectScore.Credit);
                                                // 筆數
                                                selectScore[key8].subjCount++;
                                                // 學分加總
                                                selectScore[key8].sumCredit += subjectScore.Credit;

                                                // 類別一處理
                                                // 總分
                                                selectScore[key8].sumScoreC1 += score;
                                                // 總分加權
                                                selectScore[key8].sumScoreAC1 += (score * subjectScore.Credit);
                                                // 筆數
                                                selectScore[key8].subjCountC1++;
                                                // 學分加總
                                                selectScore[key8].sumCreditC1 += subjectScore.Credit;

                                                // 類別二處理
                                                // 總分
                                                selectScore[key8].sumScoreC2 += score;
                                                // 總分加權
                                                selectScore[key8].sumScoreAC2 += (score * subjectScore.Credit);
                                                // 筆數
                                                selectScore[key8].subjCountC2++;
                                                // 學分加總
                                                selectScore[key8].sumCreditC2 += subjectScore.Credit;
                                            }
                                        }
                                        #endregion 處理篩選科目原始成績
                                    }

                                    #region 處理科目原始成績
                                    if (selectScore.ContainsKey(key6))
                                    {
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
                                            selectScore[key6].avgScoreAC2 = (decimal)(selectScore[key6].sumScoreAC2 / selectScore[key6].sumCreditC2);
                                    }



                                    if (selectScore.ContainsKey(key6))
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
                                    if (selectScore.ContainsKey(key8))
                                    {
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
                                    }

                                    if (selectScore.ContainsKey(key8))
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

                                            // 平均
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 21].PutValue(selectScore[id].avgScore);
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

                                            // 加權總分
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 27].PutValue(selectScore[id].sumScoreA);
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

                                            // 加權平均
                                            _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 33].PutValue(selectScore[id].avgScoreA);

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
                                                    if(ranks[key4].IndexOf(selectScore[id].avgScoreAC2)>=0)
                                                        _ExcelCellDict[shtName].Worksheets[shtName].Cells[rowIdx, BeginColumn + 38].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreAC2) + 1);
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

                                            // 平均
                                            _ExcelCellDict[shtName2].Worksheets[shtName2].Cells[rowIdx2, 20].PutValue(selectScore[id].avgScore);

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

                                            // 平均
                                            _ExcelCellDict[shtName2_5].Worksheets[shtName2_5].Cells[rowIdx2_5, 20].PutValue(selectScore[id].avgScore);

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

                                            // 平均
                                            _ExcelCellDict[shtName2_1].Worksheets[shtName2_1].Cells[rowIdx2_1, 20].PutValue(selectScore[id].avgScore);

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

                                            // 平均
                                            _ExcelCellDict[shtName2_2].Worksheets[shtName2_2].Cells[rowIdx2_2, 20].PutValue(selectScore[id].avgScore);

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

                                            // 平均
                                            _ExcelCellDict[shtName2_3].Worksheets[shtName2_3].Cells[rowIdx2_3, 20].PutValue(selectScore[id].avgScore);

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

                                            // 平均
                                            _ExcelCellDict[shtName2_4].Worksheets[shtName2_4].Cells[rowIdx2_4, 20].PutValue(selectScore[id].avgScore);

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
                                        // 總計總分類別1
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 6].PutValue(selectScore[id].sumScoreC1);
                                        string key1 = "";
                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key1 = "總計總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 7].PutValue(ranks[key1].IndexOf(selectScore[id].sumScoreC1) + 1);
                                        }

                                        // 總計總分類別2
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 14].PutValue(selectScore[id].sumScoreC2);
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key1 = "總計總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            if (ranks.ContainsKey(key1))
                                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 15].PutValue(ranks[key1].IndexOf(selectScore[id].sumScoreC1) + 1);
                                        }

                                        // 總分
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 22].PutValue(selectScore[id].sumScore);

                                        key1 = "總計總分班排名" + studRec.RefClass.ClassID;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 23].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "總計總分科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 24].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        key1 = "總計總分全校排名" + gradeyear;
                                        if (ranks.ContainsKey(key1))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 25].PutValue(ranks[key1].IndexOf(selectScore[id].sumScore) + 1);

                                        // 平均
                                        string key3 = "";
                                        // 平均(類別一)
                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key3 = "總計平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 8].PutValue(selectScore[id].avgScoreC1);
                                            if (ranks.ContainsKey(key3))
                                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 9].PutValue(ranks[key3].IndexOf(selectScore[id].avgScoreC1) + 1);

                                        }
                                        // 平均(類別二)
                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key3 = "總計平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 16].PutValue(selectScore[id].avgScoreC2);
                                            if (ranks.ContainsKey(key3))
                                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 17].PutValue(ranks[key3].IndexOf(selectScore[id].avgScoreC2) + 1);
                                        }
                                        // 平均
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 26].PutValue(selectScore[id].avgScore);

                                        key3 = "總計平均班排名" + studRec.RefClass.ClassID;
                                        if (ranks.ContainsKey(key3))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 27].PutValue(ranks[key3].IndexOf(selectScore[id].avgScore) + 1);

                                        key3 = "總計平均科排名" + studRec.Department + "^^^" + gradeyear;

                                        if (ranks.ContainsKey(key3))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 28].PutValue(ranks[key3].IndexOf(selectScore[id].avgScore) + 1);

                                        key3 = "總計平均全校排名" + gradeyear;
                                        if (ranks.ContainsKey(key3))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 29].PutValue(ranks[key3].IndexOf(selectScore[id].avgScore) + 1);


                                        string key2 = "";
                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key2 = "總計加權總分類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            // 加權總分(類別一)
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 10].PutValue(selectScore[id].sumScoreAC1);
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 11].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreAC1) + 1);
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key2 = "總計加權總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            // 加權總分(類別二)
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 18].PutValue(selectScore[id].sumScoreAC2);
                                            if (ranks.ContainsKey(key2))
                                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 19].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreAC2) + 1);
                                        }
                                        // 加權總分
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 30].PutValue(selectScore[id].sumScoreA);
                                        key2 = "總計加權總分班排名" + studRec.RefClass.ClassID;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 31].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1);

                                        key2 = "總計加權總分科排名" + studRec.Department + "^^^" + gradeyear;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 32].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1);

                                        key2 = "總計加權總分全校排名" + gradeyear;
                                        if (ranks.ContainsKey(key2))
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 33].PutValue(ranks[key2].IndexOf(selectScore[id].sumScoreA) + 1);


                                        string key4 = "";

                                        if (studRec.Fields.ContainsKey("tag1"))
                                        {
                                            key4 = "總計加權平均類別1排名" + studRec.Fields["tag1"] + "^^^" + gradeyear;
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 12].PutValue(selectScore[id].avgScoreAC1);
                                            if (ranks.ContainsKey(key4))
                                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 13].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreAC1) + 1);
                                        }

                                        if (studRec.Fields.ContainsKey("tag2"))
                                        {
                                            key4 = "總計加權平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                            _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 20].PutValue(selectScore[id].avgScoreAC2);
                                            if (ranks.ContainsKey(key4))
                                                _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 21].PutValue(ranks[key4].IndexOf(selectScore[id].avgScoreAC2) + 1);
                                        }

                                        // 加權平均
                                        _ExcelCellDict[shtName3].Worksheets[shtName3].Cells[rowIdx3, 34].PutValue(selectScore[id].avgScoreA);
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

                                    rowIdx3++;
                                }
                                #endregion 處理總計排名並填值

                                // Excel 存檔
                                #region Save Excl
                                foreach (string key in _ExcelCellDict.Keys)
                                {
                                    Aspose.Cells.Workbook data = _ExcelCellDict[key];

                                    string inputReportName = "多學期科目成績固定排名";
                                    string reportName = "E_" + key + inputReportName;


                                    string path = Path.Combine(System.Windows.Forms.Application.StartupPath + "\\Reports", FolderName);
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

                                        data.Save(path, Aspose.Cells.FileFormatType.Excel2003);
                                     
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

                                    int idx = 1;

                                    foreach (memoText mt in _memoText)
                                    {
                                        if (!string.IsNullOrEmpty(mt.Memo))
                                        {
                                            wbm.Worksheets[0].Cells[idx, 0].PutValue(mt.Memo);
                                            wbm.Worksheets[0].Hyperlinks.Add(idx, 1, idx, 1, mt.FileAddress);
                                            idx++;
                                        }
                                    }
                                    wbm.Worksheets[0].AutoFitColumns();
                                    string pathaa = Path.Combine(System.Windows.Forms.Application.StartupPath + "\\Reports", FolderName);
                                    wbm.Save(pathaa + "\\E_多學期科目成績固定排名_索引.xls", Aspose.Cells.FileFormatType.Excel2003);
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

                                    _table.Columns.Clear();
                                    _table.Clear();
                                    #region 合併欄位使用
                                    // 固定
                                    _table.Columns.Add("學校名稱");
                                    _table.Columns.Add("科別");
                                    _table.Columns.Add("班級");
                                    _table.Columns.Add("座號");
                                    _table.Columns.Add("學號");
                                    _table.Columns.Add("姓名");
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
                                        _table.Columns.Add("一下科目成績" + i);
                                        _table.Columns.Add("一下科目學分數" + i);
                                        _table.Columns.Add("二上科目成績" + i);
                                        _table.Columns.Add("二上科目學分數" + i);
                                        _table.Columns.Add("二下科目成績" + i);
                                        _table.Columns.Add("二下科目學分數" + i);
                                        _table.Columns.Add("三上科目成績" + i);
                                        _table.Columns.Add("三上科目學分數" + i);
                                        _table.Columns.Add("三下科目成績" + i);
                                        _table.Columns.Add("三下科目學分數" + i);
                                        _table.Columns.Add("四上科目成績" + i);
                                        _table.Columns.Add("四上科目學分數" + i);
                                        _table.Columns.Add("四下科目成績" + i);
                                        _table.Columns.Add("四下科目學分數" + i);
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
                                    }

                                    #endregion 合併欄位使用


                                    foreach (StudentRecord studRec in gradeyearStudents[gradeyear])
                                    {

                                        if (studRec.RefClass.ClassName == className)
                                        {

                                            DataRow row = _table.NewRow();
                                            row["學校名稱"] = SchoolName;
                                            row["班級"] = studRec.RefClass.ClassName;
                                            row["座號"] = studRec.SeatNo;
                                            row["學號"] = studRec.StudentNumber;
                                            row["姓名"] = studRec.StudentName;
                                            row["科別"] = studRec.Department;
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
                                                        row["一上科目成績" + subjIndex] = Utility.ParseD2(selectScore[id].gsScore11.Value);
                                                    if (selectScore[id].gsCredit11.HasValue)
                                                        row["一上科目學分數" + subjIndex] = Utility.ParseD2(selectScore[id].gsCredit11.Value);
                                                    if (selectScore[id].gsScore12.HasValue)
                                                        row["一下科目成績" + subjIndex] = Utility.ParseD2(selectScore[id].gsScore12.Value);

                                                    if (selectScore[id].gsCredit12.HasValue)
                                                        row["一下科目學分數" + subjIndex] = Utility.ParseD2(selectScore[id].gsCredit12.Value);
                                                    if (selectScore[id].gsScore21.HasValue)
                                                        row["二上科目成績" + subjIndex] = Utility.ParseD2(selectScore[id].gsScore21.Value);

                                                    if (selectScore[id].gsCredit21.HasValue)
                                                        row["二上科目學分數" + subjIndex] = Utility.ParseD2(selectScore[id].gsCredit21.Value);

                                                    if (selectScore[id].gsScore22.HasValue)
                                                        row["二下科目成績" + subjIndex] = Utility.ParseD2(selectScore[id].gsScore22.Value);

                                                    if (selectScore[id].gsCredit22.HasValue)
                                                        row["二下科目學分數" + subjIndex] = Utility.ParseD2(selectScore[id].gsCredit22.Value);

                                                    if (selectScore[id].gsScore31.HasValue)
                                                        row["三上科目成績" + subjIndex] = Utility.ParseD2(selectScore[id].gsScore31.Value);

                                                    if (selectScore[id].gsCredit31.HasValue)
                                                        row["三上科目學分數" + subjIndex] = Utility.ParseD2(selectScore[id].gsCredit31.Value);

                                                    if (selectScore[id].gsScore32.HasValue)
                                                        row["三下科目成績" + subjIndex] = Utility.ParseD2(selectScore[id].gsScore32.Value);

                                                    if (selectScore[id].gsCredit32.HasValue)
                                                        row["三下科目學分數" + subjIndex] = Utility.ParseD2(selectScore[id].gsCredit32.Value);

                                                    if (selectScore[id].gsScore41.HasValue)
                                                        row["四上科目成績" + subjIndex] = Utility.ParseD2(selectScore[id].gsScore41.Value);

                                                    if (selectScore[id].gsCredit41.HasValue)
                                                        row["四上科目學分數" + subjIndex] = Utility.ParseD2(selectScore[id].gsCredit41.Value);

                                                    if (selectScore[id].gsScore42.HasValue)
                                                        row["四下科目成績" + subjIndex] = Utility.ParseD2(selectScore[id].gsScore42.Value);

                                                    if (selectScore[id].gsCredit42.HasValue)
                                                        row["四下科目學分數" + subjIndex] = Utility.ParseD2(selectScore[id].gsCredit42.Value);

                                                    row["科目平均" + subjIndex] = Utility.ParseD2(selectScore[id].avgScore);
                                                    row["科目總分" + subjIndex] = Utility.ParseD2(selectScore[id].sumScore);
                                                    row["科目加權平均" + subjIndex] = Utility.ParseD2(selectScore[id].avgScoreA);
                                                    row["科目加權總分" + subjIndex] = Utility.ParseD2(selectScore[id].sumScoreA);

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

                                                    subjIndex++;
                                                }
                                            }
                                            #endregion 處理科目

                                            // 處理科目年級學年度
                                            row["一年級學年度"] = string.Join("/", g1List.ToArray());
                                            row["二年級學年度"] = string.Join("/", g2List.ToArray());
                                            row["三年級學年度"] = string.Join("/", g3List.ToArray());
                                            row["四年級學年度"] = string.Join("/", g4List.ToArray());

                                            if (setting.計算學業成績排名)
                                            {
                                                // 處理學業
                                                #region 處理學業
                                                string id1 = studRec.StudentID + "學業";

                                                if (selectScore.ContainsKey(id1))
                                                {
                                                    if (selectScore[id1].gsScore11.HasValue)
                                                        row["一上學業成績"] = Utility.ParseD2(selectScore[id1].gsScore11.Value);

                                                    if (selectScore[id1].gsScore12.HasValue)
                                                        row["一下學業成績"] = Utility.ParseD2(selectScore[id1].gsScore12.Value);

                                                    if (selectScore[id1].gsScore21.HasValue)
                                                        row["二上學業成績"] = Utility.ParseD2(selectScore[id1].gsScore21.Value);

                                                    if (selectScore[id1].gsScore22.HasValue)
                                                        row["二下學業成績"] = Utility.ParseD2(selectScore[id1].gsScore22.Value);

                                                    if (selectScore[id1].gsScore31.HasValue)
                                                        row["三上學業成績"] = Utility.ParseD2(selectScore[id1].gsScore31.Value);

                                                    if (selectScore[id1].gsScore32.HasValue)
                                                        row["三下學業成績"] = Utility.ParseD2(selectScore[id1].gsScore32.Value);

                                                    if (selectScore[id1].gsScore41.HasValue)
                                                        row["四上學業成績"] = Utility.ParseD2(selectScore[id1].gsScore41.Value);

                                                    if (selectScore[id1].gsScore42.HasValue)
                                                        row["四下學業成績"] = Utility.ParseD2(selectScore[id1].gsScore42.Value);


                                                    row["學業平均"] = Utility.ParseD2(selectScore[id1].avgScore);
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
                                                    row["學業總分"] = Utility.ParseD2(selectScore[id1].sumScore);
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
                                                #endregion 處理學業

                                                // 處理學業原始
                                                #region 處理學業原始
                                                string id1_5 = studRec.StudentID + "學業原始";

                                                if (selectScore.ContainsKey(id1_5))
                                                {
                                                    if (selectScore[id1_5].gsScore11.HasValue)
                                                        row["一上學業原始成績"] = Utility.ParseD2(selectScore[id1_5].gsScore11.Value);

                                                    if (selectScore[id1_5].gsScore12.HasValue)
                                                        row["一下學業原始成績"] = Utility.ParseD2(selectScore[id1_5].gsScore12.Value);

                                                    if (selectScore[id1_5].gsScore21.HasValue)
                                                        row["二上學業原始成績"] = Utility.ParseD2(selectScore[id1_5].gsScore21.Value);

                                                    if (selectScore[id1_5].gsScore22.HasValue)
                                                        row["二下學業原始成績"] = Utility.ParseD2(selectScore[id1_5].gsScore22.Value);

                                                    if (selectScore[id1_5].gsScore31.HasValue)
                                                        row["三上學業原始成績"] = Utility.ParseD2(selectScore[id1_5].gsScore31.Value);

                                                    if (selectScore[id1_5].gsScore32.HasValue)
                                                        row["三下學業原始成績"] = Utility.ParseD2(selectScore[id1_5].gsScore32.Value);

                                                    if (selectScore[id1_5].gsScore41.HasValue)
                                                        row["四上學業原始成績"] = Utility.ParseD2(selectScore[id1_5].gsScore41.Value);

                                                    if (selectScore[id1_5].gsScore42.HasValue)
                                                        row["四下學業原始成績"] = Utility.ParseD2(selectScore[id1_5].gsScore42.Value);


                                                    row["學業原始平均"] = Utility.ParseD2(selectScore[id1_5].avgScore);
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
                                                    row["學業原始總分"] = Utility.ParseD2(selectScore[id1_5].sumScore);
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
                                                #endregion 處理學業原始

                                                // 處理學業體育
                                                #region 處理學業體育
                                                string id1_1 = studRec.StudentID + "學業體育";

                                                if (selectScore.ContainsKey(id1_1))
                                                {
                                                    if (selectScore[id1_1].gsScore11.HasValue)
                                                        row["一上學業體育成績"] = Utility.ParseD2(selectScore[id1_1].gsScore11.Value);

                                                    if (selectScore[id1_1].gsScore12.HasValue)
                                                        row["一下學業體育成績"] = Utility.ParseD2(selectScore[id1_1].gsScore12.Value);

                                                    if (selectScore[id1_1].gsScore21.HasValue)
                                                        row["二上學業體育成績"] = Utility.ParseD2(selectScore[id1_1].gsScore21.Value);

                                                    if (selectScore[id1_1].gsScore22.HasValue)
                                                        row["二下學業體育成績"] = Utility.ParseD2(selectScore[id1_1].gsScore22.Value);

                                                    if (selectScore[id1_1].gsScore31.HasValue)
                                                        row["三上學業體育成績"] = Utility.ParseD2(selectScore[id1_1].gsScore31.Value);

                                                    if (selectScore[id1_1].gsScore32.HasValue)
                                                        row["三下學業體育成績"] = Utility.ParseD2(selectScore[id1_1].gsScore32.Value);

                                                    if (selectScore[id1_1].gsScore41.HasValue)
                                                        row["四上學業體育成績"] = Utility.ParseD2(selectScore[id1_1].gsScore41.Value);

                                                    if (selectScore[id1_1].gsScore42.HasValue)
                                                        row["四下學業體育成績"] = Utility.ParseD2(selectScore[id1_1].gsScore42.Value);


                                                    row["學業體育平均"] = Utility.ParseD2(selectScore[id1_1].avgScore);
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
                                                    row["學業體育總分"] = Utility.ParseD2(selectScore[id1_1].sumScore);
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
                                                #endregion 處理學業體育

                                                // 處理學業健康與護理
                                                #region 處理學業健康與護理
                                                string id1_2 = studRec.StudentID + "學業健康與護理";

                                                if (selectScore.ContainsKey(id1_2))
                                                {
                                                    if (selectScore[id1_2].gsScore11.HasValue)
                                                        row["一上學業健康與護理成績"] = Utility.ParseD2(selectScore[id1_2].gsScore11.Value);

                                                    if (selectScore[id1_2].gsScore12.HasValue)
                                                        row["一下學業健康與護理成績"] = Utility.ParseD2(selectScore[id1_2].gsScore12.Value);

                                                    if (selectScore[id1_2].gsScore21.HasValue)
                                                        row["二上學業健康與護理成績"] = Utility.ParseD2(selectScore[id1_2].gsScore21.Value);

                                                    if (selectScore[id1_2].gsScore22.HasValue)
                                                        row["二下學業健康與護理成績"] = Utility.ParseD2(selectScore[id1_2].gsScore22.Value);

                                                    if (selectScore[id1_2].gsScore31.HasValue)
                                                        row["三上學業健康與護理成績"] = Utility.ParseD2(selectScore[id1_2].gsScore31.Value);

                                                    if (selectScore[id1_2].gsScore32.HasValue)
                                                        row["三下學業健康與護理成績"] = Utility.ParseD2(selectScore[id1_2].gsScore32.Value);

                                                    if (selectScore[id1_2].gsScore41.HasValue)
                                                        row["四上學業健康與護理成績"] = Utility.ParseD2(selectScore[id1_2].gsScore41.Value);

                                                    if (selectScore[id1_2].gsScore42.HasValue)
                                                        row["四下學業健康與護理成績"] = Utility.ParseD2(selectScore[id1_2].gsScore42.Value);


                                                    row["學業健康與護理平均"] = Utility.ParseD2(selectScore[id1_2].avgScore);
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
                                                    row["學業健康與護理總分"] = Utility.ParseD2(selectScore[id1_2].sumScore);
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
                                                #endregion 處理學業健康與護理

                                                // 處理學業國防通識
                                                #region 處理學業國防通識
                                                string id1_3 = studRec.StudentID + "學業國防通識";

                                                if (selectScore.ContainsKey(id1_3))
                                                {
                                                    if (selectScore[id1_3].gsScore11.HasValue)
                                                        row["一上學業國防通識成績"] = Utility.ParseD2(selectScore[id1_3].gsScore11.Value);

                                                    if (selectScore[id1_3].gsScore12.HasValue)
                                                        row["一下學業國防通識成績"] = Utility.ParseD2(selectScore[id1_3].gsScore12.Value);

                                                    if (selectScore[id1_3].gsScore21.HasValue)
                                                        row["二上學業國防通識成績"] = Utility.ParseD2(selectScore[id1_3].gsScore21.Value);

                                                    if (selectScore[id1_3].gsScore22.HasValue)
                                                        row["二下學業國防通識成績"] = Utility.ParseD2(selectScore[id1_3].gsScore22.Value);

                                                    if (selectScore[id1_3].gsScore31.HasValue)
                                                        row["三上學業國防通識成績"] = Utility.ParseD2(selectScore[id1_3].gsScore31.Value);

                                                    if (selectScore[id1_3].gsScore32.HasValue)
                                                        row["三下學業國防通識成績"] = Utility.ParseD2(selectScore[id1_3].gsScore32.Value);

                                                    if (selectScore[id1_3].gsScore41.HasValue)
                                                        row["四上學業國防通識成績"] = Utility.ParseD2(selectScore[id1_3].gsScore41.Value);

                                                    if (selectScore[id1_3].gsScore42.HasValue)
                                                        row["四下學業國防通識成績"] = Utility.ParseD2(selectScore[id1_3].gsScore42.Value);


                                                    row["學業國防通識平均"] = Utility.ParseD2(selectScore[id1_3].avgScore);
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
                                                    row["學業國防通識總分"] = Utility.ParseD2(selectScore[id1_3].sumScore);
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
                                                #endregion 處理學業國防通識

                                                // 處理學業實習科目
                                                #region 處理學業實習科目
                                                string id1_4 = studRec.StudentID + "學業實習科目";

                                                if (selectScore.ContainsKey(id1_4))
                                                {
                                                    if (selectScore[id1_4].gsScore11.HasValue)
                                                        row["一上學業實習科目成績"] = Utility.ParseD2(selectScore[id1_4].gsScore11.Value);

                                                    if (selectScore[id1_4].gsScore12.HasValue)
                                                        row["一下學業實習科目成績"] = Utility.ParseD2(selectScore[id1_4].gsScore12.Value);

                                                    if (selectScore[id1_4].gsScore21.HasValue)
                                                        row["二上學業實習科目成績"] = Utility.ParseD2(selectScore[id1_4].gsScore21.Value);

                                                    if (selectScore[id1_4].gsScore22.HasValue)
                                                        row["二下學業實習科目成績"] = Utility.ParseD2(selectScore[id1_4].gsScore22.Value);

                                                    if (selectScore[id1_4].gsScore31.HasValue)
                                                        row["三上學業實習科目成績"] = Utility.ParseD2(selectScore[id1_4].gsScore31.Value);

                                                    if (selectScore[id1_4].gsScore32.HasValue)
                                                        row["三下學業實習科目成績"] = Utility.ParseD2(selectScore[id1_4].gsScore32.Value);

                                                    if (selectScore[id1_4].gsScore41.HasValue)
                                                        row["四上學業實習科目成績"] = Utility.ParseD2(selectScore[id1_4].gsScore41.Value);

                                                    if (selectScore[id1_4].gsScore42.HasValue)
                                                        row["四下學業實習科目成績"] = Utility.ParseD2(selectScore[id1_4].gsScore42.Value);


                                                    row["學業實習科目平均"] = Utility.ParseD2(selectScore[id1_4].avgScore);
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
                                                    row["學業實習科目總分"] = Utility.ParseD2(selectScore[id1_4].sumScore);
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
                                                #endregion 處理學業實習科目

                                            }
                                            // 處理總計
                                            #region 處理總計
                                            string id2 = studRec.StudentID + "總計成績";
                                            if (selectScore.ContainsKey(id2))
                                            {

                                                row["總計加權平均"] = Utility.ParseD2(selectScore[id2].avgScoreA);
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
                                                        row["總計加權平均類別一"] = Utility.ParseD2(selectScore[id2].avgScoreAC1);
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

                                                row["總計加權總分"] = Utility.ParseD2(selectScore[id2].sumScoreA);

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
                                                        row["總計加權總分類別一"] = Utility.ParseD2(selectScore[id2].sumScoreAC1);
                                                        row["總計加權總分類別一排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreAC1) + 1; ;
                                                        row["總計加權總分類別一排名母數"] = ranks[key2].Count;
                                                    }
                                                }

                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    key2 = "總計加權總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key2))
                                                    {
                                                        row["總計加權總分類別二"] = Utility.ParseD2(selectScore[id2].sumScoreAC2);
                                                        row["總計加權總分類別二排名"] = ranks[key2].IndexOf(selectScore[id2].sumScoreAC2) + 1; ;
                                                        row["總計加權總分類別二排名母數"] = ranks[key2].Count;
                                                    }
                                                }

                                                row["總計平均"] = Utility.ParseD2(selectScore[id2].avgScore);
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
                                                        row["總計平均類別一"] = Utility.ParseD2(selectScore[id2].avgScoreC1);
                                                        row["總計平均類別一排名"] = ranks[key3].IndexOf(selectScore[id2].avgScoreC1) + 1;
                                                        row["總計平均類別一排名母數"] = ranks[key3].Count;
                                                    }
                                                }

                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    key3 = "總計平均類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key3))
                                                    {
                                                        row["總計平均類別二"] = Utility.ParseD2(selectScore[id2].avgScoreC2);
                                                        row["總計平均類別二排名"] = ranks[key3].IndexOf(selectScore[id2].avgScoreC2) + 1;
                                                        row["總計平均類別二排名母數"] = ranks[key3].Count;
                                                    }
                                                }

                                                row["總計總分"] = Utility.ParseD2(selectScore[id2].sumScore);
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
                                                        row["總計總分類別一"] = Utility.ParseD2(selectScore[id2].sumScoreC1);
                                                        row["總計總分類別一排名"] = ranks[key1].IndexOf(selectScore[id2].sumScoreC1) + 1;
                                                        row["總計總分類別一排名母數"] = ranks[key1].Count;
                                                    }
                                                }
                                                if (studRec.Fields.ContainsKey("tag2"))
                                                {
                                                    key1 = "總計總分類別2排名" + studRec.Fields["tag2"] + "^^^" + gradeyear;
                                                    if (ranks.ContainsKey(key1))
                                                    {
                                                        row["總計總分類別二"] = Utility.ParseD2(selectScore[id2].sumScoreC2);
                                                        row["總計總分類別二排名"] = ranks[key1].IndexOf(selectScore[id2].sumScoreC2) + 1;
                                                        row["總計總分類別二排名母數"] = ranks[key1].Count;
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

                                                        row[gsS + "科目原始成績加權平均"] = Utility.ParseD2(selectScore[id5].avgScoreA);
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
                                                                row[gsS + "科目原始成績加權平均類別一"] = Utility.ParseD2(selectScore[id5].avgScoreAC1);
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
                                                                row[gsS + "科目原始成績加權平均類別二"] = Utility.ParseD2(selectScore[id5].avgScoreAC2);
                                                                row[gsS + "科目原始成績加權平均類別二排名"] = rr;
                                                                row[gsS + "科目原始成績加權平均類別二排名母數"] = ranks[key5a].Count;
                                                                row[gsS + "科目原始成績加權平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key5a].Count);
                                                            }
                                                        }
                                                    }
                                                    #endregion 處理科目原始成績

                                                    #region 處理篩選科目原始成績
                                                    if (selectScore.ContainsKey(id7))
                                                    {
                                                        row[gsS + "篩選科目原始成績加權平均"] = Utility.ParseD2(selectScore[id7].avgScoreA);

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
                                                                row[gsS + "篩選科目原始成績加權平均類別一"] = Utility.ParseD2(selectScore[id7].avgScoreAC1);
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
                                                                row[gsS + "篩選科目原始成績加權平均類別二"] = Utility.ParseD2(selectScore[id7].avgScoreAC2);
                                                                row[gsS + "篩選科目原始成績加權平均類別二排名"] = rr;
                                                                row[gsS + "篩選科目原始成績加權平均類別二排名母數"] = ranks[key7a].Count;
                                                                row[gsS + "篩選科目原始成績加權平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key7a].Count);
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

                                                row["科目原始成績加權平均平均"] = Utility.ParseD2(selectScore[id6].avgScoreA);
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
                                                        row["科目原始成績加權平均平均類別一"] = Utility.ParseD2(selectScore[id6].avgScoreAC1);
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
                                                        row["科目原始成績加權平均平均類別二"] = Utility.ParseD2(selectScore[id6].avgScoreAC2);
                                                        row["科目原始成績加權平均平均類別二排名"] = rr;
                                                        row["科目原始成績加權平均平均類別二排名母數"] = ranks[key6a].Count;
                                                        row["科目原始成績加權平均平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key6a].Count);
                                                    }
                                                }
                                            }
                                            #endregion 處理科目原始成績

                                            #region 處理篩選科目原始成績
                                            if (selectScore.ContainsKey(id8))
                                            {

                                                row["篩選科目原始成績加權平均平均"] = Utility.ParseD2(selectScore[id8].avgScoreA);
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
                                                        row["篩選科目原始成績加權平均平均類別一"] = Utility.ParseD2(selectScore[id8].avgScoreAC1);
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
                                                        row["篩選科目原始成績加權平均平均類別二"] = Utility.ParseD2(selectScore[id8].avgScoreAC2);
                                                        row["篩選科目原始成績加權平均平均類別二排名"] = rr;
                                                        row["篩選科目原始成績加權平均平均類別二排名母數"] = ranks[key8a].Count;
                                                        row["篩選科目原始成績加權平均平均類別二排名百分比"] = Utility.ParseRankPercent(rr, ranks[key8a].Count);
                                                    }
                                                }
                                            }
                                            #endregion 處理篩選科目原始成績
                                            #endregion 處理科目原始成績加權平均平均


                                            _table.Rows.Add(row);

                                        }
                                    } // data row
                                   

                                    List<string> fields = new List<string>(docTemplate.MailMerge.GetFieldNames());
                                    List<string> rmColumns = new List<string> ();
                                    
                                    foreach(DataColumn dc in _table.Columns)
                                    {
                                        if (!fields.Contains(dc.ColumnName))
                                        {
                                            rmColumns.Add(dc.ColumnName);
                                        }
                                    }
                                    // 檢查有資料才Merge
                                    foreach (string str in rmColumns)
                                        _table.Columns.Remove(str);

                                    GC.Collect();

                                    Aspose.Words.Document document = new Aspose.Words.Document();
                                    document = docTemplate;
                                    doc.Sections.Add(doc.ImportNode(document.Sections[0], true));

                                    doc.MailMerge.MergeImageField += new Aspose.Words.Reporting.MergeImageFieldEventHandler(MailMerge_MergeImageField);
                                    doc.MailMerge.Execute(_table);
                                    doc.MailMerge.RemoveEmptyParagraphs = true;
                                    doc.MailMerge.DeleteFields();
                                    _table.Clear();
                                    
                                    //_WordDocDict.Add(className, doc);

                                    #region Word 存檔
                                    string reportNameW = "W_" + className + "-多學期科目成績固定排名成績單";
                                    string pathW = Path.Combine(System.Windows.Forms.Application.StartupPath + "\\Reports", FolderName);
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
                                        doc.Save(pathW, Aspose.Words.SaveFormat.Doc);

                                    }
                                    catch (OutOfMemoryException exow)
                                    {
                                        exc = exow;
                                    }
                                    doc = null;
                                    GC.Collect();
                                    #endregion

                                }// doc
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
            };
        }

        static void MailMerge_MergeImageField(object sender, Aspose.Words.Reporting.MergeImageFieldEventArgs e)
        {
            
        }

        /// <summary>
        /// 檢查有沒有勾選"部訂必修專業科目"或"部訂必修實習科目"
        /// </summary>
        /// <param name="setting"></param>
        /// <returns></returns>
        private static bool CheckListViewItem(Configure setting)
        {
            string specialItem1 = "部訂必修專業科目";
            string specialItem2 = "部訂必修實習科目";
            bool retValue = false;
            _SpecialListViewItem.Clear();

            // 檢查勾選科目
            if (setting.useSubjectPrintList.Contains(specialItem1))
            {
                _SpecialListViewItem.Add(_ListViewName.lvwSubjectPri, specialItem1);
                retValue = true;
            }
            else if (setting.useSubjectPrintList.Contains(specialItem2))
            {
                _SpecialListViewItem.Add(_ListViewName.lvwSubjectPri, specialItem2);
                retValue = true;
            }

            // 檢查類別一勾選科目
            if (!string.IsNullOrEmpty(setting.Rank1Tag))
            {
                if (setting.useSubjecOrder1List.Contains(specialItem1))
                {
                    _SpecialListViewItem.Add(_ListViewName.lvwSubjectOrd1, specialItem1);
                    retValue = true;
                }
                else if (setting.useSubjecOrder1List.Contains(specialItem2))
                {
                    _SpecialListViewItem.Add(_ListViewName.lvwSubjectOrd1, specialItem2);
                    retValue = true;
                }
            }

            // 檢查類別二勾選科目
            if (!string.IsNullOrEmpty(setting.Rank2Tag))
            {
                if (setting.useSubjecOrder2List.Contains(specialItem1))
                {
                    _SpecialListViewItem.Add(_ListViewName.lvwSubjectOrd2, specialItem1);
                    retValue = true;
                }
                else if (setting.useSubjecOrder2List.Contains(specialItem2))
                {
                    _SpecialListViewItem.Add(_ListViewName.lvwSubjectOrd2, specialItem2);
                    retValue = true;
                }
            }
            return retValue;
        }

        /// <summary>
        /// 排除不排名的學生, 並取得類別1 2的學生
        /// </summary>
        /// <param name="gradeyearStudents">(I)所有的學生</param>
        /// <param name="setting">(I)設定</param>
        /// <param name="studentRankList">(O)排除不排名學生</param>
        /// <param name="studentTag1List">(O)類別1學生</param>
        /// <param name="studentTag2List">(O)類別2學生</param>

        private static void FilterStudents(Dictionary<string, List<StudentRecord>> gradeyearStudents,
                                            Configure setting,
                                            List<StudentRecord> studentRankList,
                                            List<string> studentTag1List,
                                            List<string> studentTag2List)
        {
            // 取得要排名的學生
            #region 取得要排名的學生
            foreach (string gradeYear in gradeyearStudents.Keys)
            {
                List<StudentRecord> studentList = gradeyearStudents[gradeYear];
                List<StudentRecord> notRankList = new List<StudentRecord>();
                // 取得不排名的學生跟類別1,2的學生
                #region 取得不排名的學生
                foreach (StudentRecord studentRec in studentList)
                {
                    foreach (var tag in studentRec.StudentCategorys)
                    {
                        string tagName = "";
                        if (tag.SubCategory == "")
                        {
                            tagName = tag.Name;
                        }
                        else
                        {
                            tagName = "[" + tag.Name + "]";
                        }

                        if (setting.NotRankTag != "" && setting.NotRankTag == tagName)
                        {
                            notRankList.Add(studentRec);
                            break;
                        }

                        if (setting.Rank1Tag != "" && setting.Rank1Tag == tagName)
                        {
                            if (!studentTag1List.Contains(studentRec.StudentID))
                                studentTag1List.Add(studentRec.StudentID);
                        }

                        if (setting.Rank2Tag != "" && setting.Rank2Tag == tagName)
                        {
                            if (!studentTag2List.Contains(studentRec.StudentID))
                                studentTag2List.Add(studentRec.StudentID);
                        }
                    }
                }
                #endregion 取得不排名的學生

                // 去掉不排名的學生
                foreach (var r in notRankList)
                {
                    studentList.Remove(r);
                }

                studentRankList.AddRange(studentList);
            }
            #endregion 取得要排名的學生
        }

        /// <summary>
        /// 更動勾選的科目
        /// </summary>
        /// <param name="setting">(O)設定</param>
        /// <param name="studentRankList">(I)排除不排名學生</param>
        /// <param name="studentTag1List">(I)類別1學生</param>
        /// <param name="studentTag2List">(I)類別2學生</param>
        private static void ReNewSelectedSubject(Configure setting,
                                                    List<StudentRecord> studentRankList,
                                                    List<string> studentTag1List,
                                                    List<string> studentTag2List)
        {
            // 清空需要更動的科目
            if (_SpecialListViewItem.ContainsKey(_ListViewName.lvwSubjectPri))
            {
                setting.useSubjectPrintList.Clear();
            }
            if (_SpecialListViewItem.ContainsKey(_ListViewName.lvwSubjectOrd1))
            {
                setting.useSubjecOrder1List.Clear();
            }
            if (_SpecialListViewItem.ContainsKey(_ListViewName.lvwSubjectOrd2))
            {
                setting.useSubjecOrder2List.Clear();
            }

            // 更動勾選的科目
            #region 更動勾選的科目
            foreach (StudentRecord student in studentRankList)
            {
                foreach (var subjectScore in student.SemesterSubjectScoreList)
                {
                    // 當成績不在勾選年級學期跳過
                    string gs = subjectScore.GradeYear + "" + subjectScore.Semester;
                    if (!setting.useGradeSemesterList.Contains(gs))
                        continue;

                    if (IsNeededSubject(subjectScore, _ListViewName.lvwSubjectPri) == true)
                    {
                        if (!setting.useSubjectPrintList.Contains(subjectScore.Subject))
                            setting.useSubjectPrintList.Add(subjectScore.Subject);
                    }

                    // 假如是類別1的學生, 檢查這個科目是否要加入
                    if (studentTag1List.Contains(student.StudentID))
                    {
                        if (IsNeededSubject(subjectScore, _ListViewName.lvwSubjectOrd1) == true)
                        {
                            if (!setting.useSubjecOrder1List.Contains(subjectScore.Subject))
                                setting.useSubjecOrder1List.Add(subjectScore.Subject);
                        }
                    }

                    // 假如是類別2的學生, 檢查這個科目是否要加入
                    if (studentTag2List.Contains(student.StudentID))
                    {
                        if (IsNeededSubject(subjectScore, _ListViewName.lvwSubjectOrd2) == true)
                        {
                            if (!setting.useSubjecOrder2List.Contains(subjectScore.Subject))
                                setting.useSubjecOrder2List.Add(subjectScore.Subject);
                        }
                    }
                }
            }
            #endregion 更動勾選的科目
        }

        /// <summary>
        /// 假如有勾選"部訂必修專業科目"或"部訂必修實習科目", 則判斷是否為指定科目
        /// 若無勾選"部訂必修專業科目"或"部訂必修實習科目", 則回傳false
        /// </summary>
        /// <param name="subjectScore"></param>
        /// <param name="listViewName"></param>
        /// <returns></returns>
        private static bool IsNeededSubject(SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo subjectScore, _ListViewName listViewName)
        {
            if (_SpecialListViewItem.ContainsKey(listViewName))
            {
                if (subjectScore.Detail.GetAttribute("修課校部訂") == "部訂" && subjectScore.Require == true)
                {
                    if (_SpecialListViewItem[listViewName] == "部訂必修專業科目" && subjectScore.Detail.GetAttribute("開課分項類別") == "專業科目")
                    {
                        return true;
                    }
                    if (_SpecialListViewItem[listViewName] == "部訂必修實習科目" && subjectScore.Detail.GetAttribute("開課分項類別") == "實習科目")
                    {
                        return true;
                    }
                }
            }

            return false;
        }

        /// <summary>
        /// 先判斷是否為指定的科目名稱, 假如有勾選"部訂必修專業科目"或"部訂必修實習科目", 則再判斷是否為指定科目
        /// </summary>
        /// <param name="subject"></param>
        /// <param name="listViewName"></param>
        /// <param name="subjectName">指定的科目名稱</param>
        /// <returns></returns>
        private static bool IsNeededSubject(SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo subject, _ListViewName listViewName, string subjectName)
        {
            if (subject.Subject == subjectName)
            {
                if (_SpecialListViewItem.ContainsKey(listViewName))
                {
                    return IsNeededSubject(subject, listViewName);
                }
                else
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// 先判斷是否為勾選的科目名稱, 假如有勾選"部訂必修專業科目"或"部訂必修實習科目", 則再判斷是否為指定科目
        /// </summary>
        /// <param name="subjectScore"></param>
        /// <param name="listViewName"></param>
        /// <param name="subjectList"></param>
        /// <returns></returns>
        private static bool IsNeededSubject(SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo subjectScore, _ListViewName listViewName, List<string> subjectList)
        {
            foreach (string subjectName in subjectList)
            {
                if (IsNeededSubject(subjectScore, listViewName, subjectName) == true)
                {
                    return true;
                }
            }
            return false;
        }

        
    }
}
