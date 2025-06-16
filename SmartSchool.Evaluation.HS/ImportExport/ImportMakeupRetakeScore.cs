using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Xml;
using System.Xml.Linq;
using FISCA.DSAUtil;
using SmartSchool.AccessControl;
using SmartSchool.API.PlugIn;
using SmartSchool.API.PlugIn.Import;
using SmartSchool.Common;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using SmartSchool.Evaluation.Process.Wizards.LearningHistory;
using SmartSchool.ImportSupport.Validators;

namespace SmartSchool.Evaluation.ImportExport
{
    [FeatureCode("00c64573-9060-4c94-ba02-cc6508c84ca8")]
    public class ImportMakeupRetakeScore : SmartSchool.API.PlugIn.Import.Importer
    {
        public ImportMakeupRetakeScore()
        {
            this.Image = null;
            this.Text = "匯入重補修成績";
        }

        public override void InitializeImport(SmartSchool.API.PlugIn.Import.ImportWizard wizard)
        {
            Dictionary<string, int> _ID_SchoolYear_Semester_GradeYear = new Dictionary<string, int>();
            Dictionary<string, List<string>> _ID_SchoolYear_Semester_Subject = new Dictionary<string, List<string>>();
            Dictionary<string, StudentRecord> _StudentCollection = new Dictionary<string, StudentRecord>();
            Dictionary<StudentRecord, Dictionary<int, decimal>> _StudentPassScore = new Dictionary<StudentRecord, Dictionary<int, decimal>>();
            AccessHelper _AccessHelper;

            VirtualRadioButton autoCheckPass = new VirtualRadioButton("自動判斷取得學分", false);
            VirtualRadioButton manulCheckPass = new VirtualRadioButton("手動判斷取得學分", true);
            autoCheckPass.CheckedChanged += delegate
            {
                if (autoCheckPass.Checked)
                {
                    if (!wizard.RequiredFields.Contains("成績年級"))
                        wizard.RequiredFields.Add("成績年級");
                    if (!wizard.RequiredFields.Contains("取得學分"))
                        wizard.RequiredFields.Add("取得學分");
                }
            };
            manulCheckPass.CheckedChanged += delegate
            {
                if (manulCheckPass.Checked)
                {
                    if (wizard.RequiredFields.Contains("成績年級"))
                        wizard.RequiredFields.Remove("成績年級");
                    if (wizard.RequiredFields.Contains("取得學分"))
                        wizard.RequiredFields.Remove("取得學分");
                }
            };
            wizard.Options.AddRange(autoCheckPass, manulCheckPass);
            wizard.PackageLimit = 3000;
            wizard.ImportableFields.AddRange("領域", "科目", "科目級別", "學年度", "學期", "英文名稱", "學分數", "分項類別", "成績年級", "必選修", "校部訂", "原始成績", "補考成績", "重修成績", "手動調整成績", "學年調整成績", "取得學分", "不計學分", "不需評分", "註記", "是否補修成績", "補修學年度", "補修學期", "重修學年度", "重修學期", "修課及格標準", "修課補考標準", "修課備註", "修課直接指定總成績", "免修", "抵免", "指定學年科目名稱", "課程代碼", "報部科目名稱", "是否重讀");

            wizard.RequiredFields.AddRange("科目", "科目級別", "學年度", "學期", "是否補修成績");
            wizard.ValidateStart += delegate (object sender, SmartSchool.API.PlugIn.Import.ValidateStartEventArgs e)
            {
                #region ValidateStart
                _ID_SchoolYear_Semester_GradeYear.Clear();
                _ID_SchoolYear_Semester_Subject.Clear();
                _StudentCollection.Clear();
                _AccessHelper = new AccessHelper();

                List<StudentRecord> list = _AccessHelper.StudentHelper.GetStudents(e.List);
                MultiThreadWorker<StudentRecord> loader = new MultiThreadWorker<StudentRecord>();
                loader.MaxThreads = 3;
                loader.PackageSize = 250;
                loader.PackageWorker += delegate (object sender1, PackageWorkEventArgs<StudentRecord> e1)
                {
                    _AccessHelper.StudentHelper.FillSemesterSubjectScore(false, e1.List);
                };
                loader.Run(list);

                foreach (StudentRecord stu in list)
                {
                    if (!_StudentCollection.ContainsKey(stu.StudentID))
                        _StudentCollection.Add(stu.StudentID, stu);
                }
                #endregion
            };
            wizard.ValidateRow += delegate (object sender, SmartSchool.API.PlugIn.Import.ValidateRowEventArgs e)
            {
                #region ValidateRow
                int t;
                decimal k;
                decimal d;
                StudentRecord student;
                if (_StudentCollection.ContainsKey(e.Data.ID))
                {
                    student = _StudentCollection[e.Data.ID];
                }
                else
                {
                    e.ErrorMessage = "壓根就沒有這個學生" + e.Data.ID;
                    return;
                }
                bool inputFormatPass = true;
                #region 驗各欄位填寫格式
                foreach (string field in e.SelectFields)
                {
                    string value = e.Data[field];
                    switch (field)
                    {
                        default:
                            break;
                        case "領域":
                            break;
                        case "科目":
                            if (value == "")
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填寫");
                            }
                            break;
                        case "科目級別":
                            if (value != "" && !int.TryParse(value, out t))
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入空白或整數");
                            }
                            break;
                        case "學年度":
                        case "學分數":
                            if (value == "" || !decimal.TryParse(value, out k))
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入數字或小數");
                            }
                            break;
                        case "成績年級":
                            if (value == "" || !int.TryParse(value, out t))
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入整數");
                            }
                            break;
                        case "學期":
                            if (value == "" || !int.TryParse(value, out t) || t > 2 || t < 1)
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入1或2");
                            }
                            break;
                        case "原始成績":
                        case "補考成績":
                        case "重修成績":
                        case "手動調整成績":
                        case "學年調整成績":
                        case "修課及格標準":
                        case "修課補考標準":
                        case "修課直接指定總成績":
                        case "補修學年度":
                        case "補修學期":
                        case "重修學年度":
                        case "重修學期":
                            //case "應修學期":
                            if (value != "" && !decimal.TryParse(value, out d))
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入空白或數值");
                            }
                            break;
                        case "取得學分":
                            if (value != "是" && value != "否" && manulCheckPass.Checked)
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入是或否");
                            }
                            else if (autoCheckPass.Checked && value != "自動")
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "使用自動判斷取得學分時，取得學分欄位必須填入「自動」");
                            }
                            break;
                        case "不需評分":
                        case "不計學分":
                        case "免修":
                        case "抵免":
                        case "是否重讀":
                            if (value != "是" && value != "否" && value != "")
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入空白或是否");
                            }
                            break;
                        case "必選修":
                            if (value != "必修" && value != "選修")
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入必修或選修");
                            }
                            break;
                        case "校部訂":
                            if (value != "校訂" && value != "部訂" && value != "部定")
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入校訂或部定");
                            }
                            break;
                        case "分項類別":
                            if (value != "學業" && value != "實習科目" && value != "專業科目")
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入 學業、專業科目或實習科目");
                            }
                            break;
                        case "是否補修成績":
                            if (string.IsNullOrWhiteSpace(value))
                            {
                                e.Data[field] = "否"; // 自動補為否
                                value = "否";
                            }
                            else if (value != "是" && value != "否")
                            {
                                inputFormatPass &= false;
                                e.ErrorFields.Add(field, "必須填入是或否");
                            }
                            break;
                    }
                }
                #endregion
                //輸入格式正確才會針對情節做檢驗
                if (inputFormatPass)
                {
                    string errorMessage = "";

                    string subject = e.Data["科目"];
                    string level = e.Data["科目級別"];
                    string schoolYear = e.Data["學年度"];
                    string semester = e.Data["學期"];
                    int? sy = null;
                    int? se = null;
                    int? le = null;
                    if (int.TryParse(level, out t))
                        le = t;
                    if (int.TryParse(schoolYear, out t))
                        sy = t;
                    if (int.TryParse(semester, out t))
                        se = t;
                    if (sy != null && se != null)
                    {
                        string key = e.Data.ID + "_" + sy + "_" + se;
                        #region 驗證新增科目成績
                        bool isNewSubjectInfo = true;
                        string message = "";
                        foreach (SemesterSubjectScoreInfo info in student.SemesterSubjectScoreList)
                        {
                            if (info.SchoolYear == sy && info.Semester == se)
                            {
                                if (info.Subject == subject && info.Level == level)
                                    isNewSubjectInfo = false;
                            }
                        }
                        if (isNewSubjectInfo)
                        {
                            if (e.Data["是否補修成績"] != "是")
                            {
                                e.ErrorFields.Add("科目", "只有「是否補修成績」為「是」時，才允許新增科目成績，請修正此列資料！");
                            }
                            else
                            {
                                if (!e.WarningFields.ContainsKey("查無此科目"))
                                    e.WarningFields.Add("查無此科目", "學生在此學期並無此筆科目成績資訊，將會新增此科目成績");
                            }

                            //if (!e.WarningFields.ContainsKey("查無此科目"))
                            //    e.WarningFields.Add("查無此科目", "學生在此學期並無此筆科目成績資訊，將會新增此科目成績");
                            foreach (string field in new string[] { "科目", "科目級別", "學年度", "學期", "學分數", "分項類別", "成績年級", "必選修", "校部訂", "取得學分" })
                            {
                                if (!e.SelectFields.Contains(field))
                                    message += (message == "" ? "發現此學期無此科目，\n將會新增成績\n缺少成績必要欄位" : "、") + field;
                            }
                            if (message != "")
                                errorMessage += (errorMessage == "" ? "" : "\n") + message;
                        }
                        #endregion
                        #region 驗證重複科目資料
                        string skey = subject + "_" + le;
                        if (!_ID_SchoolYear_Semester_Subject.ContainsKey(key))
                            _ID_SchoolYear_Semester_Subject.Add(key, new List<string>());
                        if (_ID_SchoolYear_Semester_Subject[key].Contains(skey))
                        {
                            errorMessage += (errorMessage == "" ? "" : "\n") + "同一學期不允許多筆相同科目級別的資料";
                        }
                        else
                            _ID_SchoolYear_Semester_Subject[key].Add(skey);
                        #endregion
                        if (e.SelectFields.Contains("成績年級"))
                        {
                            int gy = int.Parse(e.Data["成績年級"]);
                            #region 驗證成績年級變更
                            foreach (SemesterSubjectScoreInfo info in student.SemesterSubjectScoreList)
                            {
                                if (info.SchoolYear == sy && info.Semester == se && info.GradeYear != gy)
                                {
                                    if (!e.WarningFields.ContainsKey("成績年級"))
                                        e.WarningFields.Add("成績年級", "修改成績年級資訊將會改變此學生在該學期的所有科目成績的成績年級");
                                    else
                                        e.WarningFields["成績年級"] += "、" + "修改成績年級資訊將會改變此學生在該學期的所有科目成績的成績年級";
                                    break;
                                }
                            }
                            #endregion
                            #region 驗證同學期在匯入資料中成績年級會相同

                            if (!_ID_SchoolYear_Semester_GradeYear.ContainsKey(key))
                                _ID_SchoolYear_Semester_GradeYear.Add(key, gy);
                            else
                            {
                                if (_ID_SchoolYear_Semester_GradeYear[key] != gy)
                                {
                                    if (!e.ErrorFields.ContainsKey("成績年級"))
                                        e.ErrorFields.Add("成績年級", "發現此學生同一學期有不同成績年級資料");
                                    else
                                        e.ErrorFields["成績年級"] += "、" + "發現此學生同一學期有不同成績年級資料";
                                }
                            }
                            #endregion
                            if (autoCheckPass.Checked)
                            {
                                #region 自動判斷取得學分時
                                if (!_StudentPassScore.ContainsKey(student) || !_StudentPassScore[student].ContainsKey(gy))
                                {
                                    decimal passScore = decimal.MinValue;
                                    #region 處理計算規則
                                    XmlElement scoreCalcRule = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(student.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(student.StudentID).ScoreCalcRuleElement;
                                    if (scoreCalcRule != null)
                                    {
                                        #region 取得及格標準
                                        DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                                        decimal tryParseDecimal;
                                        foreach (XmlElement element in helper.GetElements("及格標準/學生類別"))
                                        {
                                            string cat = element.GetAttribute("類別");
                                            bool useful = false;
                                            //掃描學生的類別作比對
                                            foreach (CategoryInfo catinfo in student.StudentCategorys)
                                            {
                                                if (catinfo.Name == cat || catinfo.FullName == cat)
                                                    useful = true;
                                            }
                                            //學生是指定的類別或類別為"預設"
                                            if (cat == "預設" || useful)
                                            {
                                                switch (gy)
                                                {
                                                    case 1:
                                                        if (decimal.TryParse(element.GetAttribute("一年級及格標準"), out tryParseDecimal))
                                                        {
                                                            if (passScore == decimal.MinValue || passScore > tryParseDecimal)
                                                                passScore = tryParseDecimal;
                                                        }
                                                        break;
                                                    case 2:
                                                        if (decimal.TryParse(element.GetAttribute("二年級及格標準"), out tryParseDecimal))
                                                        {
                                                            if (passScore == decimal.MinValue || passScore > tryParseDecimal)
                                                                passScore = tryParseDecimal;
                                                        }
                                                        break;
                                                    case 3:
                                                        if (decimal.TryParse(element.GetAttribute("三年級及格標準"), out tryParseDecimal))
                                                        {
                                                            if (passScore == decimal.MinValue || passScore > tryParseDecimal)
                                                                passScore = tryParseDecimal;
                                                        }
                                                        break;
                                                    case 4:
                                                        if (decimal.TryParse(element.GetAttribute("四年級及格標準"), out tryParseDecimal))
                                                        {
                                                            if (passScore == decimal.MinValue || passScore > tryParseDecimal)
                                                                passScore = tryParseDecimal;
                                                        }
                                                        break;
                                                    default:
                                                        break;
                                                }
                                            }
                                        }
                                        #endregion
                                    }
                                    #endregion
                                    if (passScore == decimal.MinValue)
                                    {
                                        e.WarningFields.Add("成績年級", "學生沒有設定成績計算規則，將以60分做為及格標準。");
                                        passScore = 60m;
                                    }
                                    if (!_StudentPassScore.ContainsKey(student))
                                        _StudentPassScore.Add(student, new Dictionary<int, decimal>());
                                    _StudentPassScore[student].Add(gy, passScore);
                                }
                                #endregion
                            }
                        }

                        // 檢查重修學年度、重修學期
                        if (e.SelectFields.Contains("重修成績") && e.SelectFields.Contains("重修學年度") && e.SelectFields.Contains("重修學期"))
                        {
                            if (e.Data["重修成績"] != "")
                            {
                                if (e.Data["重修學年度"] == "" || e.Data["重修學期"] == "")
                                {
                                    errorMessage += (errorMessage == "" ? "" : "\n") + "重修學年度、重修學期 必填!";
                                }
                            }
                        }
                    }
                    e.ErrorMessage = errorMessage;
                }
                #endregion
            };

            wizard.ImportPackage += delegate (object sender, SmartSchool.API.PlugIn.Import.ImportPackageEventArgs e)
            {
                // (1) 宣告名冊 List
                List<SubjectScoreRec108> makeUpScoreList = new List<SubjectScoreRec108>();
                List<SubjectScoreRec108> restudyScoreList = new List<SubjectScoreRec108>();


                #region ImportPackage
                Dictionary<string, List<RowData>> id_Rows = new Dictionary<string, List<RowData>>();
                #region 分包裝
                foreach (RowData data in e.Items)
                {
                    if (!id_Rows.ContainsKey(data.ID))
                        id_Rows.Add(data.ID, new List<RowData>());
                    id_Rows[data.ID].Add(data);
                }
                #endregion
                List<SmartSchool.Feature.Score.AddScore.InsertInfo> insertList = new List<SmartSchool.Feature.Score.AddScore.InsertInfo>();
                List<SmartSchool.Feature.Score.EditScore.UpdateInfo> updateList = new List<SmartSchool.Feature.Score.EditScore.UpdateInfo>();
                //交叉比對各學生資料
                #region 交叉比對各學生資料
                StringBuilder scoreDetailLogStringBuilder = new StringBuilder();
                StringBuilder updateLogStringBuilder = new StringBuilder();
                StringBuilder insertLogStringBuilder = new StringBuilder();
                foreach (string id in id_Rows.Keys)
                {
                    XmlDocument doc = new XmlDocument();
                    StudentRecord studentRec = _StudentCollection[id];
                    //該學生的學期科目成績
                    Dictionary<int, Dictionary<int, Dictionary<string, SemesterSubjectScoreInfo>>> semesterScoreDictionary = new Dictionary<int, Dictionary<int, Dictionary<string, SemesterSubjectScoreInfo>>>();
                    #region 整理現有的成績資料
                    foreach (SemesterSubjectScoreInfo var in studentRec.SemesterSubjectScoreList)
                    {
                        string key = var.Subject + "_" + var.Level;
                        if (!semesterScoreDictionary.ContainsKey(var.SchoolYear))
                            semesterScoreDictionary.Add(var.SchoolYear, new Dictionary<int, Dictionary<string, SemesterSubjectScoreInfo>>());
                        if (!semesterScoreDictionary[var.SchoolYear].ContainsKey(var.Semester))
                            semesterScoreDictionary[var.SchoolYear].Add(var.Semester, new Dictionary<string, SemesterSubjectScoreInfo>());
                        if (!semesterScoreDictionary[var.SchoolYear][var.Semester].ContainsKey(key))
                            semesterScoreDictionary[var.SchoolYear][var.Semester].Add(key, var);
                    }
                    #endregion
                    //要匯入的學期科目成績
                    Dictionary<int, Dictionary<int, Dictionary<string, RowData>>> semesterImportScoreDictionary = new Dictionary<int, Dictionary<int, Dictionary<string, RowData>>>();
                    #region 整理要匯入的資料
                    foreach (RowData row in id_Rows[id])
                    {
                        int t;
                        string subject = row["科目"];
                        string level = row["科目級別"];
                        string schoolYear = row["學年度"];
                        string semester = row["學期"];
                        int sy = int.Parse(schoolYear);
                        int se = int.Parse(semester);
                        string key = subject + "_" + level;
                        if (!semesterImportScoreDictionary.ContainsKey(sy))
                            semesterImportScoreDictionary.Add(sy, new Dictionary<int, Dictionary<string, RowData>>());
                        if (!semesterImportScoreDictionary[sy].ContainsKey(se))
                            semesterImportScoreDictionary[sy].Add(se, new Dictionary<string, RowData>());
                        if (!semesterImportScoreDictionary[sy][se].ContainsKey(key))
                            semesterImportScoreDictionary[sy][se].Add(key, row);

                        // 處理寫入名冊需要資料
                        // (3) 同步處理名冊
                        if (row["是否補修成績"] == "是")
                        {
                            // --- 處理補修成績寫入學期歷程資料 4.3 補修成績
                            string HisClassName = "", HisStudentNumber = "";
                            int? HisSeatNo = null;
                            try
                            {

                                // 讀取學生學期對照表資料
                                if (studentRec.Fields.ContainsKey("SemesterHistory"))
                                {
                                    XmlElement xmlElement = studentRec.Fields["SemesterHistory"] as XmlElement;
                                    XElement elmRoot = XElement.Parse(xmlElement.OuterXml);

                                    //// 找到第一個符合 SchoolYear, Semester 的節點
                                    //var matched = elmRoot.Elements("History")
                                    //    .FirstOrDefault(e =>
                                    //        (string)e.Attribute("SchoolYear") == sy.ToString() &&
                                    //        (string)e.Attribute("Semester") == se.ToString());

                                    //if (matched != null)
                                    //{
                                    //    HisClassName = (string)matched.Attribute("ClassName") ?? "";
                                    //    HisStudentNumber = (string)matched.Attribute("StudentNumber") ?? "";

                                    //    // SeatNo 轉型
                                    //    int seatNo;
                                    //    if (int.TryParse((string)matched.Attribute("SeatNo"), out seatNo))
                                    //        HisSeatNo = seatNo;
                                    //    else
                                    //        HisSeatNo = null;
                                    //}
                                    //else
                                    //{

                                    //}
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                            string ScoreP = "";

                            // 根據 makeUpScoreInfo.Detail.GetAttribute("是否取得學分") 的值，如果是 "是" 就設定 ssr.ScoreP 為 1，否則為 0。
                            // Please use if-else syntax.
                            if (GetRowDataCellValue(row, "取得學分") == "是")
                                ScoreP = "1";
                            else
                                ScoreP = "0";


                            string reScore = "-1";
                            // 幫我寫一段 C# 程式碼，從 makeUpScoreInfo.Detail 取得「補考成績」和「修課及格標準」兩個屬性。如果「補考成績」能轉為數字，且在 0 到「修課及格標準」之間，則 reScore 等於該分數字串，否則 reScore = "-1"。如果「修課及格標準」無法轉數字，預設用 60。                            
                            if (row.ContainsKey("補考成績"))
                                if (decimal.TryParse(row["補考成績"].ToString(), out decimal reScoreValue2))
                                {
                                    if (decimal.TryParse(GetRowDataCellValue(row, "修課及格標準"), out decimal passingStandard))
                                    {
                                        if (reScoreValue2 >= 0 && reScoreValue2 <= passingStandard)
                                            reScore = reScoreValue2.ToString();
                                        else
                                            reScore = "-1";
                                    }
                                    else
                                    {
                                        reScore = "-1"; // 如果無法轉數字，預設為 -1
                                    }
                                }



                            // 這段式 ChatGPT 寫如果 reScore 是 "-1" 則 reScoreP 設為 "-1"。否則嘗試把 reScore 轉成數字，並把 makeUpScoreInfo.Detail.GetAttribute("修課及格標準") 轉成數字（預設 60）。如果 reScore 大於等於修課及格標準，reScoreP 設為 "1"，否則設為 "0"。
                            string reScoreP = "-1";

                            // 取得修課及格標準（預設 60）                            
                            string passScoreStr = GetRowDataCellValue(row, "修課及格標準");
                            decimal passScore = 60;
                            decimal.TryParse(passScoreStr, out passScore);

                            if (reScore == "-1")
                            {
                                reScoreP = "-1";
                            }
                            else if (decimal.TryParse(reScore, out decimal reScoreDecimal) && reScoreDecimal >= passScore)
                            {
                                reScoreP = "1";
                            }
                            else
                            {
                                reScoreP = "0";
                            }


                            // 課程代碼
                            string courseCode = GetRowDataCellValue(row,"課程代碼");

                            // 假設 makeUpScoreInfo.Detail 是 XElement 或有 GetAttribute 方法
                            string notCountCredit = GetRowDataCellValue(row, "不計學分"); 
                            string notNeedScore = GetRowDataCellValue(row, "不需評分"); 

                            // 是否採計學分 預設
                            string useCredit = "1";

                            // 不計學分為"是"
                            if (notCountCredit == "是")
                            {
                                useCredit = "2";
                            }
                            else
                            {
                                // 預設為"1"，但如果不需評分為"是"或課程代碼特殊則為"3"
                                bool setTo3 = false;

                                if (notNeedScore == "是")
                                    setTo3 = true;

                                if (!string.IsNullOrWhiteSpace(courseCode) && courseCode.Length > 22)
                                {
                                    string sub1 = courseCode.Substring(16, 1);
                                    string sub2 = courseCode.Substring(18, 1);
                                    if (sub1 == "9" && sub2 == "D")
                                        setTo3 = true;
                                }

                                if (setTo3)
                                    useCredit = "3";
                            }

                            // 檢查課程代碼 CodePass
                            bool codePass = Utility.IsValidCourseCode(courseCode);
                            codePass = true;

                            if (string.IsNullOrWhiteSpace(HisClassName))
                                HisClassName = studentRec.RefClass.ClassName;

                            if (!HisSeatNo.HasValue || HisSeatNo == 0)
                            {
                                if (int.TryParse(studentRec.SeatNo, out int seatNo))
                                    HisSeatNo = seatNo;
                            }

                            if (string.IsNullOrWhiteSpace(HisStudentNumber))
                                HisStudentNumber = studentRec.StudentNumber;


                            var rec = new SubjectScoreRec108
                            {
                                StudentID = studentRec.StudentID,
                                IDNumber = studentRec.IDNumber,
                                Birthday = Utility.ConvertChDateString(studentRec.Birthday),
                                SchoolYear = schoolYear,
                                Semester = semester,
                                CourseCode = GetRowDataCellValue(row, "課程代碼"),
                                SubjectName = GetRowDataCellValue(row, "科目"),
                                SubjectLevel = GetRowDataCellValue(row, "科目級別"),
                                GradeYear = GetRowDataCellValue(row, "成績年級"),
                                Credit = GetRowDataCellValue(row, "學分數"),
                                Score = GetRowDataCellValue(row, "原始成績"),
                                ScoreP = ScoreP,
                                ReScore = reScore,
                                ReScoreP = reScoreP,
                                ScScoreType = "3",
                                useCredit = useCredit,
                                Text = "",
                                Name = studentRec.StudentName,
                                HisClassName = HisClassName,
                                HisSeatNo = HisSeatNo,
                                HisStudentNumber = HisStudentNumber,
                                ClassName = studentRec.RefClass.ClassName,
                                SeatNo = studentRec.SeatNo,
                                StudentNumber = studentRec.StudentNumber,
                                isScScore = true,
                                checkPass = true,
                                CodePass = codePass
                            };
                            makeUpScoreList.Add(rec);
                        }
                        else if (row["重修成績"] != "")
                        {
                            // 重修成績寫入學期歷程 ---

                            string HisClassName = "", HisStudentNumber = "";
                            int? HisSeatNo = null;
                            try
                            {
                                // 讀取學生學期對照表資料
                                if (studentRec.Fields.ContainsKey("SemesterHistory"))
                                {
                                    XmlElement xmlElement = studentRec.Fields["SemesterHistory"] as XmlElement;
                                    XElement elmRoot = XElement.Parse(xmlElement.OuterXml);

                                    //// 找到第一個符合 SchoolYear, Semester 的節點
                                    //var matched = elmRoot.Elements("History")
                                    //    .FirstOrDefault(e =>
                                    //        (string)e.Attribute("SchoolYear") == sy.ToString() &&
                                    //        (string)e.Attribute("Semester") == se.ToString());

                                    //if (matched != null)
                                    //{
                                    //    HisClassName = (string)matched.Attribute("ClassName") ?? "";
                                    //    HisStudentNumber = (string)matched.Attribute("StudentNumber") ?? "";

                                    //    // SeatNo 轉型
                                    //    int seatNo;
                                    //    if (int.TryParse((string)matched.Attribute("SeatNo"), out seatNo))
                                    //        HisSeatNo = seatNo;
                                    //    else
                                    //        HisSeatNo = null;
                                    //}
                                    //else
                                    //{

                                    //}
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }

                            string ScoreP = "";

                            // 根據 makeUpScoreInfo.Detail.GetAttribute("是否取得學分") 的值，如果是 "是" 就設定 ssr.ScoreP 為 1，否則為 0。
                            // Please use if-else syntax.
                            if (GetRowDataCellValue(row, "取得學分") == "是")
                                ScoreP = "1";
                            else
                                ScoreP = "0";


                            string ReAScore = "-1";
                            // 幫我寫一段 C# 程式碼，從 makeUpScoreInfo.Detail 取得「補考成績」和「修課及格標準」兩個屬性。如果「補考成績」能轉為數字，且在 0 到「修課及格標準」之間，則 reScore 等於該分數字串，否則 reScore = "-1"。如果「修課及格標準」無法轉數字，預設用 60。
                            if (decimal.TryParse(GetRowDataCellValue(row, "重修成績"), out decimal reScoreValue2))
                            {
                                if (decimal.TryParse(GetRowDataCellValue(row, "修課及格標準"), out decimal passingStandard))
                                {

                                    if (reScoreValue2 >= 0)
                                        ReAScore = reScoreValue2.ToString();
                                    else
                                        ReAScore = "-1";
                                }
                                else
                                {
                                    ReAScore = "-1"; // 如果無法轉數字，預設為 -1
                                }
                            }



                            // 這段式 ChatGPT 寫如果 reScore 是 "-1" 則 ReAScoreP 設為 "-1"。否則嘗試把 reScore 轉成數字，並把 makeUpScoreInfo.Detail.GetAttribute("修課及格標準") 轉成數字（預設 60）。如果 reScore 大於等於修課及格標準，ReAScoreP 設為 "1"，否則設為 "0"。
                            string ReAScoreP = "-1";

                            // 取得修課及格標準（預設 60）
                            string passScoreStr = GetRowDataCellValue(row, "修課及格標準");
                            decimal passScore = 60;
                            decimal.TryParse(passScoreStr, out passScore);

                            if (ReAScore == "-1")
                            {
                                ReAScoreP = "-1";
                            }
                            else if (decimal.TryParse(ReAScore, out decimal reScoreDecimal) && reScoreDecimal >= passScore)
                            {
                                ReAScoreP = "1";
                            }
                            else
                            {
                                ReAScoreP = "0";
                            }


                            // 課程代碼
                            string courseCode = GetRowDataCellValue(row, "課程代碼");

                            // 假設 makeUpScoreInfo.Detail 是 XElement 或有 GetAttribute 方法
                            string notCountCredit = GetRowDataCellValue(row, "不計學分");
                            string notNeedScore = GetRowDataCellValue(row, "不需評分");

                            // 是否採計學分 預設
                            string useCredit = "1";

                            // 不計學分為"是"
                            if (notCountCredit == "是")
                            {
                                useCredit = "2";
                            }
                            else
                            {
                                // 預設為"1"，但如果不需評分為"是"或課程代碼特殊則為"3"
                                bool setTo3 = false;

                                if (notNeedScore == "是")
                                    setTo3 = true;

                                if (!string.IsNullOrWhiteSpace(courseCode) && courseCode.Length > 22)
                                {
                                    string sub1 = courseCode.Substring(16, 1);
                                    string sub2 = courseCode.Substring(18, 1);
                                    if (sub1 == "9" && sub2 == "D")
                                        setTo3 = true;
                                }

                                if (setTo3)
                                    useCredit = "3";
                            }

                            // 檢查課程代碼 CodePass
                            bool codePass = Utility.IsValidCourseCode(courseCode);
                            codePass = true;
                            if (string.IsNullOrWhiteSpace(HisClassName))
                                HisClassName = studentRec.RefClass.ClassName;

                            if (!HisSeatNo.HasValue || HisSeatNo == 0)
                            {
                                if (int.TryParse(studentRec.SeatNo, out int seatNo))
                                    HisSeatNo = seatNo;
                            }

                            if (string.IsNullOrWhiteSpace(HisStudentNumber))
                                HisStudentNumber = studentRec.StudentNumber;

                            var rec = new SubjectScoreRec108
                            {
                                StudentID = studentRec.StudentID,
                                IDNumber = studentRec.IDNumber,
                                Birthday = Utility.ConvertChDateString(studentRec.Birthday),
                                SchoolYear = schoolYear,
                                Semester = semester,
                                CourseCode = GetRowDataCellValue(row, "課程代碼"),
                                SubjectName = GetRowDataCellValue(row, "科目"),
                                SubjectLevel = GetRowDataCellValue(row, "科目級別"),
                                GradeYear = GetRowDataCellValue(row, "成績年級"),
                                Credit = GetRowDataCellValue(row, "學分數"),
                                Score = GetRowDataCellValue(row, "重修成績"),

                                ScoreP = ScoreP,
                                ScScoreType = "3",
                                useCredit = useCredit,
                                Text = "",
                                Name = studentRec.StudentName,
                                HisClassName = HisClassName,
                                HisSeatNo = HisSeatNo,
                                HisStudentNumber = HisStudentNumber,
                                ClassName = studentRec.RefClass.ClassName,
                                SeatNo = studentRec.SeatNo,
                                StudentNumber = studentRec.StudentNumber,
                                isScScore = true,
                                checkPass = true,
                                CodePass = codePass,
                                ReAScoreType = "3",
                                ReAScore = ReAScore,
                                ReAScoreP = ReAScoreP
                            };

                            restudyScoreList.Add(rec);
                        }

                    }
                    #endregion

                    //學期年級重整
                    Dictionary<int, Dictionary<int, int>> semesterGradeYear = new Dictionary<int, Dictionary<int, int>>();
                    //要變更成績的學期
                    Dictionary<int, List<int>> updatedSemester = new Dictionary<int, List<int>>();
                    //在變更學期中新增加的成績資料
                    Dictionary<int, Dictionary<int, List<RowData>>> updatedNewSemesterScore = new Dictionary<int, Dictionary<int, List<RowData>>>();
                    //要增加成績的學期
                    Dictionary<int, Dictionary<int, List<RowData>>> insertNewSemesterScore = new Dictionary<int, Dictionary<int, List<RowData>>>();
                    //開始處理ImportScore
                    #region 開始處理ImportScore
                    foreach (int sy in semesterImportScoreDictionary.Keys)
                    {
                        foreach (int se in semesterImportScoreDictionary[sy].Keys)
                        {
                            foreach (string key in semesterImportScoreDictionary[sy][se].Keys)
                            {
                                RowData data = semesterImportScoreDictionary[sy][se][key];
                                //如果是本來沒有這筆學期的成績就加到insertNewSemesterScore
                                if (!semesterScoreDictionary.ContainsKey(sy) || !semesterScoreDictionary[sy].ContainsKey(se))
                                {
                                    if (!insertNewSemesterScore.ContainsKey(sy))
                                        insertNewSemesterScore.Add(sy, new Dictionary<int, List<RowData>>());
                                    if (!insertNewSemesterScore[sy].ContainsKey(se))
                                        insertNewSemesterScore[sy].Add(se, new List<RowData>());
                                    insertNewSemesterScore[sy][se].Add(data);
                                    //加入學期年級
                                    int gy = int.Parse(data["成績年級"]);
                                    if (!semesterGradeYear.ContainsKey(sy))
                                        semesterGradeYear.Add(sy, new Dictionary<int, int>());
                                    if (!semesterGradeYear[sy].ContainsKey(se))
                                        semesterGradeYear[sy].Add(se, gy);
                                    else
                                        semesterGradeYear[sy][se] = gy;
                                }
                                else
                                {
                                    bool hasChanged = false;
                                    //修改已存在的資料
                                    if (semesterScoreDictionary[sy][se].ContainsKey(key))
                                    {
                                        SemesterSubjectScoreInfo score = semesterScoreDictionary[sy][se][key];
                                        #region 填入此學期的年級資料
                                        if (!semesterGradeYear.ContainsKey(sy))
                                            semesterGradeYear.Add(sy, new Dictionary<int, int>());
                                        if (!semesterGradeYear[sy].ContainsKey(se))
                                            semesterGradeYear[sy].Add(se, score.GradeYear);
                                        #endregion
                                        #region 直接修改已存在的成績資料的Detail
                                        string logLine = "學生系統編號：「" + id + "」學生姓名：「" + studentRec.StudentName + "」，更新科目：學年度「" + sy + "」、學期「" + se + "」、科目名稱「" + data["科目"] + "」、科目級別「" + data["科目級別"] + "」";
                                        foreach (string field in e.ImportFields)
                                        {
                                            string value = data[field];
                                            switch (field)
                                            {
                                                default: break;
                                                case "領域":
                                                    if (score.Detail.GetAttribute("領域") != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute("領域") + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute("領域", value);
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "學分數":
                                                    if (score.Detail.GetAttribute("開課學分數") != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute("開課學分數") + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute("開課學分數", value);
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "分項類別":
                                                    if (score.Detail.GetAttribute("開課分項類別") != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute("開課分項類別") + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute("開課分項類別", value);
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "必選修":
                                                    if (score.Detail.GetAttribute("修課必選修") != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute("修課必選修") + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute("修課必選修", value);
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "校部訂":
                                                    value = (value == "部定" ? "部訂" : value);
                                                    if (score.Detail.GetAttribute("修課校部訂") != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute("修課校部訂") == "部訂" ? "部定" : score.Detail.GetAttribute("修課校部訂") + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute("修課校部訂", value); ;
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "取得學分":
                                                    if (score.Detail.GetAttribute("是否取得學分") != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute("是否取得學分") + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute("是否取得學分", value);
                                                        hasChanged = true;
                                                    }
                                                    break;

                                                case "手動調整成績":
                                                    if (score.Detail.GetAttribute("擇優採計成績") != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute("擇優採計成績") + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute("擇優採計成績", value);
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "原始成績":
                                                case "補考成績":
                                                case "重修成績":
                                                case "學年調整成績":
                                                    if (score.Detail.GetAttribute(field) != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute(field) + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute(field, value);
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "不計學分":
                                                case "不需評分":
                                                case "免修":
                                                case "抵免":
                                                case "是否補修成績":
                                                case "是否重讀":
                                                    value = (value == "" ? "否" : value);
                                                    if (score.Detail.GetAttribute(field) != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute(field) + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute(field, value);
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "成績年級":
                                                    int gy = int.Parse(data["成績年級"]);
                                                    if (score.GradeYear != gy)
                                                    {
                                                        logLine += "、" + field + "由「" + score.GradeYear + "」變更為「" + gy + "」";
                                                        semesterGradeYear[sy][se] = gy;
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "註記":
                                                case "重修學年度":
                                                case "重修學期":
                                                case "補修學年度":
                                                case "補修學期":
                                                    if (score.Detail.GetAttribute(field) != value)
                                                    {
                                                        score.Detail.SetAttribute(field, value);
                                                        logLine += "、" + field + "「" + value + "」";
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "指定學年科目名稱":
                                                case "報部科目名稱":
                                                case "英文名稱":
                                                //if (score.Detail.GetAttribute(field) != value)
                                                //{
                                                //    logLine += "、" + field + "由「" + score.Detail.GetAttribute(field) + "」變更為「" + value + "」";
                                                //    score.Detail.SetAttribute(field, value);
                                                //    hasChanged = true;
                                                //}
                                                //break;
                                                case "修課及格標準":
                                                case "修課補考標準":
                                                case "修課直接指定總成績":
                                                case "修課備註":
                                                    if (score.Detail.GetAttribute(field) != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute(field) + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute(field, value);
                                                        hasChanged = true;
                                                    }
                                                    break;
                                                case "課程代碼":
                                                    if (score.Detail.GetAttribute("修課科目代碼") != value)
                                                    {
                                                        logLine += "、" + field + "由「" + score.Detail.GetAttribute("修課科目代碼") + "」變更為「" + value + "」";
                                                        score.Detail.SetAttribute("修課科目代碼", value);
                                                        hasChanged = true;
                                                    }
                                                    break;
                                            }
                                        }
                                        #endregion
                                        if (autoCheckPass.Checked)
                                        {
                                            int gy = int.Parse(data["成績年級"]);
                                            #region 做取得學分判斷及填入擇優採計成績
                                            //最高分
                                            decimal maxScore = decimal.MinValue;
                                            #region 抓最高分
                                            string[] scoreNames = new string[] { "原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績" };
                                            foreach (string scorename in scoreNames)
                                            {
                                                decimal s;
                                                if (decimal.TryParse(score.Detail.GetAttribute(scorename), out s))
                                                {
                                                    if (s > maxScore)
                                                    {
                                                        maxScore = s;
                                                    }
                                                }
                                            }
                                            #endregion
                                            #endregion
                                            score.Detail.SetAttribute("是否取得學分", ((score.Detail.GetAttribute("不需評分") == "是") || maxScore >= _StudentPassScore[studentRec][gy]) ? "是" : "否");
                                        }
                                        if (hasChanged)
                                            scoreDetailLogStringBuilder.AppendLine(logLine);

                                    }
                                    else//加入新成績至已存在的學期
                                    {
                                        //加入學期年級
                                        int gy = int.Parse(data["成績年級"]);
                                        if (!semesterGradeYear.ContainsKey(sy))
                                            semesterGradeYear.Add(sy, new Dictionary<int, int>());
                                        if (!semesterGradeYear[sy].ContainsKey(se))
                                            semesterGradeYear[sy].Add(se, gy);
                                        else
                                            semesterGradeYear[sy][se] = gy;
                                        //加入新成績至已存在的學期
                                        if (!updatedNewSemesterScore.ContainsKey(sy))
                                            updatedNewSemesterScore.Add(sy, new Dictionary<int, List<RowData>>());
                                        if (!updatedNewSemesterScore[sy].ContainsKey(se))
                                            updatedNewSemesterScore[sy].Add(se, new List<RowData>());
                                        updatedNewSemesterScore[sy][se].Add(data);
                                        hasChanged = true;
                                    }
                                    //真的有變更
                                    if (hasChanged)
                                    {
                                        #region 登錄有變更的學期
                                        if (!updatedSemester.ContainsKey(sy))
                                            updatedSemester.Add(sy, new List<int>());
                                        if (!updatedSemester[sy].Contains(se))
                                            updatedSemester[sy].Add(se);
                                        #endregion
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                    //處理已登錄要更新的學期成績
                    #region 處理已登錄要更新的學期成績
                    foreach (int sy in updatedSemester.Keys)
                    {
                        foreach (int se in updatedSemester[sy])
                        {
                            Dictionary<int, Dictionary<int, string>> semeScoreID = (Dictionary<int, Dictionary<int, string>>)studentRec.Fields["SemesterSubjectScoreID"];
                            string semesterScoreID = semeScoreID[sy][se];//從學期抓ID
                            string gradeyear = "" + semesterGradeYear[sy][se];//抓年級
                            XmlElement subjectScoreInfo = doc.CreateElement("SemesterSubjectScoreInfo");
                            #region 產生該學期科目成績的XML
                            foreach (SemesterSubjectScoreInfo scoreInfo in semesterScoreDictionary[sy][se].Values)
                            {
                                subjectScoreInfo.AppendChild(doc.ImportNode(scoreInfo.Detail, true));
                            }
                            if (updatedNewSemesterScore.ContainsKey(sy) && updatedNewSemesterScore[sy].ContainsKey(se))
                            {
                                foreach (RowData row in updatedNewSemesterScore[sy][se])
                                {
                                    XmlElement newScore = doc.CreateElement("Subject");
                                    string logLine = "學生系統編號：「" + id + "」學生姓名：「" + studentRec.StudentName + "」，新增科目：學年度「" + sy + "」、學期「" + se + "」";

                                    #region 建立newScore
                                    foreach (string field in new string[] { "領域", "科目", "科目級別", "學分數", "分項類別", "必選修", "校部訂", "原始成績", "補考成績", "重修成績", "手動調整成績", "學年調整成績", "取得學分", "不計學分", "不需評分", "免修", "抵免", "補修學年度", "補修學期", "是否補修成績", "指定學年科目名稱", "課程代碼", "報部科目名稱" })
                                    {
                                        if (e.ImportFields.Contains(field))
                                        {
                                            string value = row[field];
                                            switch (field)
                                            {
                                                default: break;
                                                case "領域":
                                                    newScore.SetAttribute("領域", value);
                                                    if (value != "")
                                                        logLine += "、" + field + "「" + value + "」";
                                                    break;
                                                case "科目":
                                                    newScore.SetAttribute("科目", value);
                                                    logLine += "、科目「" + value + "」";
                                                    break;
                                                case "科目級別":
                                                    int level;
                                                    if (int.TryParse(value, out level))
                                                        newScore.SetAttribute("科目級別", "" + level);
                                                    else
                                                        newScore.SetAttribute("科目級別", "");
                                                    logLine += "、科目級別「" + value + "」";
                                                    break;
                                                case "學分數":
                                                    newScore.SetAttribute("開課學分數", value);
                                                    logLine += "、學分數「" + value + "」";
                                                    break;
                                                case "分項類別":
                                                    newScore.SetAttribute("開課分項類別", value);
                                                    logLine += "、分項類別「" + value + "」";
                                                    break;
                                                case "必選修":
                                                    newScore.SetAttribute("修課必選修", value);
                                                    logLine += "、必選修「" + value + "」";
                                                    break;
                                                case "校部訂":
                                                    newScore.SetAttribute("修課校部訂", value == "部定" ? "部訂" : value);
                                                    logLine += "、校部訂「" + value + "」";
                                                    break;
                                                case "取得學分":
                                                    newScore.SetAttribute("是否取得學分", value);
                                                    if (value != "")
                                                        logLine += "、" + field + "「" + value + "」";
                                                    break;

                                                case "手動調整成績":
                                                    newScore.SetAttribute("擇優採計成績", value);
                                                    if (value != "")
                                                        logLine += "、" + field + "「" + value + "」";
                                                    break;
                                                case "註記":
                                                case "重修學年度":
                                                case "重修學期":
                                                case "補修學年度":
                                                case "補修學期":
                                                case "英文名稱":
                                                case "原始成績":
                                                case "補考成績":
                                                case "重修成績":
                                                case "學年調整成績":
                                                case "修課及格標準":
                                                case "修課補考標準":
                                                case "修課直接指定總成績":
                                                case "修課備註":
                                                    newScore.SetAttribute(field, value);
                                                    if (value != "")
                                                        logLine += "、" + field + "「" + value + "」";
                                                    break;
                                                case "指定學年科目名稱":
                                                case "報部科目名稱":
                                                    newScore.SetAttribute(field, value);
                                                    if (value != "")
                                                        logLine += "、" + field + "「" + value + "」";
                                                    break;
                                                case "不計學分":
                                                case "不需評分":
                                                case "免修":
                                                case "抵免":
                                                case "是否補修成績":
                                                case "是否重讀":
                                                    value = (value == "" ? "否" : value);
                                                    newScore.SetAttribute(field, value);
                                                    if (value != "")
                                                        logLine += "、" + field + "「" + value + "」";
                                                    break;
                                                case "課程代碼":
                                                    newScore.SetAttribute("修課科目代碼", value);
                                                    if (value != "")
                                                        logLine += "、" + field + "「" + value + "」";
                                                    break;

                                            }
                                        }
                                    }
                                    //這兩個是為二有預設值的
                                    foreach (string field in new string[] { "不計學分", "不需評分" })
                                    {
                                        if (!e.ImportFields.Contains(field))
                                        {
                                            newScore.SetAttribute(field, "否");
                                        }
                                    }
                                    if (autoCheckPass.Checked)
                                    {
                                        int gy = int.Parse(row["成績年級"]);
                                        #region 做取得學分判斷及填入擇優採計成績
                                        //最高分
                                        decimal maxScore = decimal.MinValue;
                                        #region 抓最高分
                                        string[] scoreNames = new string[] { "原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績" };
                                        foreach (string scorename in scoreNames)
                                        {
                                            decimal s;
                                            if (decimal.TryParse(newScore.GetAttribute(scorename), out s))
                                            {
                                                if (s > maxScore)
                                                {
                                                    maxScore = s;
                                                }
                                            }
                                        }
                                        #endregion
                                        #endregion
                                        newScore.SetAttribute("是否取得學分", ((newScore.GetAttribute("不需評分") == "是") || maxScore >= _StudentPassScore[studentRec][gy]) ? "是" : "否");
                                    }
                                    #endregion
                                    subjectScoreInfo.AppendChild(newScore);
                                    updateLogStringBuilder.AppendLine(logLine);
                                }
                            }
                            #endregion
                            updateList.Add(new SmartSchool.Feature.Score.EditScore.UpdateInfo(semesterScoreID, gradeyear, subjectScoreInfo));
                        }
                    }
                    #endregion
                    //處理新增成績學期
                    #region 處理新增成績學期
                    foreach (int sy in insertNewSemesterScore.Keys)
                    {
                        foreach (int se in insertNewSemesterScore[sy].Keys)
                        {
                            XmlElement subjectScoreInfo = doc.CreateElement("SemesterSubjectScoreInfo");

                            string gradeyear = "" + semesterGradeYear[sy][se];//抓年級

                            string logLine = "學生系統編號：「" + id + "」學生姓名：「" + studentRec.StudentName + "」，新增科目：學年度「" + sy + "」、學期「" + se + "」、成績年級「" + gradeyear + "」";
                            foreach (RowData row in insertNewSemesterScore[sy][se])
                            {
                                XmlElement newScore = doc.CreateElement("Subject");
                                #region 建立newScore
                                foreach (string field in new string[] { "領域", "科目", "科目級別", "學分數", "分項類別", "必選修", "校部訂", "原始成績", "補考成績", "重修成績", "手動調整成績", "學年調整成績", "取得學分", "不計學分", "不需評分", "是否補修成績", "重修學年度", "重修學期", "免修", "抵免", "補修學年度", "補修學期", "指定學年科目名稱", "課程代碼", "報部科目名稱", "是否重讀" })
                                {
                                    if (e.ImportFields.Contains(field))
                                    {
                                        string value = row[field];
                                        switch (field)
                                        {
                                            default: break;
                                            case "領域":
                                                newScore.SetAttribute("領域", value);
                                                if (value != "")
                                                    logLine += "、" + field + "「" + value + "」";
                                                break;
                                            case "科目":
                                                newScore.SetAttribute("科目", value);
                                                logLine += "、科目「" + value + "」";
                                                break;
                                            case "科目級別":
                                                int level;
                                                if (int.TryParse(value, out level))
                                                    newScore.SetAttribute("科目級別", "" + level);
                                                else
                                                    newScore.SetAttribute("科目級別", "");
                                                logLine += "、科目級別「" + value + "」";
                                                break;
                                            case "學分數":
                                                newScore.SetAttribute("開課學分數", value);
                                                logLine += "、學分數「" + value + "」";
                                                break;
                                            case "分項類別":
                                                newScore.SetAttribute("開課分項類別", value);
                                                logLine += "、分項類別「" + value + "」";
                                                break;
                                            case "必選修":
                                                newScore.SetAttribute("修課必選修", value);
                                                logLine += "、必選修「" + value + "」";
                                                break;
                                            case "校部訂":
                                                newScore.SetAttribute("修課校部訂", value == "部定" ? "部訂" : value);
                                                logLine += "、校部訂「" + value + "」";
                                                break;
                                            case "取得學分":
                                                newScore.SetAttribute("是否取得學分", value);
                                                if (value != "")
                                                    logLine += "、" + field + "「" + value + "」";
                                                break;

                                            case "手動調整成績":
                                                newScore.SetAttribute("擇優採計成績", value);
                                                if (value != "")
                                                    logLine += "、" + field + "「" + value + "」";
                                                break;
                                            case "原始成績":
                                            case "補考成績":
                                            case "重修成績":
                                            case "學年調整成績":
                                            case "重修學年度":
                                            case "重修學期":
                                            case "補修學年度":
                                            case "補修學期":
                                                newScore.SetAttribute(field, value);
                                                if (value != "")
                                                    logLine += "、" + field + "「" + value + "」";
                                                break;
                                            case "不計學分":
                                            case "不需評分":
                                            case "免修":
                                            case "抵免":
                                            case "是否補修成績":
                                            case "是否重讀":
                                                value = (value == "" ? "否" : value);
                                                newScore.SetAttribute(field, value);
                                                if (value != "")
                                                    logLine += "、" + field + "「" + value + "」";
                                                break;
                                            case "指定學年科目名稱":
                                            case "報部科目名稱":
                                                newScore.SetAttribute(field, value);
                                                if (value != "")
                                                    logLine += "、" + field + "「" + value + "」";
                                                break;
                                            case "課程代碼":
                                                newScore.SetAttribute("修課科目代碼", value);
                                                if (value != "")
                                                    logLine += "、" + field + "「" + value + "」";
                                                break;
                                        }
                                    }
                                }
                                //這兩個是為二有預設值的
                                foreach (string field in new string[] { "不計學分", "不需評分" })
                                {
                                    if (!e.ImportFields.Contains(field))
                                    {
                                        newScore.SetAttribute(field, "否");
                                    }
                                }

                                if (autoCheckPass.Checked)
                                {
                                    int gy = int.Parse(row["成績年級"]);
                                    #region 做取得學分判斷及填入擇優採計成績
                                    //最高分
                                    decimal maxScore = decimal.MinValue;
                                    #region 抓最高分
                                    string[] scoreNames = new string[] { "原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績" };
                                    foreach (string scorename in scoreNames)
                                    {
                                        decimal s;
                                        if (decimal.TryParse(newScore.GetAttribute(scorename), out s))
                                        {
                                            if (s > maxScore)
                                            {
                                                maxScore = s;
                                            }
                                        }
                                    }
                                    #endregion
                                    #endregion
                                    newScore.SetAttribute("是否取得學分", ((newScore.GetAttribute("不需評分") == "是") || maxScore >= _StudentPassScore[studentRec][gy]) ? "是" : "否");
                                }
                                #endregion
                                subjectScoreInfo.AppendChild(newScore);
                                insertLogStringBuilder.AppendLine(logLine);
                            }
                            insertList.Add(new SmartSchool.Feature.Score.AddScore.InsertInfo(studentRec.StudentID, "" + sy, "" + se, gradeyear, "", subjectScoreInfo));
                        }
                    }
                    #endregion
                }
                #endregion

                // 寫入名冊學年度學期
                int SchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
                int Semester = int.Parse(K12.Data.School.DefaultSemester);

                // (4) 兩層迴圈後，一次性寫入名冊
                if (makeUpScoreList.Count > 0)
                    new LearningHistoryDataAccess().SaveScores43(makeUpScoreList, SchoolYear, Semester);
                if (restudyScoreList.Count > 0)
                    new LearningHistoryDataAccess().SaveScores52(restudyScoreList, SchoolYear, Semester);


                if (scoreDetailLogStringBuilder.Length > 0 || updateList.Count > 0 || insertList.Count > 0)
                    FISCA.LogAgent.ApplicationLog.Log("匯入學期科目成績", "匯入", scoreDetailLogStringBuilder.ToString() + updateLogStringBuilder.ToString() + insertLogStringBuilder.ToString());
                if (updateList.Count > 0)
                {
                    //FISCA.LogAgent.ApplicationLog.Log("匯入學期科目成績", "匯入", updateLogStringBuilder.ToString());

                    #region 分批次兩路上傳
                    List<List<SmartSchool.Feature.Score.EditScore.UpdateInfo>> updatePackages = new List<List<SmartSchool.Feature.Score.EditScore.UpdateInfo>>();
                    List<List<SmartSchool.Feature.Score.EditScore.UpdateInfo>> updatePackages2 = new List<List<SmartSchool.Feature.Score.EditScore.UpdateInfo>>();
                    {
                        List<SmartSchool.Feature.Score.EditScore.UpdateInfo> package = null;
                        int count = 0;
                        foreach (SmartSchool.Feature.Score.EditScore.UpdateInfo var in updateList)
                        {
                            if (count == 0)
                            {
                                package = new List<SmartSchool.Feature.Score.EditScore.UpdateInfo>(30);
                                count = 30;
                                if ((updatePackages.Count & 1) == 0)
                                    updatePackages.Add(package);
                                else
                                    updatePackages2.Add(package);
                            }
                            package.Add(var);
                            count--;
                        }
                    }
                    Thread threadUpdateSemesterSubjectScore = new Thread(new ParameterizedThreadStart(updateSemesterSubjectScore));
                    threadUpdateSemesterSubjectScore.IsBackground = true;
                    threadUpdateSemesterSubjectScore.Start(updatePackages);
                    Thread threadUpdateSemesterSubjectScore2 = new Thread(new ParameterizedThreadStart(updateSemesterSubjectScore));
                    threadUpdateSemesterSubjectScore2.IsBackground = true;
                    threadUpdateSemesterSubjectScore2.Start(updatePackages2);

                    threadUpdateSemesterSubjectScore.Join();
                    threadUpdateSemesterSubjectScore2.Join();
                    #endregion
                }
                if (insertList.Count > 0)
                {
                    //FISCA.LogAgent.ApplicationLog.Log("匯入學期科目成績", "匯入", insertLogStringBuilder.ToString());
                    #region 分批次兩路上傳

                    List<List<SmartSchool.Feature.Score.AddScore.InsertInfo>> insertPackages = new List<List<SmartSchool.Feature.Score.AddScore.InsertInfo>>();
                    List<List<SmartSchool.Feature.Score.AddScore.InsertInfo>> insertPackages2 = new List<List<SmartSchool.Feature.Score.AddScore.InsertInfo>>();
                    {
                        List<SmartSchool.Feature.Score.AddScore.InsertInfo> package = null;
                        int count = 0;
                        foreach (SmartSchool.Feature.Score.AddScore.InsertInfo var in insertList)
                        {
                            if (count == 0)
                            {
                                package = new List<SmartSchool.Feature.Score.AddScore.InsertInfo>(30);
                                count = 30;
                                if ((insertPackages.Count & 1) == 0)
                                    insertPackages.Add(package);
                                else
                                    insertPackages2.Add(package);
                            }
                            package.Add(var);
                            count--;
                        }
                    }
                    Thread threadInsertSemesterSubjectScore = new Thread(new ParameterizedThreadStart(insertSemesterSubjectScore));
                    threadInsertSemesterSubjectScore.IsBackground = true;
                    threadInsertSemesterSubjectScore.Start(insertPackages);
                    Thread threadInsertSemesterSubjectScore2 = new Thread(new ParameterizedThreadStart(insertSemesterSubjectScore));
                    threadInsertSemesterSubjectScore2.IsBackground = true;
                    threadInsertSemesterSubjectScore2.Start(insertPackages2);

                    threadInsertSemesterSubjectScore.Join();
                    threadInsertSemesterSubjectScore2.Join();
                    #endregion
                }
                #endregion
            };
            wizard.ImportComplete += delegate
            {
                EventHub.Instance.InvokScoreChanged(new List<string>(_StudentCollection.Keys).ToArray());
            };
        }

        private string GetRowDataCellValue(RowData rowData,string name)
        {
            return rowData.ContainsKey(name) ? rowData[name].ToString() : "";
        }


        private void updateSemesterSubjectScore(object item)
        {
            List<List<SmartSchool.Feature.Score.EditScore.UpdateInfo>> updatePackages = (List<List<SmartSchool.Feature.Score.EditScore.UpdateInfo>>)item;
            foreach (List<SmartSchool.Feature.Score.EditScore.UpdateInfo> package in updatePackages)
            {
                SmartSchool.Feature.Score.EditScore.UpdateSemesterSubjectScore(package.ToArray());
            }
        }

        private void insertSemesterSubjectScore(object item)
        {
            List<List<SmartSchool.Feature.Score.AddScore.InsertInfo>> insertPackages = (List<List<SmartSchool.Feature.Score.AddScore.InsertInfo>>)item;
            foreach (List<SmartSchool.Feature.Score.AddScore.InsertInfo> package in insertPackages)
            {
                SmartSchool.Feature.Score.AddScore.InsertSemesterSubjectScore(package.ToArray());
            }
        }
    }
}
