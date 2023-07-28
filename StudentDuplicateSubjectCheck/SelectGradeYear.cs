using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using FISCA.Data;
using SHSchool.Data;
using SmartSchool.Customization.Data;
using System.Xml.Linq;
using System.Xml;
using FISCA.DSAUtil;
using SmartSchool.Evaluation.WearyDogComputerHelper;
using StudentDuplicateSubjectCheck.DAO;

namespace StudentDuplicateSubjectCheck
{
    public partial class SelectGradeYear : BaseForm
    {
        string selectGradeYear;
        string targetGradeYear;
        string schoolYear = School.DefaultSchoolYear;
        string semester = School.DefaultSemester;
        private BackgroundWorker _backgroundWorker;

        string errorMessage = "";
        DataTable dtErrorTable = new DataTable();
        List<DataRow> hasScoreList = new List<DataRow>();
        List<DataRow> hasSubjectCodeList = new List<DataRow>();

        // 指定學年科目名稱
        Dictionary<string, CourseSpecifySubjectNameInfo> hasSpecifySubjectNameDict = new Dictionary<string, CourseSpecifySubjectNameInfo>();

        // 課程規劃表科目學分數上下學期不同
        Dictionary<string, List<GPlanSubjectCreditDifInfo>> hasDiffSubjectCreditDict = new Dictionary<string, List<GPlanSubjectCreditDifInfo>>();

        // 課程規劃表科目學分數上下學期不同
        Dictionary<string, List<GPlanSubjectCreditDifInfo>> SubejctCreditDifCheckData = new Dictionary<string, List<GPlanSubjectCreditDifInfo>>();

        Dictionary<string, List<GPlanSubjectCreditDifInfo>> SubejctCreditDifDataList = new Dictionary<string, List<GPlanSubjectCreditDifInfo>>();

        DataTable dtSubjectCode = new DataTable();

        List<SCAttendRecord> scaDuplicateList = new List<SCAttendRecord>();

        bool checkStandSocrePass = false;
        Dictionary<string, List<string>> dataCompareDict = new Dictionary<string, List<string>>(); // <sid,<科目名稱 + _ + 級別>>

        public SelectGradeYear()
        {
            InitializeComponent();


            this.numericUpDown1.Minimum = 1;
            this.numericUpDown1.Maximum = 5;
            this.numericUpDown1.Value = 1;

            labelX2.Text = "本功能將會對本學期學生的修課紀錄有3項檢查：\n" +
                "1.檢查成績計算規則，並將及格標準、補考標準、學生身分寫入修課紀錄。\n" +
                //"2.檢查課程規劃，並將課程代碼寫入修課紀錄。" + Environment.NewLine +
                "2.及格標準、補考標準只會寫入一次並不會覆蓋已寫入的紀錄。" + Environment.NewLine +
                "3.假若與先前的學期成績有重覆的科目級別則會列出，由人工設定計算方式。";

            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.DoWork += new DoWorkEventHandler(_backgroundWorker_DoWork);
            _backgroundWorker.ProgressChanged += new ProgressChangedEventHandler(_backgroundWorker_ProgressChanged);
            _backgroundWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_backgroundWorker_RunWorkerCompleted);
            _backgroundWorker.WorkerReportsProgress = true;

        }

        private void buttonX1_Click(object sender, EventArgs e)
        {

            targetGradeYear = "" + this.numericUpDown1.Value;
            FreezeUI();
            _backgroundWorker.RunWorkerAsync();
        }

        private void _backgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            ActivateUI();

            if (e.Cancelled)
            {
                if (errorMessage != "")
                {
                    ErrorMessageForm emf = new ErrorMessageForm();
                    emf.SetDataTable(dtErrorTable);
                    emf.SetMessage(errorMessage);
                    emf.ShowDialog();
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("檢查發生錯誤。");
                }
            }
            else
            {
                if (e.Error != null)
                {
                    MessageBox.Show("Error:" + e.Error.Message);
                }
                else
                {

                    FISCA.Presentation.MotherForm.SetStatusBarMessage("檢查本學期重覆修課完畢。");

                    // 有資料才show
                    if (scaDuplicateList.Count != 0)
                    {
                        CheckCalculatedLogicForm cclf = new CheckCalculatedLogicForm(scaDuplicateList);
                        cclf.ShowDialog();
                    }
                    else
                    {
                        MsgBox.Show("本年級本學期並無重覆修課的紀錄。");
                    }
                }

            }
        }

        private void _backgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (e.ProgressPercentage == 10)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("取得基本資料中...", e.ProgressPercentage);

            }
            else if (e.ProgressPercentage == 20)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("取得修課資料中...", e.ProgressPercentage);
            }
            else if (e.ProgressPercentage == 40)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("取得成績計算規則資料中...", e.ProgressPercentage);
            }
            else if (e.ProgressPercentage == 50)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("寫入及格標準、補考標準、學生身分中...", e.ProgressPercentage);
            }
            else if (e.ProgressPercentage == 60)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("取得課程規劃資料中...", e.ProgressPercentage);

            }
            else if (e.ProgressPercentage == 65)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("檢查課程規劃資料中...", e.ProgressPercentage);

            }
            else if (e.ProgressPercentage == 70)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("取得課程代碼中...", e.ProgressPercentage);
            }
            else if (e.ProgressPercentage == 80)
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("取得指定學年科目名稱中...", e.ProgressPercentage);
            }

            else if (e.ProgressPercentage == 90)
            { FISCA.Presentation.MotherForm.SetStatusBarMessage("取得重覆級別科目中...", e.ProgressPercentage); }
            else
            {
                FISCA.Presentation.MotherForm.SetStatusBarMessage("資料處理中...", e.ProgressPercentage);
            }


            MyUserState mus = e.UserState as MyUserState;
            if (mus != null)
            {
                if (mus.Name == "及格補考標準")
                {
                    if (hasScoreList.Count > 0)
                    {
                        HasPassScoreForm hpsf = new HasPassScoreForm();
                        hpsf.StartPosition = FormStartPosition.CenterScreen;
                        hpsf.SetDataRows(hasScoreList);
                        hpsf.ShowDialog();
                    }
                }

                if (mus.Name == "課程代碼")
                {

                    bool isUpdateNull = false;
                    bool isUpdate = false;
                    if (hasSubjectCodeList.Count > 0)
                    {
                        hasSubjectCodeForm hscf = new hasSubjectCodeForm();
                        hscf.StartPosition = FormStartPosition.CenterScreen;
                        hscf.SetDataRows(hasSubjectCodeList);
                        // 確定覆蓋
                        if (hscf.ShowDialog() == DialogResult.Yes)
                        {
                            isUpdate = true;
                        }
                        else
                        {
                            // 當選離開時，只更新有空值得課程代碼。
                            isUpdateNull = true;
                        }
                    }
                    else
                    {
                        isUpdate = true;
                    }

                    List<string> updateScSubjCodeList = new List<string>();

                    if (isUpdate)
                    {
                        // 回寫課程規劃課程代碼到修課紀錄上科目代碼                        
                        foreach (DataRow dr in dtSubjectCode.Rows)
                        {
                            string sc_id = dr["sc_attend_id"].ToString();
                            string subj_code = "";
                            if (dr["subject_code"] != null)
                            {
                                subj_code = dr["subject_code"].ToString();
                            }

                            string updateStr = "UPDATE " +
                                "sc_attend " +
                                "SET subject_code = '" + subj_code + "' " +
                                "WHERE " +
                                "id =" + sc_id + ";";
                            updateScSubjCodeList.Add(updateStr);
                        }
                    }

                    if (isUpdateNull)
                    {
                        // 回寫課程規劃課程代碼到修課紀錄上科目代碼空值                        
                        foreach (DataRow dr in dtSubjectCode.Rows)
                        {
                            bool add = true;
                            string sc_id = dr["sc_attend_id"].ToString();
                            string subj_code = "";
                            if (dr["subject_code"] != null)
                            {
                                subj_code = dr["subject_code"].ToString();
                            }

                            // 原資料
                            if (dr["subject_code_old"] != null)
                            {
                                if (!string.IsNullOrWhiteSpace(dr["subject_code_old"].ToString()) || string.IsNullOrWhiteSpace(subj_code))
                                    add = false;
                            }


                            if (add)
                            {
                                string updateStr = "UPDATE " +
                               "sc_attend " +
                               "SET subject_code = '" + subj_code + "' " +
                               "WHERE " +
                               "id =" + sc_id + ";";
                                updateScSubjCodeList.Add(updateStr);
                            }
                        }
                    }

                    if (updateScSubjCodeList.Count > 0)
                    {
                        try
                        {
                            UpdateHelper uhSubj = new UpdateHelper();
                            uhSubj.Execute(updateScSubjCodeList);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("覆蓋課程代碼發生錯誤" + ex.Message);

                            //Console.WriteLine(ex.Message);
                        }
                    }

                }


                if (mus.Name == "檢查課程規劃表")
                {
                    if (SubejctCreditDifDataList.Count > 0)
                    {
                        hasSubjectCreditDiffForm hf = new hasSubjectCreditDiffForm();
                        hf.StartPosition = FormStartPosition.CenterScreen;
                        hf.SetData(SubejctCreditDifDataList);
                        hf.ShowDialog();
                    }
                }

                if (mus.Name == "指定學年科目名稱")
                {

                    bool isUpdateNull = false;
                    bool isUpdate = false;
                    if (hasSpecifySubjectNameDict.Count > 0)
                    {
                        hasSpecifySubjectNameForm hscf = new hasSpecifySubjectNameForm();
                        hscf.StartPosition = FormStartPosition.CenterScreen;
                        hscf.SetData(hasSpecifySubjectNameDict);
                        // 確定覆蓋
                        if (hscf.ShowDialog() == DialogResult.Yes)
                        {
                            isUpdate = true;
                        }
                        else
                        {
                            // 當選離開時，只更新有空值得課程代碼。
                            isUpdateNull = true;
                        }
                    }
                    else
                    {
                        isUpdate = true;
                    }

                    List<string> updateCourseSpecifySubjectNameList = new List<string>();

                    if (isUpdate)
                    {
                        // 整理需要回寫課程指定學年科目名稱資料
                        foreach (CourseSpecifySubjectNameInfo ci in hasSpecifySubjectNameDict.Values)
                        {
                            string updateStr = "UPDATE " +
                                "course " +
                                "SET specify_subject_name = '" + ci.SpecifySubjectName + "' " +
                                "WHERE " +
                                "id =" + ci.CourseID + ";";
                            updateCourseSpecifySubjectNameList.Add(updateStr);
                        }
                    }

                    if (isUpdateNull)
                    {
                        // 整理需要回寫課程指定學年科目名稱資料，只寫入原本是空白                        
                        foreach (CourseSpecifySubjectNameInfo ci in hasSpecifySubjectNameDict.Values)
                        {
                            // 舊資料室空白
                            if (ci.SpecifySubjectNameOld == "" && ci.SpecifySubjectName != "")
                            {
                                string updateStr = "UPDATE " +
                             "course " +
                             "SET specify_subject_name = '" + ci.SpecifySubjectName + "' " +
                             "WHERE " +
                             "id =" + ci.CourseID + ";";
                                updateCourseSpecifySubjectNameList.Add(updateStr);
                            }
                        }
                    }

                    if (updateCourseSpecifySubjectNameList.Count > 0)
                    {
                        try
                        {
                            UpdateHelper uhSubj = new UpdateHelper();
                            uhSubj.Execute(updateCourseSpecifySubjectNameList);
                        }
                        catch (Exception ex)
                        {
                            MessageBox.Show("覆蓋指定學年科目名稱發生錯誤" + ex.Message);

                            //Console.WriteLine(ex.Message);
                        }
                    }

                }
            }
        }

        private void _backgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            this.errorMessage = "";
            this.dtErrorTable.Rows.Clear();
            _backgroundWorker.ReportProgress(10);

            //清空上次資料
            scaDuplicateList.Clear();



            List<string> sidList = new List<string>();

            // 學生班級基本資料(學生狀態為一般)
            QueryHelper qh1 = new QueryHelper();
            string query1 = "select student.id,class.grade_year,student_number,class.class_name,seat_no,student.name from student inner join class on student.ref_class_id=class.id where student.status=1 and class.grade_year ='" + targetGradeYear + "' order by class.grade_year,student_number;";
            DataTable dt1 = qh1.Select(query1);
            foreach (DataRow dr in dt1.Rows)
            {
                sidList.Add("" + dr["id"]);
            }

            _backgroundWorker.ReportProgress(20);

            #region 舊的抓取成績方式 (有問題!)
            // 2018/7/4 穎驊協助 嘉詮處理客服 https://ischool.zendesk.com/agent/tickets/6133 ，發現 ischool API  SHSemesterScore.SelectByStudentID 
            // 預設會將 學期歷程重覆的 重讀學期資料濾掉， 需要設定為 false，才會有。
            //List<SHSchool.Data.SHSemesterScoreRecord> ssList = SHSchool.Data.SHSemesterScore.SelectByStudentIDs(sidList, false);

            //foreach (SHSemesterScoreRecord data in ssList)
            //{
            //    // 2018/6/12 穎驊筆記，佳樺測出 會抓到同學期的成績比對，在此濾掉， 同學期的重覆不算，因為學校可能會先算過本學期的 成績， 那本學期成績 與本學期修課重覆 是很正常的事
            //    if ("" + data.SchoolYear == schoolYear && "" + data.Semester == semester)
            //    {
            //        continue;
            //    }
            //    // 以前學期，學期科目成績紀錄(題外話 如果未來覺得 scaList 使用ischool API的抓法太沒效率，可以考慮直接寫一串SQL)
            //    if (!dataCompareDict.ContainsKey(data.RefStudentID))
            //    {
            //        dataCompareDict.Add(data.RefStudentID, new List<string>());

            //        foreach (var subjectData in data.Subjects.Values)
            //        {
            //            if (!dataCompareDict[data.RefStudentID].Contains(subjectData.Subject + "_" + subjectData.Level))
            //            {
            //                dataCompareDict[data.RefStudentID].Add(subjectData.Subject + "_" + subjectData.Level);
            //            }
            //        }
            //    }
            //    else
            //    {
            //        foreach (var subjectData in data.Subjects.Values)
            //        {
            //            if (!dataCompareDict[data.RefStudentID].Contains(subjectData.Subject + "_" + subjectData.Level))
            //            {
            //                dataCompareDict[data.RefStudentID].Add(subjectData.Subject + "_" + subjectData.Level);
            //            }
            //        }
            //    }
            //} 
            #endregion

            // 2018/9/12 穎驊註解， 因應康橋結業計算詢問的問題:"有學生有重覆科目級別，卻在本功能無法列出"，
            // 檢查後發現 使用 ischool API List<SHSchool.Data.SHSemesterScoreRecord> ssList = SHSchool.Data.SHSemesterScore.SelectByStudentIDs(sidList, false);
            // 抓取的學期科目成績， 若同一學期 有重覆的兩個科目， 此API 回傳的只會有第一筆，因此會造成部分學生的重覆修課檢查無法在本功能被列出
            // 跟恩正討論後，建議與成績計算邏輯 WearyDogComputer 抓取學生學期科目成績使用方法一致，因此做了下列改寫

            #region 新寫法
            AccessHelper accesshelper = new AccessHelper();

            List<SmartSchool.Customization.Data.StudentRecord> students = new List<SmartSchool.Customization.Data.StudentRecord>();

            students = new List<SmartSchool.Customization.Data.StudentRecord>();
            foreach (SmartSchool.Customization.Data.ClassRecord classrecord in accesshelper.ClassHelper.GetAllClass())
            {
                int tryParseGradeYear;
                if (int.TryParse(classrecord.GradeYear, out tryParseGradeYear) && tryParseGradeYear == (int)numericUpDown1.Value)
                    students.AddRange(classrecord.Students);
            }

            // 預設會將 學期歷程重覆的 重讀學期資料濾掉， 需要設定為 false，才會有。
            accesshelper.StudentHelper.FillSemesterSubjectScore(false, students);
            _backgroundWorker.ReportProgress(30);



            _backgroundWorker.ReportProgress(40);

            Dictionary<string, decimal> passingStandardDict = new Dictionary<string, decimal>();
            Dictionary<string, decimal> makeupStandardDict = new Dictionary<string, decimal>();
            Dictionary<string, string> remarkDict = new Dictionary<string, string>();
            Dictionary<string, string> studentGraduationPlanDict = new Dictionary<string, string>();
            //Dictionary<string, string> studentGroupCodeDict = new Dictionary<string, string>();
            dtErrorTable.Columns.Clear();
            dtErrorTable.Columns.Add("student_number");
            dtErrorTable.Columns.Add("class_name");
            dtErrorTable.Columns.Add("seat_no");
            dtErrorTable.Columns.Add("student_name");
            dtErrorTable.Columns.Add("student_id");

            List<string> studentIDList = new List<string>();
            bool chkCalcHasError = false;
            foreach (SmartSchool.Customization.Data.StudentRecord student in students)
            {
                studentIDList.Add(student.StudentID);
                XmlElement scoreCalcRule = SmartSchool.Evaluation.ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(student.StudentID) == null ? null : SmartSchool.Evaluation.ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(student.StudentID).ScoreCalcRuleElement;
                if (scoreCalcRule == null)
                {
                    chkCalcHasError = true;
                    // 沒有設定成績計算規則
                    if (!remarkDict.ContainsKey(student.StudentID))
                        remarkDict.Add(student.StudentID, "沒有設定成績計算規則");

                    DataRow dr = this.dtErrorTable.NewRow();

                    dr["student_number"] = student.StudentNumber;
                    dr["class_name"] = student.RefClass.ClassName;
                    dr["seat_no"] = student.SeatNo;
                    dr["student_name"] = student.StudentName;
                    dr["student_id"] = student.StudentID;
                    this.dtErrorTable.Rows.Add(dr);
                }
                else
                {

                    DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                    decimal passingStandard = -1;
                    decimal makeupStandard = -1;

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

                        if (useful)
                        {
                            if (!remarkDict.ContainsKey(student.StudentID))
                            {
                                remarkDict.Add(student.StudentID, "" + cat);
                            }
                            else
                            {
                                remarkDict[student.StudentID] += "," + cat;
                            }
                        }

                        //學生是指定的類別或類別為"預設"
                        if (cat == "預設" || useful)
                        {
                            switch (targetGradeYear)
                            {
                                case "1":
                                    decimal.TryParse(element.GetAttribute("一年級及格標準"), out passingStandard);
                                    decimal.TryParse(element.GetAttribute("一年級補考標準"), out makeupStandard);
                                    break;
                                case "2":
                                    decimal.TryParse(element.GetAttribute("二年級及格標準"), out passingStandard);
                                    decimal.TryParse(element.GetAttribute("二年級補考標準"), out makeupStandard);
                                    break;
                                case "3":
                                    decimal.TryParse(element.GetAttribute("三年級及格標準"), out passingStandard);
                                    decimal.TryParse(element.GetAttribute("三年級補考標準"), out makeupStandard);
                                    break;
                                case "4":
                                    decimal.TryParse(element.GetAttribute("四年級及格標準"), out passingStandard);
                                    decimal.TryParse(element.GetAttribute("四年級補考標準"), out makeupStandard);
                                    break;
                                default:
                                    break;
                            }

                            if (passingStandardDict.ContainsKey(student.StudentID))
                            {
                                if (passingStandardDict[student.StudentID] > passingStandard)
                                    passingStandardDict[student.StudentID] = passingStandard;
                            }
                            else
                            {
                                passingStandardDict.Add(student.StudentID, passingStandard);
                            }

                            if (makeupStandardDict.ContainsKey(student.StudentID))
                            {
                                if (makeupStandardDict[student.StudentID] > makeupStandard)
                                {
                                    makeupStandardDict[student.StudentID] = makeupStandard;
                                }
                            }
                            else
                            {
                                makeupStandardDict.Add(student.StudentID, makeupStandard);
                            }
                        }
                    }
                }


                foreach (SmartSchool.Customization.Data.StudentExtension.SemesterSubjectScoreInfo subjectScore in student.SemesterSubjectScoreList)
                {

                    // 2018/6/12 穎驊筆記，佳樺測出 會抓到同學期的成績比對，在此濾掉， 同學期的重覆不算，因為學校可能會先算過本學期的 成績， 那本學期成績 與本學期修課重覆 是很正常的事
                    if ("" + subjectScore.SchoolYear == schoolYear && "" + subjectScore.Semester == semester)
                    {
                        continue;
                    }
                    // 以前學期，學期科目成績紀錄(題外話 如果未來覺得 scaList 使用ischool API的抓法太沒效率，可以考慮直接寫一串SQL)
                    if (!dataCompareDict.ContainsKey(student.StudentID))
                    {
                        dataCompareDict.Add(student.StudentID, new List<string>());

                        if (!dataCompareDict[student.StudentID].Contains(subjectScore.Subject + "_" + subjectScore.Level))
                        {
                            dataCompareDict[student.StudentID].Add(subjectScore.Subject + "_" + subjectScore.Level);
                        }
                    }
                    else
                    {
                        if (!dataCompareDict[student.StudentID].Contains(subjectScore.Subject + "_" + subjectScore.Level))
                        {
                            dataCompareDict[student.StudentID].Add(subjectScore.Subject + "_" + subjectScore.Level);
                        }
                    }
                }
            }
            #endregion


            if (chkCalcHasError)
            {
                e.Cancel = true;
                errorMessage = "沒有設定成績計算規則學生名單：";
                return;
            }

            // 取得學生當學期修課
            QueryHelper qhScAttend = new QueryHelper();
            string qryScAttend = "SELECT " +
                "sc_attend.id AS sc_attend_id" +
                ",course.id AS course_id" +
                ",course.school_year AS school_year" +
                ",course.semester AS semester" +
                ",course.course_name AS course_name" +
                ",student.id AS student_id" +
                ",student.name AS student_name" +
                ",student.student_number AS student_number" +
                ",sc_attend.passing_standard" +
                ",sc_attend.makeup_standard" +
                ",sc_attend.remark" +
                ",sc_attend.subject_code" +
                " FROM " +
                "course INNER JOIN sc_attend" +
                " ON course.id = sc_attend.ref_course_id INNER JOIN" +
                " student ON sc_attend.ref_student_id = student.id " +
                " WHERE student.id IN(" + string.Join(",", studentIDList.ToArray()) + ") AND course.school_year = " + schoolYear + " AND course.semester = " + semester + "" +
                " ORDER BY course.course_name,student.seat_no; ";

            DataTable dtScAttend = qh1.Select(qryScAttend);

            List<DataRow> updateScoreList = new List<DataRow>();
            hasScoreList.Clear();
            hasSubjectCodeList.Clear();
            // 比對填入 data table
            foreach (DataRow dr in dtScAttend.Rows)
            {
                string sid = dr["student_id"].ToString();
                bool hasScore = false;
                // 已有及格標準或補考標準不寫入
                // 及格標準
                if (dr["passing_standard"] != null)
                    if (dr["passing_standard"].ToString() != "")
                        hasScore = true;

                // 補考標準
                if (dr["makeup_standard"] != null)
                    if (dr["makeup_standard"].ToString() != "")
                        hasScore = true;


                // 是否有課程代碼
                bool hasSubjectCode = false;

                if (dr["subject_code"] != null)
                {
                    if (dr["subject_code"].ToString() != "")
                        hasSubjectCode = true;
                }

                if (hasSubjectCode)
                {
                    hasSubjectCodeList.Add(dr);
                }


                if (hasScore)
                {
                    hasScoreList.Add(dr);
                }
                else
                {
                    // 及格標準
                    if (passingStandardDict.ContainsKey(sid))
                    {
                        dr["passing_standard"] = passingStandardDict[sid];
                    }

                    // 補考標準
                    if (makeupStandardDict.ContainsKey(sid))
                    {
                        dr["makeup_standard"] = makeupStandardDict[sid];
                    }

                    // 備註
                    if (remarkDict.ContainsKey(sid))
                    {
                        dr["remark"] = "學生身分：" + remarkDict[sid];
                    }
                    updateScoreList.Add(dr);
                }

            }

            MyUserState mus = new MyUserState();
            mus.Name = "及格補考標準";
            mus.Value = "50";
            _backgroundWorker.ReportProgress(50, mus);
            // _backgroundWorker.ReportProgress(50);


            // 更新修課紀錄
            List<string> sbUpdateScAttend = new List<string>();
            // 只更新沒有及格補考標準
            foreach (DataRow dr in updateScoreList)
            {
                string passing_standard = "null";
                string makeup_standard = "null";
                string rem = "";

                if (dr["passing_standard"] != null)
                    passing_standard = dr["passing_standard"].ToString();

                if (dr["makeup_standard"] != null)
                    makeup_standard = dr["makeup_standard"].ToString();

                if (dr["remark"] != null)
                    rem = dr["remark"].ToString();

                string sc_id = dr["sc_attend_id"].ToString();
                string updateStr = "" +
                    "UPDATE sc_attend " +
                    " SET passing_standard=" + passing_standard +
                    ",makeup_standard =" + makeup_standard +
                    ",remark='" + rem + "'" +
                    " WHERE id=" + sc_id + ";";
                sbUpdateScAttend.Add(updateStr);
            }

            // 更新修課
            UpdateHelper upScattend = new UpdateHelper();
            try
            {
                upScattend.Execute(sbUpdateScAttend);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }


            _backgroundWorker.ReportProgress(60);

            // 處理課程代碼



            // 學生取得課程規劃 
            string qrygpPlan = @"
            SELECT 
	            student.id AS student_id,
	            COALESCE(student.ref_graduation_plan_id,class.ref_graduation_plan_id) AS graduation_plan_id
            FROM 
	            student 
            LEFT JOIN 
	            class 
		            ON student.ref_class_id = class.id 
	            WHERE 
            student.id IN(" + string.Join(",", studentIDList.ToArray()) + ");";

            QueryHelper qhGpPlan = new QueryHelper();
            DataTable dtGpPlan = qhGpPlan.Select(qrygpPlan);
            List<string> noGpPlanStudentIDList = new List<string>();
            foreach (DataRow dr in dtGpPlan.Rows)
            {
                string sid = dr["student_id"] + "";
                string gpID = dr["graduation_plan_id"] + "";


                if (gpID == "")
                    noGpPlanStudentIDList.Add(sid);

                if (!studentGraduationPlanDict.ContainsKey(sid))
                    studentGraduationPlanDict.Add(sid, gpID);
            }

            // 有學生沒有課程規劃表
            if (noGpPlanStudentIDList.Count > 0)
            {
                errorMessage = "沒有設定課程規劃表學生清單：";
                dtErrorTable.Rows.Clear();
                dtErrorTable.Columns.Clear();
                this.dtErrorTable.Columns.Add("student_number");
                this.dtErrorTable.Columns.Add("class_name");
                this.dtErrorTable.Columns.Add("seat_no");
                this.dtErrorTable.Columns.Add("student_name");
                this.dtErrorTable.Columns.Add("student_id");
                foreach (SmartSchool.Customization.Data.StudentRecord student in students)
                {
                    if (noGpPlanStudentIDList.Contains(student.StudentID))
                    {
                        DataRow dr = this.dtErrorTable.NewRow();
                        dr["student_number"] = student.StudentNumber;
                        dr["class_name"] = student.RefClass.ClassName;
                        dr["seat_no"] = student.SeatNo;
                        dr["student_name"] = student.StudentName;
                        dr["student_id"] = student.StudentID;
                        this.dtErrorTable.Rows.Add(dr);
                    }
                }
                e.Cancel = true;
                return;
            }

            // 分類有哪些課程規劃ID並取得資料
            List<string> gpIDList = new List<string>();
            foreach (string id in studentGraduationPlanDict.Values)
            {
                if (string.IsNullOrWhiteSpace(id))
                    continue;

                if (!gpIDList.Contains(id))
                    gpIDList.Add(id);
            }


            //取得課程規劃內容
            string qryGpPlanData = "SELECT id,name,content FROM graduation_plan WHERE id IN(" + string.Join(",", gpIDList.ToArray()) + ") ";

            DataTable dtGpPlanData = qhGpPlan.Select(qryGpPlanData);

            Dictionary<string, XElement> gpDataDict = new Dictionary<string, XElement>();
            // 課規id,name
            Dictionary<string, string> gpIDNameDict = new Dictionary<string, string>();
            foreach (DataRow dr in dtGpPlanData.Rows)
            {
                string id = dr["id"] + "";
                string name = dr["name"] + "";
                XElement elm = XElement.Parse(dr["content"].ToString());
                gpDataDict.Add(id, elm);
                if (!gpIDNameDict.ContainsKey(id))
                    gpIDNameDict.Add(id, name);
            }

            SubejctCreditDifCheckData.Clear();
            SubejctCreditDifDataList.Clear();

            // 檢查課程規劃表內科目同一年上下學期學分數不同，又沒有設定指定學年科目名稱
            foreach (string id in gpDataDict.Keys)
            {
                foreach (XElement elm in gpDataDict[id].Elements("Subject"))
                {
                    string key = id + "_" + elm.Attribute("GradeYear").Value + "_" + elm.Attribute("SubjectName").Value;

                    if (!SubejctCreditDifCheckData.ContainsKey(key))
                        SubejctCreditDifCheckData.Add(key, new List<GPlanSubjectCreditDifInfo>());

                    GPlanSubjectCreditDifInfo gi = new GPlanSubjectCreditDifInfo();
                    gi.GPlanID = id;
                    if (gpIDNameDict.ContainsKey(id))
                        gi.GPlanName = gpIDNameDict[id];

                    gi.GradeYear = elm.Attribute("GradeYear").Value;
                    gi.SubjectName = elm.Attribute("SubjectName").Value;
                    gi.Credit = elm.Attribute("Credit").Value;
                    gi.Semester = elm.Attribute("Semester").Value;

                    gi.SpecifySubjectName = "";
                    if (elm.Attribute("指定學年科目名稱") != null)
                    {
                        gi.SpecifySubjectName = elm.Attribute("指定學年科目名稱").Value;
                    }


                    SubejctCreditDifCheckData[key].Add(gi);
                }
            }

            // 檢查有科目同年級學分數不同
            foreach (string key in SubejctCreditDifCheckData.Keys)
            {
                // 有上下學期才檢查
                if (SubejctCreditDifCheckData[key].Count == 2)
                {
                    // 指定學年科目名稱都空白才需要判斷
                    if(SubejctCreditDifCheckData[key][0].SpecifySubjectName == "" && SubejctCreditDifCheckData[key][1].SpecifySubjectName == "")
                    {
                        // 上下學期學分數不同
                        if (SubejctCreditDifCheckData[key][0].Credit != SubejctCreditDifCheckData[key][1].Credit)
                        {
                            if (!SubejctCreditDifDataList.ContainsKey(key))
                                SubejctCreditDifDataList.Add(key, new List<GPlanSubjectCreditDifInfo>());

                            SubejctCreditDifDataList[key].Add(SubejctCreditDifCheckData[key][0]);
                            SubejctCreditDifDataList[key].Add(SubejctCreditDifCheckData[key][1]);
                        }
                    }
                    
                }
            }

            Console.WriteLine(SubejctCreditDifDataList.Count);


            // 取得學生修課科目代碼
            string qryScAttendSubjectCode = "" +
                "SELECT " +
                "sc_attend.id AS sc_attend_id" +
                ",course.id AS course_id" +
                ",course.course_name" +
                ",course.school_year" +
                ",course.semester" +
                ",course.subject AS subject_name" +
                ",course.subj_level AS subj_level" +
                ",ref_student_id AS student_id" +
                ",sc_attend.subject_code " +
                ", sc_attend.subject_code AS subject_code_old " +
                ", CASE WHEN course.c_required_by = 1 THEN '部定' ELSE '校訂' END AS required_by " +
                ", CASE WHEN course. c_is_required = '1' THEN '必修' ELSE '選修' END AS is_required " +
                ", course.score_type AS entry " +
                ", course.credit " +
                ", course.specify_subject_name " +
                "FROM " +
                "course " +
                "INNER JOIN sc_attend " +
                "ON course.id = sc_attend.ref_course_id " +
                "WHERE ref_student_id IN (" + string.Join(",", studentIDList.ToArray()) + ") " +
                "AND course.school_year=" + schoolYear + " AND course.semester=" + semester + " AND subject is not null";
            QueryHelper qhSubjectCode = new QueryHelper();
            dtSubjectCode = qhSubjectCode.Select(qryScAttendSubjectCode);

            hasSpecifySubjectNameDict.Clear();

            #region 課程規劃表
            foreach (DataRow dr in dtSubjectCode.Rows)
            {
                string sid = dr["student_id"] + "";
                string coid = dr["course_id"] + "";

                if (studentGraduationPlanDict.ContainsKey(sid))
                {
                    string gpid = studentGraduationPlanDict[sid];
                    if (gpDataDict.ContainsKey(gpid))
                    {
                        foreach (XElement elm in gpDataDict[gpid].Elements("Subject"))
                        {
                            string subjName = "", subjLevel = "", subjCode = "";
                            if (elm.Attribute("SubjectName") != null)
                                subjName = elm.Attribute("SubjectName").Value;
                            if (elm.Attribute("Level") != null)
                                subjLevel = elm.Attribute("Level").Value;

                            if (elm.Attribute("課程代碼") != null)
                                subjCode = elm.Attribute("課程代碼").Value;

                            // 科目名稱+科目級別 相同
                            if (dr["subject_name"] + "" == subjName && dr["subj_level"] + "" == subjLevel)
                            {
                                dr["subject_code"] = subjCode;

                                // 處理指定學年科目名稱
                                if (!hasSpecifySubjectNameDict.ContainsKey(coid))
                                {
                                    CourseSpecifySubjectNameInfo ci = new CourseSpecifySubjectNameInfo();
                                    ci.CourseID = coid;
                                    ci.CourseName = dr["course_name"] + "";
                                    ci.SchoolYear = dr["school_year"] + "";
                                    ci.Semester = dr["semester"] + "";
                                    ci.SubjectName = subjName;
                                    ci.SubjectLevel = subjLevel;
                                    ci.SpecifySubjectNameOld = dr["specify_subject_name"] + "";
                                    ci.SpecifySubjectName = "";

                                    if (elm.Attribute("指定學年科目名稱") != null)
                                    {
                                        ci.SpecifySubjectName = elm.Attribute("指定學年科目名稱").Value;
                                    }

                                    if (ci.SpecifySubjectName != "")
                                        hasSpecifySubjectNameDict.Add(coid, ci);
                                }
                            }

                        }
                    }
                }
            }
            #endregion

            MyUserState mus01 = new MyUserState();
            mus01.Name = "檢查課程規劃表";
            mus01.Value = "650";
            _backgroundWorker.ReportProgress(65, mus01);

            MyUserState mus1 = new MyUserState();
            mus1.Name = "課程代碼";
            mus1.Value = "70";
            _backgroundWorker.ReportProgress(70, mus1);


            //// 回寫課程規劃課程代碼到修課紀錄上科目代碼
            //List<string> updateScSubjCodeList = new List<string>();
            //foreach (DataRow dr in dtSubjectCode.Rows)
            //{
            //    string sc_id = dr["sc_attend_id"] + "";
            //    string subj_code = "";
            //    if (dr["subject_code"] != null)
            //    {
            //        subj_code = dr["subject_code"] + "";
            //    }

            //    string updateStr = "UPDATE " +
            //        "sc_attend " +
            //        "SET subject_code = '" + subj_code + "' " +
            //        "WHERE " +
            //        "id =" + sc_id + ";";
            //    updateScSubjCodeList.Add(updateStr);
            //}

            //if (updateScSubjCodeList.Count > 0)
            //{
            //    try
            //    {
            //        UpdateHelper uhSubj = new UpdateHelper();
            //        uhSubj.Execute(updateScSubjCodeList);
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine(ex.Message);
            //    }
            //}

            MyUserState mus2 = new MyUserState();
            mus2.Name = "指定學年科目名稱";
            mus2.Value = "80";
            _backgroundWorker.ReportProgress(80, mus2);



            _backgroundWorker.ReportProgress(90);

            #region 舊寫法
            // 2018/5/23 穎驊 筆記 使用ischool API 抓不到 scattend 中的 extensions 資料，故採用SQL 來抓
            //List<SHSchool.Data.SHSCAttendRecord> scaList = SHSchool.Data.SHSCAttend.SelectByStudentIDs(sidList);

            //// 是本學期的修課紀錄(題外話 如果未來覺得 scaList 使用ischool API的抓法太沒效率，可以考慮直接寫一串SQL)
            //foreach (SHSCAttendRecord data in scaList)
            //{
            //    // 本學期的修課紀錄(題外話 如果未來覺得 scaList 使用ischool API的抓法太沒效率，可以考慮直接寫一串SQL)
            //    if (data.Course.SchoolYear == int.Parse(schoolYear) && data.Course.Semester == int.Parse(semester))
            //    {
            //        if (dataCompareDict.ContainsKey(data.RefStudentID))
            //        {
            //            if (dataCompareDict[data.RefStudentID].Contains(data.Course.Subject + "_" + data.Course.Level))
            //            {
            //                scaDuplicateList.Add(data); // 假如有找到重覆的科目級別資料，將 本學期修課紀錄加入List 處理
            //            }
            //        }
            //    }
            //} 
            #endregion

            //取得學生的修課紀錄、學生基本資料、課程名稱、級別

            string query2 = string.Format(@"SELECT sc_attend.id
	,sc_attend.extensions AS extensions
	,student.id AS refStudentID	
	,student.name AS studentName
	,student.student_number AS studentNumber
	,student.seat_no AS seatNo  
	,class.class_name AS className
	,class.grade_year AS gradeYear
	,course.id AS refCourseID
    ,course.course_name AS courseName
	,course.subject AS subjectName
	,course.subj_level AS subjectLevel
	FROM sc_attend 
	LEFT JOIN student ON sc_attend.ref_student_id =student.id 
	LEFT JOIN class ON student.ref_class_id =class.id  
	LEFT JOIN course ON sc_attend.ref_course_id =course.id  
	WHERE 
	student.status ='1' 
	AND course.school_year = '{0}'
    AND course.semester = '{1}'
    AND class.grade_year = '{2}'
    ORDER BY courseName,className, seatNo ASC", schoolYear, semester, targetGradeYear);

            DataTable dt_SCAttend = qh1.Select(query2);

            _backgroundWorker.ReportProgress(90);

            foreach (DataRow dr in dt_SCAttend.Rows)
            {
                if (dataCompareDict.ContainsKey("" + dr["refStudentID"]))
                {
                    if (dataCompareDict["" + dr["refStudentID"]].Contains(dr["subjectName"] + "_" + dr["subjectLevel"]))
                    {
                        SCAttendRecord scar = new SCAttendRecord();

                        scar.ID = "" + dr["id"];
                        scar.Extensions = "" + dr["extensions"];
                        scar.RefStudentID = "" + dr["refStudentID"];
                        scar.Name = "" + dr["studentName"];
                        scar.StudentNumber = "" + dr["studentNumber"];
                        scar.SeatNo = "" + dr["seatNo"];
                        scar.RefCourseID = "" + dr["refStudentID"];
                        scar.ClassName = "" + dr["className"];
                        scar.GradeYear = "" + dr["gradeYear"];
                        scar.RefCourseID = "" + dr["refCourseID"];
                        scar.CourseName = "" + dr["courseName"];
                        scar.SubjectName = "" + dr["subjectName"];
                        scar.SubjectLevel = "" + dr["subjectLevel"];

                        scaDuplicateList.Add(scar); // 假如有找到重覆的科目級別資料，將 本學期修課紀錄加入List 處理
                    }
                }
            }
            _backgroundWorker.ReportProgress(100);
        }


        //關閉
        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FreezeUI()
        {
            buttonX1.Enabled = false;
            buttonX2.Enabled = false;
        }

        private void ActivateUI()
        {
            buttonX1.Enabled = true;
            buttonX2.Enabled = true;
        }

        private void SelectGradeYear_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;
        }
    }
}

