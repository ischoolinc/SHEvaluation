using Aspose.Words;
using FISCA.Data;
using FISCA.DSAUtil;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using SmartSchool.Evaluation.Process.Wizards.LearningHistory;
using SmartSchool.Evaluation.WearyDogComputerHelper;
using SmartSchool.StudentRelated.RibbonBars.Export.RequestHandler.Formater;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Xml;
using System.Xml.Linq;

namespace SmartSchool.Evaluation
{
    /// <summary>
    /// 成績計算及畢業判斷檢查類別
    /// </summary>
    public class WearyDogComputer
    {
        /// <summary>
        /// 異常課程提示清單
        /// </summary>
        public List<string> _WarningList;

        /// <summary>
        /// 進位方式
        /// </summary>
        public enum RoundMode { 四捨五入, 無條件進位, 無條件捨去 }

        /// <summary>
        /// 根據進位條件取得成績
        /// </summary>
        /// <param name="score">成績</param>
        /// <param name="decimals">小數點</param>
        /// <param name="mode">進位模式</param>
        /// <returns></returns>
        internal static decimal GetRoundScore(decimal score, int decimals, RoundMode mode)
        {
            decimal seed = Convert.ToDecimal(Math.Pow(0.1, Convert.ToDouble(decimals)));
            switch (mode)
            {
                default:
                case RoundMode.四捨五入:
                    score = decimal.Round(score, decimals, MidpointRounding.AwayFromZero);
                    break;
                case RoundMode.無條件捨去:
                    score /= seed;
                    score = decimal.Floor(score);
                    score *= seed;
                    break;
                case RoundMode.無條件進位:
                    decimal d2 = GetRoundScore(score, decimals, RoundMode.無條件捨去);
                    if (d2 != score)
                        score = d2 + seed;
                    else
                        score = d2;
                    break;
            }
            string ss = "0.";
            for (int i = 0; i < decimals; i++)
            {
                ss += "0";
            }
            return Convert.ToDecimal(Math.Round(score, decimals).ToString(ss));
        }

        /// <summary>
        /// 計算學期科目成績
        /// </summary>
        /// <param name="schoolyear">學年度</param>
        /// <param name="semester">學期</param>
        /// <param name="accesshelper">資料存取類別</param>
        /// <param name="students">學生物件列表</param>
        /// <returns></returns>
        public Dictionary<StudentRecord, List<string>> FillSemesterSubjectCalcScore(int schoolyear, int semester, AccessHelper accesshelper, List<StudentRecord> students)
        {
            Dictionary<StudentRecord, List<string>> _ErrorList = new Dictionary<StudentRecord, List<string>>();
            //抓學生學期歷程
            accesshelper.StudentHelper.FillField("SemesterHistory", students);
            //抓學生學期修課紀錄
            accesshelper.StudentHelper.FillAttendCourse(schoolyear, semester, students);
            //抓學生歷年學期科目成績
            accesshelper.StudentHelper.FillSemesterSubjectScore(false, students);

            // 2018/5/24 穎驊新增 紀錄 於 [H成績][06] 課程重讀重修設定，的新設定，假如學生有重覆科目級別被抓出來，且設定了計算方式，
            // 此字典可用於對照， 格式 <studentID_subject_level,重修(寫回原學期)||重讀(擇優採計成績)||視為一般修課>
            Dictionary<string, string> duplicateSubjectLevelMethodDict = new Dictionary<string, string>();

            Dictionary<string, string> duplicateSubjectLevelMethodDict_Afterfilter = new Dictionary<string, string>(); // 真正過濾後，有重覆科目級別的項目


            // 整理所欲計算學期科目學生的ID
            List<string> sidList = new List<string>();

            foreach (StudentRecord sr in students)
            {
                if (!sidList.Contains(sr.StudentID))
                {
                    sidList.Add(sr.StudentID);
                }
            }

            string sid = string.Join(",", sidList);

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
    , course.specify_subject_name
	, COALESCE(sc_attend.required_by, course.c_required_by)  AS required_by
	, COALESCE(sc_attend.is_required, course.c_is_required)  AS is_required
    , course.score_type
	FROM sc_attend 
	LEFT JOIN student ON sc_attend.ref_student_id =student.id 
	LEFT JOIN class ON student.ref_class_id =class.id  
	LEFT JOIN course ON sc_attend.ref_course_id =course.id  
	WHERE 
	student.status ='1' 
	AND course.school_year = '{0}'
    AND course.semester = '{1}'
    AND student.id IN ({2})
    ORDER BY courseName,className, seatNo ASC", schoolyear, semester, sid);

            QueryHelper qh1 = new QueryHelper();

            DataTable dt_SCAttend = qh1.Select(query2);

            // 2019/8/23 by CT，取得學生有直接指定總成績，如果有將取代原始成績,key:student_id,subject+subjectLevel
            Dictionary<string, Dictionary<string, DataRow>> studentFinalScoreDict = new Dictionary<string, Dictionary<string, DataRow>>();

            // 2022/5/12 by Cynthia，取得課程上的指定學年科目名稱，存入「指定學年科目名稱」屬性,key:student_id,  subject+subjectLevel
            Dictionary<string, Dictionary<string, string>> specifySubjectNameDict = new Dictionary<string, Dictionary<string, string>>();

            // 2022/6/2 by Cynthia  檢查設定必選修。key:student_id,  course_name
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> courseRequied = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            // 各科及格標準
            Dictionary<string, Dictionary<string, decimal>> studentPassScoreDict = new Dictionary<string, Dictionary<string, decimal>>();

            // 學生有指定總成績
            Dictionary<string, Dictionary<string, decimal>> studentHasFinalScoreDict = new Dictionary<string, Dictionary<string, decimal>>();

            string strFinalScore = string.Format(@"
            SELECT
                sc_attend.id AS sc_attend_id,
                student.id AS student_id,
                course.subject,
                course.subj_level,
                sc_attend.passing_standard,
                sc_attend.makeup_standard,
                sc_attend.remark,
                designate_final_score,
                sc_attend.subject_code AS subject_code,
                course.course_name
            FROM
                sc_attend
                INNER JOIN course ON sc_attend.ref_course_id = course.id
                INNER JOIN student ON sc_attend.ref_student_id = student.id
            WHERE
                course.school_year = {0} 
                AND course.semester = {1} 
                AND student.id IN({2});
            ", schoolyear, semester, sid);

            DataTable dtFinalScore = qh1.Select(strFinalScore);

            foreach (DataRow dr in dtFinalScore.Rows)
            {
                string student_id = dr["student_id"].ToString();
                string level = "";
                if (dr["subj_level"] != null)
                    level = dr["subj_level"].ToString().Trim();
                string key = dr["subject"].ToString().Trim() + "_" + level;
                if (!studentFinalScoreDict.ContainsKey(student_id))
                    studentFinalScoreDict.Add(student_id, new Dictionary<string, DataRow>());

                if (!studentFinalScoreDict[student_id].ContainsKey(key))
                    studentFinalScoreDict[student_id].Add(key, dr);
                else
                    studentFinalScoreDict[student_id][key] = dr;


                // 放入及格標準
                if (!studentPassScoreDict.ContainsKey(student_id))
                    studentPassScoreDict.Add(student_id, new Dictionary<string, decimal>());

                // 放指定總成績
                if (!studentHasFinalScoreDict.ContainsKey(student_id))
                    studentHasFinalScoreDict.Add(student_id, new Dictionary<string, decimal>());

                decimal passScore;


                if (!studentPassScoreDict[student_id].ContainsKey(key))
                {
                    studentPassScoreDict[student_id].Add(key, -1);
                }


                if (dr["passing_standard"] != null)
                {
                    if (decimal.TryParse(dr["passing_standard"].ToString(), out passScore))
                    {
                        studentPassScoreDict[student_id][key] = passScore;
                    }
                }

                if (dr["designate_final_score"] != null)
                {
                    decimal finalScore;
                    if (decimal.TryParse(dr["designate_final_score"].ToString(), out finalScore))
                    {
                        if (!studentHasFinalScoreDict[student_id].ContainsKey(key))
                            studentHasFinalScoreDict[student_id].Add(key, finalScore);
                    }
                }

            }


            // 將學生本學期的修課紀錄整理出來 (包含了重覆科目級別的計算處理方式會在extensions內)
            foreach (DataRow dr in dt_SCAttend.Rows)
            {
                string key = "" + dr["refStudentID"] + "_" + dr["subjectName"] + "_" + dr["subjectLevel"];

                string xmlStr = "<root>" + dr["extensions"] + "</root>";
                string method = "";

                XElement elmRoot = XElement.Parse(xmlStr);

                if (elmRoot != null)
                {
                    if (elmRoot.Element("Extensions") != null)
                    {
                        foreach (XElement ex in elmRoot.Element("Extensions").Elements("Extension"))
                        {
                            if (ex.Attribute("Name").Value == "DuplicatedLevelSubjectCalRule")
                            {
                                method = ex.Element("Rule").Value;
                            }
                        }
                    }
                }
                if (!duplicateSubjectLevelMethodDict.ContainsKey(key))
                {
                    duplicateSubjectLevelMethodDict.Add("" + dr["refStudentID"] + "_" + dr["subjectName"] + "_" + dr["subjectLevel"], method);
                }

                #region 指定學年科目名稱
                string student_id = dr["refStudentID"].ToString();
                string level = "";
                if (dr["subjectLevel"] != null)
                    level = dr["subjectLevel"].ToString().Trim();

                string subjectName = dr["subjectName"].ToString().Trim();
                string subjectKey = subjectName + "_" + level;
                string specifySubjectName = dr["specify_subject_name"].ToString().Trim();
                string courseName = dr["courseName"].ToString();
                string required_by = dr["required_by"].ToString();
                string is_required = dr["is_required"].ToString();
                string score_type = dr["score_type"].ToString();
                if (!specifySubjectNameDict.ContainsKey(student_id))
                    specifySubjectNameDict.Add(student_id, new Dictionary<string, string>());

                if (!specifySubjectNameDict[student_id].ContainsKey(subjectKey))
                    specifySubjectNameDict[student_id].Add(subjectKey, specifySubjectName);
                else
                    specifySubjectNameDict[student_id][subjectKey] = specifySubjectName;

                // 2024/5/27 CT,調整當有科目名稱才檢查，這三項
                if (!string.IsNullOrEmpty(subjectName))
                {
                    #region 校部定/必選修/分項類別
                    if (!courseRequied.ContainsKey(student_id))
                        courseRequied.Add(student_id, new Dictionary<string, Dictionary<string, string>>());
                    if (!courseRequied[student_id].ContainsKey(courseName))
                        courseRequied[student_id].Add(courseName, new Dictionary<string, string>());
                    if (!courseRequied[student_id][courseName].ContainsKey("required_by"))
                        courseRequied[student_id][courseName].Add("required_by", required_by);
                    if (!courseRequied[student_id][courseName].ContainsKey("is_required"))
                        courseRequied[student_id][courseName].Add("is_required", is_required);
                    if (!courseRequied[student_id][courseName].ContainsKey("score_type"))
                        courseRequied[student_id][courseName].Add("score_type", score_type);

                    #endregion
                }
                #endregion
            }

            // 透過學生系統編號，取得封存成績
            Dictionary<string, Dictionary<string, decimal>> dataCompareDict1 = new Dictionary<string, Dictionary<string, decimal>>();

            string query1 = string.Format(@"
            SELECT
                ext.ref_student_id,
                ext.grade_year,
                ext.semester,
                ext.school_year,
                array_to_string(xpath('//Subject/@科目', subj_score_ele), '')::text AS 科目,
                array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '')::text AS 科目級別,
                array_to_string(xpath('//Subject/@不計學分', subj_score_ele), '')::text AS 不計學分,
                array_to_string(xpath('//Subject/@不需評分', subj_score_ele), '')::text AS 不需評分,
                array_to_string(xpath('//Subject/@原始成績', subj_score_ele), '')::text AS 原始成績,
                array_to_string(xpath('//Subject/@學年調整成績', subj_score_ele), '')::text AS 學年調整成績,
                array_to_string(xpath('//Subject/@擇優採計成績', subj_score_ele), '')::text AS 擇優採計成績,
                array_to_string(xpath('//Subject/@補考成績', subj_score_ele), '')::text AS 補考成績,
                array_to_string(xpath('//Subject/@重修成績', subj_score_ele), '')::text AS 重修成績
            FROM (
                SELECT 
                    archive.*,
                    unnest(xpath('//SemesterSubjectScoreInfo/Subject', xmlparse(content score_info))) as subj_score_ele
                FROM $semester_subject_score_archive archive
                WHERE archive.uid IN (
                    SELECT uid FROM (
                        SELECT
                            uid,
                            ref_student_id,
                            school_year,
                            semester,
                            ROW_NUMBER() OVER (
                                PARTITION BY ref_student_id, school_year, semester
                                ORDER BY last_update DESC
                            ) AS rn
                        FROM $semester_subject_score_archive
                        WHERE ref_student_id IN ({0})
                    ) t
                    WHERE t.rn = 1
                )
            ) ext
            ORDER BY ext.ref_student_id, ext.school_year DESC, ext.semester DESC, ext.grade_year DESC

            ", sid);

            try
            {
                QueryHelper qhd1 = new QueryHelper();
                DataTable dtd1 = qhd1.Select(query1);
                foreach (DataRow dr in dtd1.Rows)
                {
                    string studentID = dr["ref_student_id"].ToString();
                    string subject = dr["科目"].ToString();
                    string level = dr["科目級別"].ToString();
                    string s1 = dr["不計學分"].ToString();
                    string s2 = dr["不需評分"].ToString();

                    if (s2 == "是")
                    {
                        // 不需評分的科目，跳過
                        continue;
                    }

                    //最高分
                    decimal maxScore = 0;// = sacRecord.FinalScore;


                    string[] scoreNames = new string[] { "原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績" };

                    foreach (string scorename in scoreNames)
                    {
                        decimal s;
                        if (dr[scorename] != null)
                            if (decimal.TryParse(dr[scorename].ToString(), out s))
                            {
                                if (s > maxScore)
                                {
                                    maxScore = s;
                                }
                            }
                    }


                    string key = subject + "_" + level;
                    // 只要有成績的就加入
                    if (!dataCompareDict1.ContainsKey(studentID))
                    {
                        dataCompareDict1.Add(studentID, new Dictionary<string, decimal>());
                    }

                    if (!dataCompareDict1[studentID].ContainsKey(key))
                    {
                        dataCompareDict1[studentID].Add(key, maxScore);
                    }

                }
            }
            catch (Exception ex)
            {
                FISCA.LogAgent.ApplicationLog.Log("成績計算", "取得封存成績異常", ex.Message);
            }

            // 補修成績使用
            List<SubjectScoreRec108> makeUpScoreList = new List<SubjectScoreRec108>();

            // 重修成績使用
            List<SubjectScoreRec108> restudyScoreList = new List<SubjectScoreRec108>();

            // 4.4 轉學轉科名冊
            List<SubjectScoreRec108> transferScoreList = new List<SubjectScoreRec108>();

            // 5.3 重讀成績名冊
            List<SubjectScoreRec108> repeatScoreList = new List<SubjectScoreRec108>();


            foreach (StudentRecord var in students)
            {
                //成績年級
                int? gradeYear = null;
                //精準位數
                int decimals = 2;
                //進位模式
                RoundMode mode = RoundMode.四捨五入;
                //使用擇優採計成績
                bool choseBetter = true;
                //重修登錄至原學期
                //bool writeToFirstSemester = false;
                //及格標準<年及,及格標準>
                Dictionary<int, decimal> applyLimit = new Dictionary<int, decimal>();
                //applyLimit.Add(1, 60);
                //applyLimit.Add(2, 60);
                //applyLimit.Add(3, 60);
                //applyLimit.Add(4, 60);
                //成績年級及計算規則皆存在，允許計算成績
                bool canCalc = true;
                #region 取得成績年級跟計算規則
                {
                    #region 處理成績年級
                    XmlElement semesterHistory = (XmlElement)var.Fields["SemesterHistory"];
                    if (semesterHistory == null)
                    {
                        if (!_ErrorList.ContainsKey(var))
                            _ErrorList.Add(var, new List<string>());
                        _ErrorList[var].Add("沒有學期歷程紀錄，無法判斷成績年級。");
                        canCalc &= false;
                    }
                    else
                    {
                        foreach (XmlElement history in new DSXmlHelper(semesterHistory).GetElements("History"))
                        {
                            int year, sems, gradeyear;
                            if (
                                int.TryParse(history.GetAttribute("SchoolYear"), out year) &&
                                int.TryParse(history.GetAttribute("Semester"), out sems) &&
                                int.TryParse(history.GetAttribute("GradeYear"), out gradeyear) &&
                                year == schoolyear && sems == semester
                                )
                                gradeYear = gradeyear;
                        }
                        if (gradeYear == null)
                        {
                            if (!_ErrorList.ContainsKey(var))
                                _ErrorList.Add(var, new List<string>());
                            _ErrorList[var].Add("學期歷程中沒有" + schoolyear + "學年度第" + semester + "學期的紀錄，無法判斷成績年級。");
                            canCalc &= false;
                        }
                    }
                    #endregion


                    #region 檢查修課及格標準
                    if (studentPassScoreDict.ContainsKey(var.StudentID))
                    {
                        bool hasError = false;
                        List<string> errorName = new List<string>();
                        foreach (string key in studentPassScoreDict[var.StudentID].Keys)
                        {
                            // 分數 -1 沒設定
                            if (studentPassScoreDict[var.StudentID][key] == -1)
                            {
                                string course_name = "";

                                if (studentFinalScoreDict.ContainsKey(var.StudentID))
                                {
                                    if (studentFinalScoreDict[var.StudentID].ContainsKey(key))
                                        course_name = studentFinalScoreDict[var.StudentID][key]["course_name"].ToString();
                                }

                                // 修課沒有設及格標準
                                if (!_ErrorList.ContainsKey(var))
                                    _ErrorList.Add(var, new List<string>());

                                _ErrorList[var].Add("沒有" + course_name + "的修課及格標準，無法計算。");

                                hasError = true;
                            }
                        }

                        if (hasError)
                        {
                            canCalc = false;
                        }
                    }
                    else
                    {
                        // 完全沒設
                        if (!_ErrorList.ContainsKey(var))
                            _ErrorList.Add(var, new List<string>());

                        _ErrorList[var].Add("沒有" + schoolyear + "學年度第" + semester + "學期的修課及格標準，無法計算。");
                        canCalc = false;
                    }
                    #endregion

                    #region 檢查課程的校部定、必選修、分項類別
                    //bool hasError = false;
                    if (courseRequied.ContainsKey(var.StudentID))
                    {
                        bool hasError = false;
                        foreach (string courseName in courseRequied[var.StudentID].Keys)
                        {
                            if (courseRequied[var.StudentID][courseName]["required_by"] == "")
                            {
                                if (!_ErrorList.ContainsKey(var))
                                    _ErrorList.Add(var, new List<string>());

                                _ErrorList[var].Add("沒有" + courseName + "的校部訂，無法計算。");

                                hasError = true;
                            }

                            if (courseRequied[var.StudentID][courseName]["is_required"] == "")
                            {
                                if (!_ErrorList.ContainsKey(var))
                                    _ErrorList.Add(var, new List<string>());

                                _ErrorList[var].Add("沒有" + courseName + "的必選修，無法計算。");

                                hasError = true;
                            }

                            if (courseRequied[var.StudentID][courseName]["score_type"] == "")
                            {

                                if (!_ErrorList.ContainsKey(var))
                                    _ErrorList.Add(var, new List<string>());

                                _ErrorList[var].Add("沒有" + courseName + "的分項類別，無法計算。");

                                hasError = true;
                            }
                            else
                            {
                                //允許的分項類別清單
                                List<string> entrys = new List<string>(new string[] { "學業", "體育", "國防通識", "健康與護理", "實習科目", "專業科目" });

                                // 分項類別錯誤，無法計算
                                string entry = courseRequied[var.StudentID][courseName]["score_type"];
                                if (!entrys.Contains(entry))
                                {
                                    if (!_ErrorList.ContainsKey(var))
                                        _ErrorList.Add(var, new List<string>());

                                    _ErrorList[var].Add("分項類別錯誤，無法計算。");

                                    hasError = true;
                                }
                            }
                        }
                        if (hasError)
                        {
                            canCalc = false;
                        }
                    }
                    #endregion


                    #region 處理計算規則
                    XmlElement scoreCalcRule = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID).ScoreCalcRuleElement;
                    if (scoreCalcRule == null)
                    {
                        if (!_ErrorList.ContainsKey(var))
                            _ErrorList.Add(var, new List<string>());
                        _ErrorList[var].Add("沒有設定成績計算規則。");
                        canCalc &= false;
                    }
                    else
                    {
                        DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                        bool tryParsebool;
                        int tryParseint;
                        decimal tryParseDecimal;

                        if (scoreCalcRule.SelectSingleNode("各項成績計算位數/科目成績計算位數") != null)
                        {
                            if (int.TryParse(helper.GetText("各項成績計算位數/科目成績計算位數/@位數"), out tryParseint))
                                decimals = tryParseint;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/科目成績計算位數/@四捨五入"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.四捨五入;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/科目成績計算位數/@無條件捨去"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.無條件捨去;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/科目成績計算位數/@無條件進位"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.無條件進位;
                        }
                        //if (scoreCalcRule.SelectSingleNode("延修及重讀成績處理規則/重讀成績") != null) 2018/5/29 穎驊註解 ，因應 [H成績][04] 計算學期科目成績調整 ，不再使用這個屬性 ，若有重讀，此屬性永遠是true，擇優
                        //{
                        //    if (bool.TryParse(helper.GetText("延修及重讀成績處理規則/重讀成績/@擇優採計成績"), out tryParsebool))
                        //        choseBetter = tryParsebool;
                        //}
                        //if (bool.TryParse(helper.GetText("重修成績/@登錄至原學期"), out tryParsebool)) 2018/5/29 穎驊註解 ，因應 [H成績][04] 計算學期科目成績調整 ，不再使用這個屬性，若有重覆科目級別，請至教務作業 /成績作業 課程重讀重修設定
                        //    writeToFirstSemester = tryParsebool;
                        foreach (XmlElement element in helper.GetElements("及格標準/學生類別"))
                        {
                            string cat = element.GetAttribute("類別");
                            bool useful = false;
                            //掃描學生的類別作比對
                            foreach (CategoryInfo catinfo in var.StudentCategorys)
                            {
                                if (catinfo.Name == cat || catinfo.FullName == cat)
                                    useful = true;
                            }
                            //學生是指定的類別或類別為"預設"
                            if (cat == "預設" || useful)
                            {
                                for (int gyear = 1; gyear <= 4; gyear++)
                                {
                                    switch (gyear)
                                    {
                                        case 1:
                                            if (decimal.TryParse(element.GetAttribute("一年級及格標準"), out tryParseDecimal))
                                            {
                                                if (!applyLimit.ContainsKey(gyear))
                                                    applyLimit.Add(gyear, tryParseDecimal);
                                                if (applyLimit[gyear] > tryParseDecimal)
                                                    applyLimit[gyear] = tryParseDecimal;
                                            }
                                            break;
                                        case 2:
                                            if (decimal.TryParse(element.GetAttribute("二年級及格標準"), out tryParseDecimal))
                                            {
                                                if (!applyLimit.ContainsKey(gyear))
                                                    applyLimit.Add(gyear, tryParseDecimal);
                                                if (applyLimit[gyear] > tryParseDecimal)
                                                    applyLimit[gyear] = tryParseDecimal;
                                            }
                                            break;
                                        case 3:
                                            if (decimal.TryParse(element.GetAttribute("三年級及格標準"), out tryParseDecimal))
                                            {
                                                if (!applyLimit.ContainsKey(gyear))
                                                    applyLimit.Add(gyear, tryParseDecimal);
                                                if (applyLimit[gyear] > tryParseDecimal)
                                                    applyLimit[gyear] = tryParseDecimal;
                                            }
                                            break;
                                        case 4:
                                            if (decimal.TryParse(element.GetAttribute("四年級及格標準"), out tryParseDecimal))
                                            {
                                                if (!applyLimit.ContainsKey(gyear))
                                                    applyLimit.Add(gyear, tryParseDecimal);
                                                if (applyLimit[gyear] > tryParseDecimal)
                                                    applyLimit[gyear] = tryParseDecimal;
                                            }
                                            break;
                                        default:
                                            break;
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                }
                #endregion

                XmlDocument doc = new XmlDocument();
                if (var.Fields.ContainsKey("SemesterSubjectCalcScore"))
                    var.Fields.Remove("SemesterSubjectCalcScore");
                XmlElement semesterSubjectCalcScoreElement = doc.CreateElement("SemesterSubjectCalcScore");
                if (canCalc)
                {
                    //已存在的學期科目成績<學年度<學期,<科目+級別,成績>>>
                    Dictionary<int, Dictionary<int, Dictionary<string, SemesterSubjectScoreInfo>>> semesterSubjectScoreList = new Dictionary<int, Dictionary<int, Dictionary<string, SemesterSubjectScoreInfo>>>();
                    //當學年度已紀錄之成績
                    Dictionary<SemesterSubjectScoreInfo, string> currentSubjectScoreList = new Dictionary<SemesterSubjectScoreInfo, string>();
                    ////重讀成績
                    //Dictionary<SemesterSubjectScoreInfo, string> repeatSubjectScoreList = new Dictionary<SemesterSubjectScoreInfo, string>();
                    ////重修成績
                    //Dictionary<SemesterSubjectScoreInfo, string> restudySubjectScoreList = new Dictionary<SemesterSubjectScoreInfo, string>();

                    // 重修成績
                    Dictionary<string, List<SemesterSubjectScoreInfo>> previousSubjectScoreMap = new Dictionary<string, List<SemesterSubjectScoreInfo>>();

                    // 補修成績
                    Dictionary<string, List<SemesterSubjectScoreInfo>> makeupSubjectScoreMap = new Dictionary<string, List<SemesterSubjectScoreInfo>>();

                    // 科目名稱+級別最高
                    Dictionary<string, decimal> dataCompareDict = new Dictionary<string, decimal>();


                    #region 先掃一遍把學生成績分類
                    foreach (SemesterSubjectScoreInfo scoreinfo in var.SemesterSubjectScoreList)
                    {
                        string key = scoreinfo.Subject.Trim() + "_" + scoreinfo.Level.Trim();

                        string studentSubjectKey = var.StudentID + "_" + key;

                        if (!dataCompareDict.ContainsKey(key))
                            dataCompareDict.Add(key, 0);

                        //最高分
                        decimal maxScore = 0;//                   

                        string[] scoreNames = new string[] { "原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績" };

                        foreach (string scorename in scoreNames)
                        {
                            decimal s;
                            if (decimal.TryParse(scoreinfo.Detail.GetAttribute(scorename), out s))
                            {
                                if (s > maxScore)
                                {
                                    maxScore = s;
                                }
                            }
                        }



                        // 目前學年度 currentSubjectScoreList
                        if (scoreinfo.SchoolYear == schoolyear && scoreinfo.Semester == semester)
                        {
                            currentSubjectScoreList.Add(scoreinfo, key);
                        }
                        else
                        {
                            if (maxScore > dataCompareDict[key])
                                dataCompareDict[key] = maxScore;

                            // 重修                            
                            if ((scoreinfo.SchoolYear < schoolyear || (scoreinfo.SchoolYear == schoolyear && scoreinfo.Semester < semester)))
                            {

                                // 重修成績
                                if (previousSubjectScoreMap.ContainsKey(studentSubjectKey))
                                {
                                    previousSubjectScoreMap[studentSubjectKey].Add(scoreinfo);
                                }
                                else
                                {
                                    previousSubjectScoreMap.Add(studentSubjectKey, new List<SemesterSubjectScoreInfo> { scoreinfo });
                                }

                                // 補修成績挖洞
                                if (scoreinfo.Detail.GetAttribute("是否補修成績") == "是")
                                {
                                    if (makeupSubjectScoreMap.ContainsKey(studentSubjectKey))
                                    {
                                        makeupSubjectScoreMap[studentSubjectKey].Add(scoreinfo);
                                    }
                                    else
                                    {
                                        List<SemesterSubjectScoreInfo> scoreinfoList = new List<SemesterSubjectScoreInfo> { scoreinfo };
                                        makeupSubjectScoreMap.Add(studentSubjectKey, scoreinfoList);
                                    }
                                }
                            }
                        }

                        //if (duplicateSubjectLevelMethodDict.ContainsKey(var.StudentID + "_" + key) && !duplicateSubjectLevelMethodDict_Afterfilter.ContainsKey(var.StudentID + "_" + key) && !currentSubjectScoreList.ContainsValue(key)) //假如有key 又非本學期的成績(可能本學期末已經先算過一次了) 則加入重覆計算的規則
                        //{
                        //    duplicateSubjectLevelMethodDict_Afterfilter.Add(var.StudentID + "_" + key, duplicateSubjectLevelMethodDict[var.StudentID + "_" + key]);
                        //}

                        // 只要 duplicateSubjectLevelMethodDict 有值且 value 非空白，且 Afterfilter 還沒有，就加入
                        //string dictKey = var.StudentID + "_" + key;
                        //if (duplicateSubjectLevelMethodDict.ContainsKey(dictKey)
                        //    && !duplicateSubjectLevelMethodDict_Afterfilter.ContainsKey(dictKey)
                        //    && !string.IsNullOrWhiteSpace(duplicateSubjectLevelMethodDict[dictKey]))
                        //{
                        //    duplicateSubjectLevelMethodDict_Afterfilter.Add(dictKey, duplicateSubjectLevelMethodDict[dictKey]);
                        //}




                        if (!semesterSubjectScoreList.ContainsKey(scoreinfo.SchoolYear))
                            semesterSubjectScoreList.Add(scoreinfo.SchoolYear, new Dictionary<int, Dictionary<string, SemesterSubjectScoreInfo>>());
                        if (!semesterSubjectScoreList[scoreinfo.SchoolYear].ContainsKey(scoreinfo.Semester))
                            semesterSubjectScoreList[scoreinfo.SchoolYear].Add(scoreinfo.Semester, new Dictionary<string, SemesterSubjectScoreInfo>());
                        if (!semesterSubjectScoreList[scoreinfo.SchoolYear][scoreinfo.Semester].ContainsKey(key))
                            semesterSubjectScoreList[scoreinfo.SchoolYear][scoreinfo.Semester].Add(key, scoreinfo);
                        else
                            semesterSubjectScoreList[scoreinfo.SchoolYear][scoreinfo.Semester][key] = scoreinfo;
                    }
                    #endregion

                    // 過濾取得重覆科目級別的計算處理方式
                    foreach (var kv in duplicateSubjectLevelMethodDict)
                    {
                        if (!string.IsNullOrWhiteSpace(kv.Value) && !duplicateSubjectLevelMethodDict_Afterfilter.ContainsKey(kv.Key))
                        {
                            duplicateSubjectLevelMethodDict_Afterfilter.Add(kv.Key, kv.Value);
                        }
                    }


                    #region 移除重讀跟重修成績重複年級的學期成績
                    // 因為整個過濾會產生只重讀某學年度學期資料會被過濾
                    //CleanUpRepeat(repeatSubjectScoreList);

                    //CleanUpRepeat(restudySubjectScoreList);
                    #endregion
                    //新增的學期成績資料
                    Dictionary<int, Dictionary<int, Dictionary<string, XmlElement>>> insertSemesterSubjectScoreList = new Dictionary<int, Dictionary<int, Dictionary<string, XmlElement>>>();
                    //修改的學期成績資料
                    Dictionary<int, Dictionary<int, Dictionary<string, XmlElement>>> updateSemesterSubjectScoreList = new Dictionary<int, Dictionary<int, Dictionary<string, XmlElement>>>();
                    #region 掃描修課紀錄填入新增或修改的清單中
                    foreach (StudentAttendCourseRecord sacRecord in var.AttendCourseList)
                    {
                        string key = sacRecord.Subject.Trim() + "_" + sacRecord.SubjectLevel.Trim();

                        bool noFScore = true;
                        // 檢查是否有指定總成績
                        if (studentHasFinalScoreDict.ContainsKey(var.StudentID))
                        {
                            if (studentHasFinalScoreDict[var.StudentID].ContainsKey(key))
                                noFScore = false;
                        }

                        if (noFScore)
                        {
                            if (!sacRecord.HasFinalScore && !sacRecord.NotIncludedInCalc)
                            {
                                if (!_ErrorList.ContainsKey(var))
                                    _ErrorList.Add(var, new List<string>());
                                _ErrorList[var].Add("" + sacRecord.CourseName + "沒有修課總成績，無法計算。");
                                continue;
                            }
                        }

                        //// 這判斷不確定作法先註解
                        //if (duplicateSubjectLevelMethodDict_Afterfilter.ContainsKey(var.StudentID + "_" + key) ? duplicateSubjectLevelMethodDict_Afterfilter[var.StudentID + "_" + key] == "" : false) // 如果使用者 沒有設定，要擋下，逼他們設定完畢才可以計算完畢
                        //{
                        //    int sy = 0, se = 0;
                        //    #region 找到最近一次修課紀錄
                        //    foreach (SemesterSubjectScoreInfo si in restudySubjectScoreList.Keys)
                        //    {
                        //        if (restudySubjectScoreList[si] == key)
                        //        {
                        //            if (si.SchoolYear > sy || (si.SchoolYear == sy && si.Semester > se))
                        //            {
                        //                sy = si.SchoolYear;
                        //                se = si.Semester;
                        //            }
                        //        }
                        //    }
                        //    #endregion

                        //    if (!_ErrorList.ContainsKey(var))
                        //        _ErrorList.Add(var, new List<string>());
                        //    _ErrorList[var].Add("課程名稱 :" + sacRecord.CourseName + "， 已於" + sy + "學年度 第" + se + "學期修習過，請至 教務作業/成績作業 重覆修課採計方式 設定。");
                        //    continue;
                        //}

                        // 取得該學生的處理規則
                        string subjectKey = sacRecord.Subject.Trim() + "_" + sacRecord.SubjectLevel.Trim();
                        string studentSubjectKey = var.StudentID + "_" + subjectKey;
                        string rule = duplicateSubjectLevelMethodDict_Afterfilter.ContainsKey(studentSubjectKey)
    ? duplicateSubjectLevelMethodDict_Afterfilter[studentSubjectKey]
    : "";

                        // 統一分類
                        string ruleType = "";
                        if (rule == "重修(寫回原學期)" || rule == "重修成績")
                            ruleType = "重修成績";
                        else if (rule == "重讀(擇優採計成績)" || rule == "再次修習")
                            ruleType = "再次修習";
                        else if (rule == "補修成績")
                            ruleType = "補修成績";

                        // 根據處理規則進行分支寫入
                        if (ruleType == "再次修習")
                        {
                            int sy = 0, se = 0;

                            //填入本學期科目成績
                            if (currentSubjectScoreList.ContainsValue(key))
                            {
                                #region 修改此學期已存在之成績
                                SemesterSubjectScoreInfo updateScoreInfo = null;
                                foreach (SemesterSubjectScoreInfo s in currentSubjectScoreList.Keys)
                                {
                                    if (currentSubjectScoreList[s] == key)
                                    {
                                        updateScoreInfo = s;
                                        break;
                                    }
                                }
                                sy = schoolyear;
                                se = semester;
                                if (!updateSemesterSubjectScoreList.ContainsKey(sy) || !updateSemesterSubjectScoreList[sy].ContainsKey(se) || !updateSemesterSubjectScoreList[sy][se].ContainsKey(key))
                                {
                                    //修改成績
                                    XmlElement updateScoreElement = updateScoreInfo.Detail;
                                    #region 重新填入課程資料


                                    updateScoreElement.SetAttribute("不計學分", sacRecord.NotIncludedInCredit ? "是" : "否");
                                    updateScoreElement.SetAttribute("不需評分", sacRecord.NotIncludedInCalc ? "是" : "否");
                                    updateScoreElement.SetAttribute("修課必選修", sacRecord.Required ? "必修" : "選修");
                                    updateScoreElement.SetAttribute("修課校部訂", (sacRecord.RequiredBy == "部訂" ? sacRecord.RequiredBy : "校訂"));
                                    updateScoreElement.SetAttribute("領域", sacRecord.Domain);
                                    updateScoreElement.SetAttribute("科目", sacRecord.Subject);
                                    updateScoreElement.SetAttribute("科目級別", sacRecord.SubjectLevel);
                                    updateScoreElement.SetAttribute("開課分項類別", sacRecord.Entry);

                                    updateScoreElement.SetAttribute("開課學分數", "" + sacRecord.CreditDec());

                                    if (specifySubjectNameDict.ContainsKey(sacRecord.StudentID))
                                    {
                                        string sKey = sacRecord.Subject + "_" + sacRecord.SubjectLevel;
                                        if (specifySubjectNameDict[sacRecord.StudentID].ContainsKey(sKey))
                                            updateScoreElement.SetAttribute("指定學年科目名稱", specifySubjectNameDict[sacRecord.StudentID][sKey]);
                                    }


                                    #endregion

                                    decimal? sfinalScore = null;

                                    bool fromPrevSemester = false;
                                    bool fromArchive = false;


                                    // 課程成績
                                    if (sacRecord.HasFinalScore)
                                        sfinalScore = GetRoundScore(sacRecord.FinalScore, decimals, mode);

                                    string sfKey = sacRecord.Subject + "_" + sacRecord.SubjectLevel;

                                    // 比對來自學期科目成績，科目名稱+級別相同
                                    if (dataCompareDict.ContainsKey(sfKey))
                                    {
                                        // 來自之前學期成績
                                        fromPrevSemester = true;

                                        if (sfinalScore.HasValue)
                                            if (dataCompareDict[sfKey] > sfinalScore.Value)
                                                sfinalScore = dataCompareDict[sfKey];
                                    }



                                    // 當沒有來自之前學期科目成績才比對來自封存成績，科目名稱+級別相同
                                    if (fromPrevSemester == false)
                                        if (dataCompareDict1.ContainsKey(sacRecord.StudentID))
                                        {
                                            if (dataCompareDict1[sacRecord.StudentID].ContainsKey(sfKey))
                                            {
                                                fromArchive = true; // 來源封存成績                           
                                                if (sfinalScore.HasValue)
                                                    if (dataCompareDict1[sacRecord.StudentID][sfKey] > sfinalScore.Value)
                                                    {
                                                        sfinalScore = dataCompareDict1[sacRecord.StudentID][sfKey];
                                                    }
                                            }
                                        }

                                    // 沒有修課成績填空，一定會有修課，擇優成績來自之前學期成績或封存成績，只能擇一。
                                    if (sfinalScore.HasValue)
                                        updateScoreElement.SetAttribute("原始成績", ("" + sfinalScore.Value));
                                    else
                                        updateScoreElement.SetAttribute("原始成績", "");


                                    // 當有直接指定總成績覆蓋
                                    if (studentFinalScoreDict.ContainsKey(sacRecord.StudentID))
                                    {
                                        string sKey = sacRecord.Subject + "_" + sacRecord.SubjectLevel;

                                        if (studentFinalScoreDict[sacRecord.StudentID].ContainsKey(sKey))
                                        {
                                            DataRow dr = studentFinalScoreDict[sacRecord.StudentID][sKey];

                                            string passing_standard = "", makeup_standard = "", remark = "", designate_final_score = "", subject_code = "";

                                            decimal passing_standard_score, makeup_standard_score, designate_final_score_score;

                                            if (dr["passing_standard"] != null)
                                                passing_standard = dr["passing_standard"].ToString();

                                            if (dr["makeup_standard"] != null)
                                                makeup_standard = dr["makeup_standard"].ToString();

                                            if (dr["remark"] != null)
                                                remark = dr["remark"].ToString();

                                            if (dr["subject_code"] != null)
                                                subject_code = dr["subject_code"].ToString();

                                            if (dr["designate_final_score"] != null)
                                                designate_final_score = dr["designate_final_score"].ToString();

                                            if (decimal.TryParse(passing_standard, out passing_standard_score))
                                                updateScoreElement.SetAttribute("修課及格標準", ("" + GetRoundScore(passing_standard_score, decimals, mode)));
                                            else
                                                updateScoreElement.SetAttribute("修課及格標準", "");

                                            if (decimal.TryParse(makeup_standard, out makeup_standard_score))
                                                updateScoreElement.SetAttribute("修課補考標準", ("" + GetRoundScore(makeup_standard_score, decimals, mode)));
                                            else
                                                updateScoreElement.SetAttribute("修課補考標準", "");

                                            updateScoreElement.SetAttribute("註記", "");

                                            if (decimal.TryParse(designate_final_score, out designate_final_score_score))
                                            {
                                                updateScoreElement.SetAttribute("修課直接指定總成績", ("" + GetRoundScore(designate_final_score_score, decimals, mode)));

                                                // 註解是因經過2024/4/26討論，修課直接指定總成績不應該覆蓋原始成績，需要保留原始成績。                                              
                                                updateScoreElement.SetAttribute("原始成績", ("" + GetRoundScore(designate_final_score_score, decimals, mode)));


                                                //updateScoreElement.SetAttribute("原始成績", (sacRecord.NotIncludedInCalc ? "" : "" + GetRoundScore(designate_final_score_score, decimals, mode)));


                                                updateScoreElement.SetAttribute("註記", "修課成績：" + ("" + GetRoundScore(sacRecord.FinalScore, decimals, mode)));
                                            }
                                            else
                                            {
                                                updateScoreElement.SetAttribute("修課直接指定總成績", "");
                                            }

                                            updateScoreElement.SetAttribute("修課備註", remark);
                                            updateScoreElement.SetAttribute("修課科目代碼", subject_code);
                                        }
                                    }

                                    //做取得學分判斷
                                    #region 做取得學分判斷及填入擇優採計成績
                                    //最高分
                                    decimal maxScore = 0;// sacRecord.FinalScore;
                                    #region 抓最高分
                                    string[] scoreNames = new string[] { "原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績" };
                                    foreach (string scorename in scoreNames)
                                    {
                                        decimal s;
                                        if (decimal.TryParse(updateScoreElement.GetAttribute(scorename), out s))
                                        {
                                            if (s > maxScore)
                                            {
                                                maxScore = s;
                                            }
                                        }
                                    }
                                    #endregion


                                    decimal passscore;

                                    // 新寫及格標準
                                    passscore = 100;
                                    if (studentPassScoreDict.ContainsKey(var.StudentID))
                                    {
                                        if (studentPassScoreDict[var.StudentID].ContainsKey(key))
                                        {
                                            passscore = studentPassScoreDict[var.StudentID][key];
                                        }
                                        else
                                        {
                                            if (!applyLimit.ContainsKey(updateScoreInfo.GradeYear))
                                                passscore = 60;
                                            else
                                                passscore = applyLimit[updateScoreInfo.GradeYear];
                                        }
                                    }


                                    // 2024/7/5 會議決議，需要計算學分使用成績判斷是否取得學分
                                    if (sacRecord.NotIncludedInCredit == false)
                                    {
                                        updateScoreElement.SetAttribute("是否取得學分", maxScore >= passscore ? "是" : "否");
                                    }

                                    updateScoreElement.SetAttribute("註記", "修課成績：" + GetRoundScore(sacRecord.FinalScore, decimals, mode));

                                    #endregion
                                    if (!updateSemesterSubjectScoreList.ContainsKey(sy)) updateSemesterSubjectScoreList.Add(sy, new Dictionary<int, Dictionary<string, XmlElement>>());
                                    if (!updateSemesterSubjectScoreList[sy].ContainsKey(se)) updateSemesterSubjectScoreList[sy].Add(se, new Dictionary<string, XmlElement>());
                                    updateSemesterSubjectScoreList[sy][se].Add(key, updateScoreElement);


                                    // 取得學期歷程資料
                                    string HisClassName = "", HisStudentNumber = "";
                                    int? HisSeatNo = null;
                                    try
                                    {
                                        if (var.Fields.ContainsKey("SemesterHistory"))
                                        {
                                            XmlElement xmlElement = var.Fields["SemesterHistory"] as XmlElement;
                                            XElement elmRoot = XElement.Parse(xmlElement.OuterXml);

                                            var matched = elmRoot.Elements("History")
                                                .FirstOrDefault(e =>
                                                    (string)e.Attribute("SchoolYear") == schoolyear.ToString() &&
                                                    (string)e.Attribute("Semester") == semester.ToString());

                                            if (matched != null)
                                            {
                                                HisClassName = (string)matched.Attribute("ClassName") ?? "";
                                                HisStudentNumber = (string)matched.Attribute("StudentNumber") ?? "";
                                                int seatNo;
                                                if (int.TryParse((string)matched.Attribute("SeatNo"), out seatNo))
                                                    HisSeatNo = seatNo;
                                            }
                                        }
                                    }
                                    catch { }

                                    // 判斷是否取得學分
                                    string ScoreP = "";
                                    if ((sfinalScore.HasValue ? sfinalScore.Value : 0) >= 60) // 這裡 60 可以改成你的及格標準變數
                                        ScoreP = "1";
                                    else
                                        ScoreP = "0";


                                    // 課程代碼
                                    string courseCode = updateScoreElement.GetAttribute("修課科目代碼");

                                    // 是否採計學分，預設"1"
                                    string useCredit = "1";
                                    if (sacRecord.NotIncludedInCredit)
                                        useCredit = "2";
                                    else if (sacRecord.NotIncludedInCalc)
                                        useCredit = "3";

                                    // 檢查課程代碼 CodePass
                                    bool codePass = Utility.IsValidCourseCode(courseCode);

                                    // 只有來自封存成績，填入轉學轉科工作表
                                    if (fromArchive)
                                    {
                                        var transferRec = new SubjectScoreRec108
                                        {
                                            StudentID = sacRecord.StudentID,
                                            IDNumber = sacRecord.IDNumber,
                                            Birthday = sacRecord.Birthday,
                                            SchoolYear = schoolyear.ToString(),
                                            Semester = semester.ToString(),
                                            CourseCode = courseCode,
                                            SubjectName = sacRecord.Subject,
                                            SubjectLevel = sacRecord.SubjectLevel,
                                            GradeYear = gradeYear.HasValue ? gradeYear.Value.ToString() : "",
                                            Credit = sacRecord.CreditDec().ToString(),
                                            Score = sfinalScore.HasValue ? sfinalScore.Value.ToString() : "",
                                            ScoreP = ScoreP,
                                            useCredit = useCredit,
                                            Name = var.StudentName,
                                            HisClassName = HisClassName,
                                            HisSeatNo = HisSeatNo,
                                            HisStudentNumber = HisStudentNumber,
                                            ClassName = var.RefClass.ClassName,
                                            SeatNo = sacRecord.SeatNo,
                                            StudentNumber = sacRecord.StudentNumber,
                                            isScScore = true,
                                            checkPass = true,
                                            CodePass = codePass,
                                            RepeatMemo = "2",
                                            RepeatScoreP = ScoreP,
                                            RepeatScore = updateScoreElement.GetAttribute("原始成績")
                                        };
                                        // 加入 4.4 轉學轉科名冊
                                        transferScoreList.Add(transferRec);
                                    }

                                    // 來自之前學期成績，填入重修重讀名冊，重讀工作表
                                    if (fromPrevSemester)
                                    {
                                        var repeatRec = new SubjectScoreRec108
                                        {
                                            StudentID = sacRecord.StudentID,
                                            IDNumber = sacRecord.IDNumber,
                                            Birthday = sacRecord.Birthday,
                                            SchoolYear = schoolyear.ToString(),
                                            Semester = semester.ToString(),
                                            CourseCode = courseCode,
                                            SubjectName = sacRecord.Subject,
                                            SubjectLevel = sacRecord.SubjectLevel,
                                            GradeYear = gradeYear.HasValue ? gradeYear.Value.ToString() : "",
                                            Credit = sacRecord.CreditDec().ToString(),
                                            Score = GetRoundScore(sacRecord.FinalScore, decimals, mode).ToString(),
                                            ScoreP = ScoreP,
                                            useCredit = useCredit,
                                            Name = var.StudentName,
                                            HisClassName = HisClassName,
                                            HisSeatNo = HisSeatNo,
                                            HisStudentNumber = HisStudentNumber,
                                            ClassName = var.RefClass.ClassName,
                                            SeatNo = sacRecord.SeatNo,
                                            StudentNumber = sacRecord.StudentNumber,
                                            isScScore = true,
                                            checkPass = true,
                                            CodePass = codePass,
                                            RepeatMemo = "2",
                                            RepeatScoreP = ScoreP,
                                            RepeatScore = updateScoreElement.GetAttribute("原始成績")
                                        };
                                        // 加入 5.3 重讀成績名冊
                                        repeatScoreList.Add(repeatRec);
                                    }
                                }
                                #endregion
                            }
                            else
                            {
                                #region 新增一筆成績
                                sy = schoolyear; se = semester;
                                if (!insertSemesterSubjectScoreList.ContainsKey(sy) || !insertSemesterSubjectScoreList[sy].ContainsKey(se) || !insertSemesterSubjectScoreList[sy][se].ContainsKey(key))
                                {
                                    //科目名稱空白有錯誤時執行，2024/5/23討論，當課程的科目空白不計算
                                    if (string.IsNullOrEmpty(sacRecord.Subject) || string.IsNullOrWhiteSpace(sacRecord.Subject))
                                    {
                                        //_WarningList為空就先建立
                                        if (_WarningList == null) _WarningList = new List<string>();
                                        //_WarningList不存在此課程名稱才加入
                                        if (!_WarningList.Contains(sacRecord.CourseName))
                                        {
                                            _WarningList.Add(sacRecord.CourseName);
                                        }
                                        continue;
                                    }


                                    #region 加入新的資料
                                    XmlElement newScoreInfo = doc.CreateElement("Subject");
                                    newScoreInfo.SetAttribute("不計學分", sacRecord.NotIncludedInCredit ? "是" : "否");
                                    newScoreInfo.SetAttribute("不需評分", sacRecord.NotIncludedInCalc ? "是" : "否");
                                    newScoreInfo.SetAttribute("修課必選修", sacRecord.Required ? "必修" : "選修");
                                    newScoreInfo.SetAttribute("修課校部訂", (sacRecord.RequiredBy == "部訂" ? sacRecord.RequiredBy : "校訂"));
                                    newScoreInfo.SetAttribute("領域", sacRecord.Domain);
                                    newScoreInfo.SetAttribute("科目", sacRecord.Subject);
                                    newScoreInfo.SetAttribute("科目級別", sacRecord.SubjectLevel);
                                    newScoreInfo.SetAttribute("開課分項類別", sacRecord.Entry);
                                    newScoreInfo.SetAttribute("開課學分數", "" + sacRecord.CreditDec());

                                    decimal? sfinalScore = null;
                                    bool fromPrevSemester = false;
                                    bool fromArchive = false;

                                    // 課程成績
                                    if (sacRecord.HasFinalScore)
                                        sfinalScore = GetRoundScore(sacRecord.FinalScore, decimals, mode);

                                    string sfKey = sacRecord.Subject + "_" + sacRecord.SubjectLevel;

                                    // 比對來自學期科目成績，科目名稱+級別相同
                                    if (dataCompareDict.ContainsKey(sfKey))
                                    {
                                        // 來自之前學期成績
                                        fromPrevSemester = true;

                                        if (sfinalScore.HasValue)
                                            if (dataCompareDict[sfKey] > sfinalScore.Value)
                                                sfinalScore = dataCompareDict[sfKey];
                                    }




                                    // 當沒有之前學期成績，才使用自封存成績來比對，科目名稱+級別相同
                                    if (fromPrevSemester == false)
                                        if (dataCompareDict1.ContainsKey(sacRecord.StudentID))
                                        {
                                            if (dataCompareDict1[sacRecord.StudentID].ContainsKey(sfKey))
                                            {
                                                fromArchive = true; // 來源封存成績

                                                if (sfinalScore.HasValue)
                                                    if (dataCompareDict1[sacRecord.StudentID][sfKey] > sfinalScore.Value)
                                                    {
                                                        sfinalScore = dataCompareDict1[sacRecord.StudentID][sfKey];

                                                    }
                                            }
                                        }



                                    // 沒有修課成績填空
                                    if (sfinalScore.HasValue)
                                        newScoreInfo.SetAttribute("原始成績", ("" + sfinalScore.Value));
                                    else
                                        newScoreInfo.SetAttribute("原始成績", "");


                                    if (specifySubjectNameDict.ContainsKey(sacRecord.StudentID))
                                    {
                                        string sKey = sacRecord.Subject + "_" + sacRecord.SubjectLevel;
                                        if (specifySubjectNameDict[sacRecord.StudentID].ContainsKey(sKey))
                                            newScoreInfo.SetAttribute("指定學年科目名稱", specifySubjectNameDict[sacRecord.StudentID][sKey]);
                                    }

                                    // 課程代碼
                                    string CourseCode = "";

                                    // 當有直接指定總成績覆蓋
                                    if (studentFinalScoreDict.ContainsKey(sacRecord.StudentID))
                                    {
                                        string sKey = sacRecord.Subject + "_" + sacRecord.SubjectLevel;

                                        if (studentFinalScoreDict[sacRecord.StudentID].ContainsKey(sKey))
                                        {
                                            DataRow dr = studentFinalScoreDict[sacRecord.StudentID][sKey];

                                            string passing_standard = "", makeup_standard = "", remark = "", designate_final_score = "", subject_code = "";

                                            decimal passing_standard_score, makeup_standard_score, designate_final_score_score;

                                            if (dr["passing_standard"] != null)
                                                passing_standard = dr["passing_standard"].ToString();

                                            if (dr["makeup_standard"] != null)
                                                makeup_standard = dr["makeup_standard"].ToString();

                                            if (dr["remark"] != null)
                                                remark = dr["remark"].ToString();

                                            if (dr["subject_code"] != null)
                                                subject_code = dr["subject_code"].ToString();

                                            if (dr["designate_final_score"] != null)
                                                designate_final_score = dr["designate_final_score"].ToString();

                                            if (decimal.TryParse(passing_standard, out passing_standard_score))
                                                newScoreInfo.SetAttribute("修課及格標準", ("" + GetRoundScore(passing_standard_score, decimals, mode)));
                                            else
                                                newScoreInfo.SetAttribute("修課及格標準", "");

                                            if (decimal.TryParse(makeup_standard, out makeup_standard_score))
                                                newScoreInfo.SetAttribute("修課補考標準", ("" + GetRoundScore(makeup_standard_score, decimals, mode)));
                                            else
                                                newScoreInfo.SetAttribute("修課補考標準", "");

                                            if (decimal.TryParse(designate_final_score, out designate_final_score_score))
                                            {
                                                newScoreInfo.SetAttribute("修課直接指定總成績", ("" + GetRoundScore(designate_final_score_score, decimals, mode)));

                                                // 註解是因經過2024/4/26討論，修課直接指定總成績不應該覆蓋原始成績，需要保留原始成績。

                                                newScoreInfo.SetAttribute("原始成績", ("" + GetRoundScore(designate_final_score_score, decimals, mode)));


                                                newScoreInfo.SetAttribute("註記", "修課成績：" + ("" + GetRoundScore(sacRecord.FinalScore, decimals, mode)));
                                            }
                                            else
                                            {
                                                newScoreInfo.SetAttribute("修課直接指定總成績", "");
                                            }

                                            newScoreInfo.SetAttribute("修課備註", remark);
                                            newScoreInfo.SetAttribute("修課科目代碼", subject_code);
                                            CourseCode = subject_code;
                                        }
                                    }

                                    newScoreInfo.SetAttribute("重修成績", "");
                                    newScoreInfo.SetAttribute("學年調整成績", "");
                                    newScoreInfo.SetAttribute("擇優採計成績", "");
                                    newScoreInfo.SetAttribute("補考成績", "");
                                    //做取得學分判斷
                                    #region 做取得學分判斷及填入擇優採計成績
                                    //最高分
                                    decimal maxScore = 0;// = sacRecord.FinalScore;
                                    #region 抓最高分
                                    string[] scoreNames = new string[] { "原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績" };
                                    foreach (string scorename in scoreNames)
                                    {
                                        decimal s;
                                        if (decimal.TryParse(newScoreInfo.GetAttribute(scorename), out s))
                                        {
                                            if (s > maxScore)
                                            {
                                                maxScore = s;
                                            }
                                        }
                                    }
                                    #endregion


                                    decimal passscore;

                                    // 新寫及格標準
                                    passscore = 100;
                                    if (studentPassScoreDict.ContainsKey(var.StudentID))
                                    {
                                        if (studentPassScoreDict[var.StudentID].ContainsKey(key))
                                        {
                                            passscore = studentPassScoreDict[var.StudentID][key];
                                        }
                                        else
                                        {
                                            if (!applyLimit.ContainsKey((int)gradeYear))
                                                passscore = 60;
                                            else
                                                passscore = applyLimit[(int)gradeYear];
                                        }
                                    }

                                    #endregion


                                    // 2024/7/5 會議決議，需要計算學分使用成績判斷是否取得學分
                                    if (sacRecord.NotIncludedInCredit == false)
                                    {
                                        newScoreInfo.SetAttribute("是否取得學分", maxScore >= passscore ? "是" : "否");
                                    }

                                    newScoreInfo.SetAttribute("註記", "修課成績：" + GetRoundScore(sacRecord.FinalScore, decimals, mode));

                                    #endregion
                                    if (!insertSemesterSubjectScoreList.ContainsKey(sy)) insertSemesterSubjectScoreList.Add(sy, new Dictionary<int, Dictionary<string, XmlElement>>());
                                    if (!insertSemesterSubjectScoreList[sy].ContainsKey(se)) insertSemesterSubjectScoreList[sy].Add(se, new Dictionary<string, XmlElement>());
                                    insertSemesterSubjectScoreList[sy][se].Add(key, newScoreInfo);

                                    // --- 寫入學期歷程
                                    // 取得學期歷程資料
                                    string HisClassName = "", HisStudentNumber = "";
                                    int? HisSeatNo = null;
                                    try
                                    {
                                        if (var.Fields.ContainsKey("SemesterHistory"))
                                        {
                                            XmlElement xmlElement = var.Fields["SemesterHistory"] as XmlElement;
                                            XElement elmRoot = XElement.Parse(xmlElement.OuterXml);

                                            var matched = elmRoot.Elements("History")
                                                .FirstOrDefault(e =>
                                                    (string)e.Attribute("SchoolYear") == schoolyear.ToString() &&
                                                    (string)e.Attribute("Semester") == semester.ToString());

                                            if (matched != null)
                                            {
                                                HisClassName = (string)matched.Attribute("ClassName") ?? "";
                                                HisStudentNumber = (string)matched.Attribute("StudentNumber") ?? "";
                                                int seatNo;
                                                if (int.TryParse((string)matched.Attribute("SeatNo"), out seatNo))
                                                    HisSeatNo = seatNo;
                                            }
                                        }
                                    }
                                    catch { }

                                    // 判斷是否取得學分
                                    string ScoreP = "";
                                    if ((sfinalScore.HasValue ? sfinalScore.Value : 0) >= 60) // 這裡 60 可以改成你的及格標準變數
                                        ScoreP = "1";
                                    else
                                        ScoreP = "0";

                                    // 課程代碼
                                    string courseCode = CourseCode;

                                    // 是否採計學分，預設"1"
                                    string useCredit = "1";
                                    if (sacRecord.NotIncludedInCredit)
                                        useCredit = "2";
                                    else if (sacRecord.NotIncludedInCalc)
                                        useCredit = "3";

                                    // 檢查課程代碼 CodePass
                                    bool codePass = Utility.IsValidCourseCode(courseCode);

                                    // 只有來自封存成績，填入轉學轉科工作表
                                    if (fromArchive)
                                    {
                                        var transferRec = new SubjectScoreRec108
                                        {
                                            StudentID = sacRecord.StudentID,
                                            IDNumber = sacRecord.IDNumber,
                                            Birthday = sacRecord.Birthday,
                                            SchoolYear = schoolyear.ToString(),
                                            Semester = semester.ToString(),
                                            CourseCode = courseCode,
                                            SubjectName = sacRecord.Subject,
                                            SubjectLevel = sacRecord.SubjectLevel,
                                            GradeYear = gradeYear.HasValue ? gradeYear.Value.ToString() : "",
                                            Credit = sacRecord.CreditDec().ToString(),
                                            Score = sfinalScore.HasValue ? sfinalScore.Value.ToString() : "",
                                            ScoreP = ScoreP,
                                            useCredit = useCredit,
                                            Name = var.StudentName,
                                            HisClassName = HisClassName,
                                            HisSeatNo = HisSeatNo,
                                            HisStudentNumber = HisStudentNumber,
                                            ClassName = var.RefClass.ClassName,
                                            SeatNo = sacRecord.SeatNo,
                                            StudentNumber = sacRecord.StudentNumber,
                                            isScScore = true,
                                            checkPass = true,
                                            CodePass = codePass,
                                            RepeatMemo = "2",
                                            RepeatScoreP = ScoreP,
                                            RepeatScore = newScoreInfo.GetAttribute("原始成績")
                                        };
                                        // 加入 4.4 轉學轉科名冊
                                        transferScoreList.Add(transferRec);
                                    }

                                    // 來自之前學期成績，填入重修重讀名冊，重讀工作表
                                    if (fromPrevSemester)
                                    {
                                        var repeatRec = new SubjectScoreRec108
                                        {
                                            StudentID = sacRecord.StudentID,
                                            IDNumber = sacRecord.IDNumber,
                                            Birthday = sacRecord.Birthday,
                                            SchoolYear = schoolyear.ToString(),
                                            Semester = semester.ToString(),
                                            CourseCode = courseCode,
                                            SubjectName = sacRecord.Subject,
                                            SubjectLevel = sacRecord.SubjectLevel,
                                            GradeYear = gradeYear.HasValue ? gradeYear.Value.ToString() : "",
                                            Credit = sacRecord.CreditDec().ToString(),
                                            Score = GetRoundScore(sacRecord.FinalScore, decimals, mode).ToString(),
                                            ScoreP = ScoreP,
                                            useCredit = useCredit,
                                            Name = var.StudentName,
                                            HisClassName = HisClassName,
                                            HisSeatNo = HisSeatNo,
                                            HisStudentNumber = HisStudentNumber,
                                            ClassName = var.RefClass.ClassName,
                                            SeatNo = sacRecord.SeatNo,
                                            StudentNumber = sacRecord.StudentNumber,
                                            isScScore = true,
                                            checkPass = true,
                                            CodePass = codePass,
                                            RepeatMemo = "2",
                                            RepeatScoreP = ScoreP,
                                            RepeatScore = newScoreInfo.GetAttribute("原始成績")
                                        };

                                        // 加入 5.3 重讀成績名冊
                                        repeatScoreList.Add(repeatRec);
                                    }

                                }
                                #endregion
                            }
                        }
                        else if (ruleType == "重修成績")
                        {
                            if (previousSubjectScoreMap.ContainsKey(studentSubjectKey))
                            {
                                #region 寫入重修成績回原學期
                                int sy = 0, se = 0;

                                foreach (SemesterSubjectScoreInfo previousSubjectScoreInfo in previousSubjectScoreMap[studentSubjectKey])
                                {
                                    sy = previousSubjectScoreInfo.SchoolYear;
                                    se = previousSubjectScoreInfo.Semester;
                                    string key1 = previousSubjectScoreInfo.Subject.Trim() + "_" + previousSubjectScoreInfo.Level.Trim();

                                    // 假設 sacRecord.FinalScore 為本次重修課程成績
                                    decimal roundedScore = GetRoundScore(sacRecord.FinalScore, decimals, mode);

                                    //寫入重修紀錄
                                    XmlElement updateScoreElement = previousSubjectScoreInfo.Detail;
                                    //updateScoreElement.SetAttribute("重修成績", "" + GetRoundScore(sacRecord.FinalScore, decimals, mode));

                                    // C# 寫一段程式，只有當新分數比舊的「重修成績」高時才寫入 SetAttribute("重修成績", ...)；否則不寫入。
                                    decimal previousScore;
                                    if (decimal.TryParse(updateScoreElement.GetAttribute("重修成績"), out previousScore))
                                    {
                                        if (roundedScore > previousScore)
                                        {
                                            updateScoreElement.SetAttribute("重修成績", "" + roundedScore);
                                        }
                                    }
                                    else
                                    {
                                        updateScoreElement.SetAttribute("重修成績", "" + roundedScore);
                                    }

                                    previousSubjectScoreInfo.Detail.SetAttribute("重修學年度", schoolyear.ToString());
                                    previousSubjectScoreInfo.Detail.SetAttribute("重修學期", semester.ToString());


                                    //做取得學分判斷
                                    #region 做取得學分判斷
                                    //最高分
                                    decimal maxScore = 0;// = sacRecord.FinalScore;
                                    #region 抓最高分

                                    string[] scoreNames = new string[] { "原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績" };

                                    foreach (string scorename in scoreNames)
                                    {
                                        decimal s;
                                        if (decimal.TryParse(updateScoreElement.GetAttribute(scorename), out s))
                                        {
                                            if (s > maxScore)
                                            {
                                                maxScore = s;
                                            }
                                        }
                                    }
                                    #endregion
                                    decimal passscore;


                                    // 新寫及格標準
                                    passscore = 100;
                                    if (studentPassScoreDict.ContainsKey(var.StudentID))
                                    {
                                        if (studentPassScoreDict[var.StudentID].ContainsKey(key))
                                        {
                                            passscore = studentPassScoreDict[var.StudentID][key];
                                        }
                                        else
                                        {
                                            if (!applyLimit.ContainsKey(previousSubjectScoreInfo.GradeYear))
                                                passscore = 60;
                                            else
                                                passscore = applyLimit[previousSubjectScoreInfo.GradeYear];
                                        }
                                    }

                                    // 2024/7/5 會議決議，需要計算學分使用成績判斷是否取得學分
                                    if (updateScoreElement.GetAttribute("不計學分") == "否")
                                    {
                                        updateScoreElement.SetAttribute("是否取得學分", maxScore >= passscore ? "是" : "否");
                                    }


                                    #endregion
                                    if (!updateSemesterSubjectScoreList.ContainsKey(sy)) updateSemesterSubjectScoreList.Add(sy, new Dictionary<int, Dictionary<string, XmlElement>>());
                                    if (!updateSemesterSubjectScoreList[sy].ContainsKey(se)) updateSemesterSubjectScoreList[sy].Add(se, new Dictionary<string, XmlElement>());
                                    updateSemesterSubjectScoreList[sy][se].Add(key, updateScoreElement);

                                    // 重修成績寫入學期歷程 ---

                                    string HisClassName = "", HisStudentNumber = "";
                                    int? HisSeatNo = null;
                                    try
                                    {

                                        // 讀取學生學期對照表資料
                                        if (var.Fields.ContainsKey("SemesterHistory"))
                                        {
                                            XmlElement xmlElement = var.Fields["SemesterHistory"] as XmlElement;
                                            XElement elmRoot = XElement.Parse(xmlElement.OuterXml);

                                            // 找到第一個符合 SchoolYear, Semester 的節點
                                            var matched = elmRoot.Elements("History")
                                                .FirstOrDefault(e =>
                                                    (string)e.Attribute("SchoolYear") == sy.ToString() &&
                                                    (string)e.Attribute("Semester") == se.ToString());

                                            if (matched != null)
                                            {
                                                HisClassName = (string)matched.Attribute("ClassName") ?? "";
                                                HisStudentNumber = (string)matched.Attribute("StudentNumber") ?? "";

                                                // SeatNo 轉型
                                                int seatNo;
                                                if (int.TryParse((string)matched.Attribute("SeatNo"), out seatNo))
                                                    HisSeatNo = seatNo;
                                                else
                                                    HisSeatNo = null;
                                            }
                                            else
                                            {

                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine(ex.Message);
                                    }

                                    string ScoreP = "";

                                    // 根據 makeUpScoreInfo.Detail.GetAttribute("是否取得學分") 的值，如果是 "是" 就設定 ssr.ScoreP 為 1，否則為 0。
                                    // Please use if-else syntax.
                                    if (previousSubjectScoreInfo.Detail.GetAttribute("是否取得學分") == "是")
                                        ScoreP = "1";
                                    else
                                        ScoreP = "0";


                                    string ReAScore = "-1";
                                    // 幫我寫一段 C# 程式碼，從 makeUpScoreInfo.Detail 取得「補考成績」和「修課及格標準」兩個屬性。如果「補考成績」能轉為數字，且在 0 到「修課及格標準」之間，則 reScore 等於該分數字串，否則 reScore = "-1"。如果「修課及格標準」無法轉數字，預設用 60。
                                    if (decimal.TryParse(previousSubjectScoreInfo.Detail.GetAttribute("重修成績"), out decimal reScoreValue2))
                                    {
                                        if (decimal.TryParse(previousSubjectScoreInfo.Detail.GetAttribute("修課及格標準"), out decimal passingStandard))
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
                                    string passScoreStr = previousSubjectScoreInfo.Detail.GetAttribute("修課及格標準");
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
                                    string courseCode = previousSubjectScoreInfo.Detail.GetAttribute("修課科目代碼").Trim();

                                    // 假設 makeUpScoreInfo.Detail 是 XElement 或有 GetAttribute 方法
                                    string notCountCredit = previousSubjectScoreInfo.Detail.GetAttribute("不計學分");
                                    string notNeedScore = previousSubjectScoreInfo.Detail.GetAttribute("不需評分");

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

                                    var rec = new SubjectScoreRec108
                                    {
                                        StudentID = sacRecord.StudentID,
                                        IDNumber = sacRecord.IDNumber,
                                        Birthday = sacRecord.Birthday,
                                        SchoolYear = previousSubjectScoreInfo.SchoolYear.ToString(),
                                        Semester = previousSubjectScoreInfo.Semester.ToString(),
                                        CourseCode = courseCode,
                                        SubjectName = previousSubjectScoreInfo.Subject.Trim(),
                                        SubjectLevel = previousSubjectScoreInfo.Level.Trim(),
                                        GradeYear = previousSubjectScoreInfo.GradeYear.ToString(),
                                        Credit = previousSubjectScoreInfo.CreditDec().ToString(),
                                        Score = previousSubjectScoreInfo.Score.ToString(),

                                        ScoreP = ScoreP,
                                        ScScoreType = "3",
                                        useCredit = useCredit,
                                        Text = "",
                                        Name = var.StudentName,
                                        HisClassName = HisClassName,
                                        HisSeatNo = HisSeatNo,
                                        HisStudentNumber = HisStudentNumber,
                                        ClassName = var.RefClass.ClassName,
                                        SeatNo = sacRecord.SeatNo,
                                        StudentNumber = sacRecord.StudentNumber,
                                        isScScore = true,
                                        checkPass = true,
                                        CodePass = codePass,
                                        ReAScoreType = "3",
                                        ReAScore = ReAScore,
                                        ReAScoreP = ReAScoreP
                                    };
                                    restudyScoreList.Add(rec);

                                }
                                #endregion
                            }
                        }
                        else if (ruleType == "補修成績")
                        {
                            if (makeupSubjectScoreMap.ContainsKey(studentSubjectKey))
                            {
                                foreach (SemesterSubjectScoreInfo makeUpScoreInfo in makeupSubjectScoreMap[studentSubjectKey])
                                {
                                    int sy = makeUpScoreInfo.SchoolYear;
                                    int se = makeUpScoreInfo.Semester;
                                    string key1 = makeUpScoreInfo.Subject.Trim() + "_" + makeUpScoreInfo.Level.Trim();

                                    // 假設 sacRecord.FinalScore 為本次補修課程成績
                                    decimal roundedScore = GetRoundScore(sacRecord.FinalScore, decimals, mode);
                                    makeUpScoreInfo.Detail.SetAttribute("原始成績", roundedScore.ToString());
                                    makeUpScoreInfo.Detail.SetAttribute("補修學年度", schoolyear.ToString());
                                    makeUpScoreInfo.Detail.SetAttribute("補修學期", semester.ToString());

                                    decimal passscore;
                                    passscore = 100;
                                    if (studentPassScoreDict.ContainsKey(var.StudentID))
                                    {
                                        if (studentPassScoreDict[var.StudentID].ContainsKey(key))
                                        {
                                            passscore = studentPassScoreDict[var.StudentID][key];
                                        }
                                        else
                                        {
                                            if (!applyLimit.ContainsKey(makeUpScoreInfo.GradeYear))
                                                passscore = 60;
                                            else
                                                passscore = applyLimit[makeUpScoreInfo.GradeYear];
                                        }
                                    }


                                    // 2024/7/5 會議決議，需要計算學分使用成績判斷是否取得學分
                                    if (sacRecord.NotIncludedInCredit == false)
                                    {
                                        makeUpScoreInfo.Detail.SetAttribute("是否取得學分", roundedScore >= passscore ? "是" : "否");
                                    }


                                    if (!updateSemesterSubjectScoreList.ContainsKey(sy))
                                        updateSemesterSubjectScoreList.Add(sy, new Dictionary<int, Dictionary<string, XmlElement>>());
                                    if (!updateSemesterSubjectScoreList[sy].ContainsKey(se))
                                        updateSemesterSubjectScoreList[sy].Add(se, new Dictionary<string, XmlElement>());
                                    updateSemesterSubjectScoreList[sy][se][key1] = makeUpScoreInfo.Detail;

                                    // --- 處理補修成績寫入學期歷程資料 4.3 補修成績
                                    string HisClassName = "", HisStudentNumber = "";
                                    int? HisSeatNo = null;
                                    try
                                    {

                                        // 讀取學生學期對照表資料
                                        if (var.Fields.ContainsKey("SemesterHistory"))
                                        {
                                            XmlElement xmlElement = var.Fields["SemesterHistory"] as XmlElement;
                                            XElement elmRoot = XElement.Parse(xmlElement.OuterXml);

                                            // 找到第一個符合 SchoolYear, Semester 的節點
                                            var matched = elmRoot.Elements("History")
                                                .FirstOrDefault(e =>
                                                    (string)e.Attribute("SchoolYear") == sy.ToString() &&
                                                    (string)e.Attribute("Semester") == se.ToString());

                                            if (matched != null)
                                            {
                                                HisClassName = (string)matched.Attribute("ClassName") ?? "";
                                                HisStudentNumber = (string)matched.Attribute("StudentNumber") ?? "";

                                                // SeatNo 轉型
                                                int seatNo;
                                                if (int.TryParse((string)matched.Attribute("SeatNo"), out seatNo))
                                                    HisSeatNo = seatNo;
                                                else
                                                    HisSeatNo = null;
                                            }
                                            else
                                            {

                                            }
                                        }
                                    }
                                    catch (Exception ex)
                                    {
                                        Console.WriteLine(ex.Message);
                                    }

                                    string ScoreP = "";

                                    // 根據 makeUpScoreInfo.Detail.GetAttribute("是否取得學分") 的值，如果是 "是" 就設定 ssr.ScoreP 為 1，否則為 0。
                                    // Please use if-else syntax.
                                    if (makeUpScoreInfo.Detail.GetAttribute("是否取得學分") == "是")
                                        ScoreP = "1";
                                    else
                                        ScoreP = "0";


                                    string reScore = "-1";
                                    // 幫我寫一段 C# 程式碼，從 makeUpScoreInfo.Detail 取得「補考成績」和「修課及格標準」兩個屬性。如果「補考成績」能轉為數字，且在 0 到「修課及格標準」之間，則 reScore 等於該分數字串，否則 reScore = "-1"。如果「修課及格標準」無法轉數字，預設用 60。
                                    if (decimal.TryParse(makeUpScoreInfo.Detail.GetAttribute("補考成績"), out decimal reScoreValue2))
                                    {
                                        if (decimal.TryParse(makeUpScoreInfo.Detail.GetAttribute("修課及格標準"), out decimal passingStandard))
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
                                    string passScoreStr = makeUpScoreInfo.Detail.GetAttribute("修課及格標準");
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
                                    string courseCode = makeUpScoreInfo.Detail.GetAttribute("修課科目代碼").Trim();

                                    // 假設 makeUpScoreInfo.Detail 是 XElement 或有 GetAttribute 方法
                                    string notCountCredit = makeUpScoreInfo.Detail.GetAttribute("不計學分");
                                    string notNeedScore = makeUpScoreInfo.Detail.GetAttribute("不需評分");

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


                                    var rec = new SubjectScoreRec108
                                    {
                                        StudentID = sacRecord.StudentID,
                                        IDNumber = sacRecord.IDNumber,
                                        Birthday = sacRecord.Birthday,
                                        SchoolYear = makeUpScoreInfo.SchoolYear.ToString(),
                                        Semester = makeUpScoreInfo.Semester.ToString(),
                                        CourseCode = courseCode,
                                        SubjectName = makeUpScoreInfo.Subject.Trim(),
                                        SubjectLevel = makeUpScoreInfo.Level.Trim(),
                                        GradeYear = makeUpScoreInfo.GradeYear.ToString(),
                                        Credit = makeUpScoreInfo.CreditDec().ToString(),
                                        Score = makeUpScoreInfo.Score.ToString(),

                                        ScoreP = ScoreP,
                                        ReScore = reScore,
                                        ReScoreP = reScoreP,
                                        ScScoreType = "3",
                                        useCredit = useCredit,
                                        Text = "",
                                        Name = var.StudentName,
                                        HisClassName = HisClassName,
                                        HisSeatNo = HisSeatNo,
                                        HisStudentNumber = HisStudentNumber,
                                        ClassName = var.RefClass.ClassName,
                                        SeatNo = sacRecord.SeatNo,
                                        StudentNumber = sacRecord.StudentNumber,
                                        isScScore = true,
                                        checkPass = true,
                                        CodePass = codePass
                                    };
                                    makeUpScoreList.Add(rec);
                                }

                            }
                        }
                        else // 一般修課或空白
                        {
                            int sy = schoolyear, se = semester;

                            // 1. 檢查這一筆科目是否已被補修成績處理（避免重複）
                            if (updateSemesterSubjectScoreList.ContainsKey(sy) &&
                                updateSemesterSubjectScoreList[sy].ContainsKey(se) &&
                                updateSemesterSubjectScoreList[sy][se].ContainsKey(key))
                            {
                                continue; // 已補修，略過本次處理
                            }

                            //填入本學期科目成績
                            if (currentSubjectScoreList.ContainsValue(key))
                            {
                                #region 修改此學期已存在之成績
                                SemesterSubjectScoreInfo updateScoreInfo = null;
                                foreach (SemesterSubjectScoreInfo s in currentSubjectScoreList.Keys)
                                {
                                    if (currentSubjectScoreList[s] == key)
                                    {
                                        updateScoreInfo = s;
                                        break;
                                    }
                                }
                                sy = schoolyear;
                                se = semester;
                                if (!updateSemesterSubjectScoreList.ContainsKey(sy) || !updateSemesterSubjectScoreList[sy].ContainsKey(se) || !updateSemesterSubjectScoreList[sy][se].ContainsKey(key))
                                {
                                    //修改成績
                                    XmlElement updateScoreElement = updateScoreInfo.Detail;
                                    #region 重新填入課程資料


                                    updateScoreElement.SetAttribute("不計學分", sacRecord.NotIncludedInCredit ? "是" : "否");
                                    updateScoreElement.SetAttribute("不需評分", sacRecord.NotIncludedInCalc ? "是" : "否");
                                    updateScoreElement.SetAttribute("修課必選修", sacRecord.Required ? "必修" : "選修");
                                    updateScoreElement.SetAttribute("修課校部訂", (sacRecord.RequiredBy == "部訂" ? sacRecord.RequiredBy : "校訂"));
                                    updateScoreElement.SetAttribute("領域", sacRecord.Domain);
                                    updateScoreElement.SetAttribute("科目", sacRecord.Subject);
                                    updateScoreElement.SetAttribute("科目級別", sacRecord.SubjectLevel);
                                    updateScoreElement.SetAttribute("開課分項類別", sacRecord.Entry);

                                    updateScoreElement.SetAttribute("開課學分數", "" + sacRecord.CreditDec());

                                    if (specifySubjectNameDict.ContainsKey(sacRecord.StudentID))
                                    {
                                        string sKey = sacRecord.Subject + "_" + sacRecord.SubjectLevel;
                                        if (specifySubjectNameDict[sacRecord.StudentID].ContainsKey(sKey))
                                            updateScoreElement.SetAttribute("指定學年科目名稱", specifySubjectNameDict[sacRecord.StudentID][sKey]);
                                    }


                                    #endregion

                                    // 沒有修課成績填空
                                    if (sacRecord.HasFinalScore)
                                        updateScoreElement.SetAttribute("原始成績", ("" + GetRoundScore(sacRecord.FinalScore, decimals, mode)));
                                    else
                                        updateScoreElement.SetAttribute("原始成績", "");


                                    // 當有直接指定總成績覆蓋
                                    if (studentFinalScoreDict.ContainsKey(sacRecord.StudentID))
                                    {
                                        string sKey = sacRecord.Subject + "_" + sacRecord.SubjectLevel;

                                        if (studentFinalScoreDict[sacRecord.StudentID].ContainsKey(sKey))
                                        {
                                            DataRow dr = studentFinalScoreDict[sacRecord.StudentID][sKey];

                                            string passing_standard = "", makeup_standard = "", remark = "", designate_final_score = "", subject_code = "";

                                            decimal passing_standard_score, makeup_standard_score, designate_final_score_score;

                                            if (dr["passing_standard"] != null)
                                                passing_standard = dr["passing_standard"].ToString();

                                            if (dr["makeup_standard"] != null)
                                                makeup_standard = dr["makeup_standard"].ToString();

                                            if (dr["remark"] != null)
                                                remark = dr["remark"].ToString();

                                            if (dr["subject_code"] != null)
                                                subject_code = dr["subject_code"].ToString();

                                            if (dr["designate_final_score"] != null)
                                                designate_final_score = dr["designate_final_score"].ToString();

                                            if (decimal.TryParse(passing_standard, out passing_standard_score))
                                                updateScoreElement.SetAttribute("修課及格標準", ("" + GetRoundScore(passing_standard_score, decimals, mode)));
                                            else
                                                updateScoreElement.SetAttribute("修課及格標準", "");

                                            if (decimal.TryParse(makeup_standard, out makeup_standard_score))
                                                updateScoreElement.SetAttribute("修課補考標準", ("" + GetRoundScore(makeup_standard_score, decimals, mode)));
                                            else
                                                updateScoreElement.SetAttribute("修課補考標準", "");

                                            updateScoreElement.SetAttribute("註記", "");

                                            if (decimal.TryParse(designate_final_score, out designate_final_score_score))
                                            {
                                                updateScoreElement.SetAttribute("修課直接指定總成績", ("" + GetRoundScore(designate_final_score_score, decimals, mode)));

                                                // 註解是因經過2024/4/26討論，修課直接指定總成績不應該覆蓋原始成績，需要保留原始成績。                                              
                                                updateScoreElement.SetAttribute("原始成績", ("" + GetRoundScore(designate_final_score_score, decimals, mode)));


                                                //updateScoreElement.SetAttribute("原始成績", (sacRecord.NotIncludedInCalc ? "" : "" + GetRoundScore(designate_final_score_score, decimals, mode)));


                                                updateScoreElement.SetAttribute("註記", "修課成績：" + ("" + GetRoundScore(sacRecord.FinalScore, decimals, mode)));
                                            }
                                            else
                                            {
                                                updateScoreElement.SetAttribute("修課直接指定總成績", "");
                                            }

                                            updateScoreElement.SetAttribute("修課備註", remark);
                                            updateScoreElement.SetAttribute("修課科目代碼", subject_code);
                                        }
                                    }

                                    //做取得學分判斷
                                    #region 做取得學分判斷及填入擇優採計成績
                                    //最高分
                                    decimal maxScore = 0;// sacRecord.FinalScore;
                                    #region 抓最高分
                                    string[] scoreNames = new string[] { "原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績" };
                                    foreach (string scorename in scoreNames)
                                    {
                                        decimal s;
                                        if (decimal.TryParse(updateScoreElement.GetAttribute(scorename), out s))
                                        {
                                            if (s > maxScore)
                                            {
                                                maxScore = s;
                                            }
                                        }
                                    }
                                    #endregion

                                    //如果有擇優採計成績且重讀學期有修過課
                                    if (duplicateSubjectLevelMethodDict_Afterfilter.ContainsKey(var.StudentID + "_" + key) ? duplicateSubjectLevelMethodDict_Afterfilter[var.StudentID + "_" + key] == "重讀(擇優採計成績)" : false)
                                    {
                                        #region 填入擇優採計成績
                                        //foreach (SemesterSubjectScoreInfo s in repeatSubjectScoreList.Keys)
                                        //{
                                        //    //之前的成績比現在的成績好
                                        //    if (repeatSubjectScoreList[s] == key && s.Score > maxScore)
                                        //    {
                                        //        updateScoreElement.SetAttribute("原始成績", "" + GetRoundScore(s.Score, decimals, mode));
                                        //        updateScoreElement.SetAttribute("註記", "修課成績：" + sacRecord.FinalScore);
                                        //        maxScore = s.Score;
                                        //    }
                                        //}
                                        #endregion
                                    }
                                    decimal passscore;
                                    //if (!applyLimit.ContainsKey(updateScoreInfo.GradeYear))
                                    //    passscore = 60;
                                    //else
                                    //    passscore = applyLimit[updateScoreInfo.GradeYear];
                                    // 新寫及格標準
                                    passscore = 100;
                                    if (studentPassScoreDict.ContainsKey(var.StudentID))
                                    {
                                        if (studentPassScoreDict[var.StudentID].ContainsKey(key))
                                        {
                                            passscore = studentPassScoreDict[var.StudentID][key];
                                        }
                                        else
                                        {
                                            if (!applyLimit.ContainsKey(updateScoreInfo.GradeYear))
                                                passscore = 60;
                                            else
                                                passscore = applyLimit[updateScoreInfo.GradeYear];
                                        }
                                    }

                                    updateScoreElement.SetAttribute("是否取得學分", (sacRecord.NotIncludedInCalc || maxScore >= passscore) ? "是" : "否");

                                    // 2024/7/5 會議決議，需要計算學分使用成績判斷是否取得學分
                                    if (sacRecord.NotIncludedInCredit == false)
                                    {
                                        updateScoreElement.SetAttribute("是否取得學分", maxScore >= passscore ? "是" : "否");
                                    }

                                    #endregion
                                    if (!updateSemesterSubjectScoreList.ContainsKey(sy)) updateSemesterSubjectScoreList.Add(sy, new Dictionary<int, Dictionary<string, XmlElement>>());
                                    if (!updateSemesterSubjectScoreList[sy].ContainsKey(se)) updateSemesterSubjectScoreList[sy].Add(se, new Dictionary<string, XmlElement>());
                                    updateSemesterSubjectScoreList[sy][se].Add(key, updateScoreElement);
                                }
                                #endregion
                            }
                            else
                            {
                                #region 新增一筆成績
                                sy = schoolyear; se = semester;
                                if (!insertSemesterSubjectScoreList.ContainsKey(sy) || !insertSemesterSubjectScoreList[sy].ContainsKey(se) || !insertSemesterSubjectScoreList[sy][se].ContainsKey(key))
                                {
                                    //科目名稱空白有錯誤時執行，2024/5/23討論，當課程的科目空白不計算
                                    if (string.IsNullOrEmpty(sacRecord.Subject) || string.IsNullOrWhiteSpace(sacRecord.Subject))
                                    {
                                        //_WarningList為空就先建立
                                        if (_WarningList == null) _WarningList = new List<string>();
                                        //_WarningList不存在此課程名稱才加入
                                        if (!_WarningList.Contains(sacRecord.CourseName))
                                        {
                                            _WarningList.Add(sacRecord.CourseName);
                                        }
                                        continue;
                                    }


                                    #region 加入新的資料
                                    XmlElement newScoreInfo = doc.CreateElement("Subject");
                                    newScoreInfo.SetAttribute("不計學分", sacRecord.NotIncludedInCredit ? "是" : "否");
                                    newScoreInfo.SetAttribute("不需評分", sacRecord.NotIncludedInCalc ? "是" : "否");
                                    newScoreInfo.SetAttribute("修課必選修", sacRecord.Required ? "必修" : "選修");
                                    newScoreInfo.SetAttribute("修課校部訂", (sacRecord.RequiredBy == "部訂" ? sacRecord.RequiredBy : "校訂"));
                                    newScoreInfo.SetAttribute("領域", sacRecord.Domain);
                                    newScoreInfo.SetAttribute("科目", sacRecord.Subject);
                                    newScoreInfo.SetAttribute("科目級別", sacRecord.SubjectLevel);
                                    newScoreInfo.SetAttribute("開課分項類別", sacRecord.Entry);
                                    newScoreInfo.SetAttribute("開課學分數", "" + sacRecord.CreditDec());

                                    // 沒有修課成績填空
                                    if (sacRecord.HasFinalScore)
                                        newScoreInfo.SetAttribute("原始成績", ("" + GetRoundScore(sacRecord.FinalScore, decimals, mode)));
                                    else
                                        newScoreInfo.SetAttribute("原始成績", "");


                                    if (specifySubjectNameDict.ContainsKey(sacRecord.StudentID))
                                    {
                                        string sKey = sacRecord.Subject + "_" + sacRecord.SubjectLevel;
                                        if (specifySubjectNameDict[sacRecord.StudentID].ContainsKey(sKey))
                                            newScoreInfo.SetAttribute("指定學年科目名稱", specifySubjectNameDict[sacRecord.StudentID][sKey]);
                                    }

                                    // 當有直接指定總成績覆蓋
                                    if (studentFinalScoreDict.ContainsKey(sacRecord.StudentID))
                                    {
                                        string sKey = sacRecord.Subject + "_" + sacRecord.SubjectLevel;

                                        if (studentFinalScoreDict[sacRecord.StudentID].ContainsKey(sKey))
                                        {
                                            DataRow dr = studentFinalScoreDict[sacRecord.StudentID][sKey];

                                            string passing_standard = "", makeup_standard = "", remark = "", designate_final_score = "", subject_code = "";

                                            decimal passing_standard_score, makeup_standard_score, designate_final_score_score;

                                            if (dr["passing_standard"] != null)
                                                passing_standard = dr["passing_standard"].ToString();

                                            if (dr["makeup_standard"] != null)
                                                makeup_standard = dr["makeup_standard"].ToString();

                                            if (dr["remark"] != null)
                                                remark = dr["remark"].ToString();

                                            if (dr["subject_code"] != null)
                                                subject_code = dr["subject_code"].ToString();

                                            if (dr["designate_final_score"] != null)
                                                designate_final_score = dr["designate_final_score"].ToString();

                                            if (decimal.TryParse(passing_standard, out passing_standard_score))
                                                newScoreInfo.SetAttribute("修課及格標準", ("" + GetRoundScore(passing_standard_score, decimals, mode)));
                                            else
                                                newScoreInfo.SetAttribute("修課及格標準", "");

                                            if (decimal.TryParse(makeup_standard, out makeup_standard_score))
                                                newScoreInfo.SetAttribute("修課補考標準", ("" + GetRoundScore(makeup_standard_score, decimals, mode)));
                                            else
                                                newScoreInfo.SetAttribute("修課補考標準", "");

                                            if (decimal.TryParse(designate_final_score, out designate_final_score_score))
                                            {
                                                newScoreInfo.SetAttribute("修課直接指定總成績", ("" + GetRoundScore(designate_final_score_score, decimals, mode)));

                                                // 註解是因經過2024/4/26討論，修課直接指定總成績不應該覆蓋原始成績，需要保留原始成績。

                                                newScoreInfo.SetAttribute("原始成績", ("" + GetRoundScore(designate_final_score_score, decimals, mode)));


                                                newScoreInfo.SetAttribute("註記", "修課成績：" + ("" + GetRoundScore(sacRecord.FinalScore, decimals, mode)));
                                            }
                                            else
                                            {
                                                newScoreInfo.SetAttribute("修課直接指定總成績", "");
                                            }

                                            newScoreInfo.SetAttribute("修課備註", remark);
                                            newScoreInfo.SetAttribute("修課科目代碼", subject_code);
                                        }
                                    }

                                    newScoreInfo.SetAttribute("重修成績", "");
                                    newScoreInfo.SetAttribute("學年調整成績", "");
                                    newScoreInfo.SetAttribute("擇優採計成績", "");
                                    newScoreInfo.SetAttribute("補考成績", "");
                                    //做取得學分判斷
                                    #region 做取得學分判斷及填入擇優採計成績
                                    //最高分
                                    decimal maxScore = 0;// = sacRecord.FinalScore;
                                    #region 抓最高分
                                    string[] scoreNames = new string[] { "原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績" };
                                    foreach (string scorename in scoreNames)
                                    {
                                        decimal s;
                                        if (decimal.TryParse(newScoreInfo.GetAttribute(scorename), out s))
                                        {
                                            if (s > maxScore)
                                            {
                                                maxScore = s;
                                            }
                                        }
                                    }
                                    #endregion

                                    //如果有擇優採計成績且重讀學期有修過課
                                    if (duplicateSubjectLevelMethodDict_Afterfilter.ContainsKey(var.StudentID + "_" + key) ? duplicateSubjectLevelMethodDict_Afterfilter[var.StudentID + "_" + key] == "重讀(擇優採計成績)" : false)
                                    {
                                        #region 填入擇優採計成績
                                        //foreach (SemesterSubjectScoreInfo s in repeatSubjectScoreList.Keys)
                                        //{
                                        //    //之前的成績比現在的成績好
                                        //    if (repeatSubjectScoreList[s] == key && s.Score > maxScore)
                                        //    {
                                        //        //newScoreInfo.SetAttribute("擇優採計成績", "" + GetRoundScore(s.Score, decimals, mode));
                                        //        newScoreInfo.SetAttribute("原始成績", "" + GetRoundScore(s.Score, decimals, mode));
                                        //        newScoreInfo.SetAttribute("註記", "修課成績：" + sacRecord.FinalScore);
                                        //        maxScore = s.Score;
                                        //    }
                                        //}
                                        #endregion
                                    }
                                    decimal passscore;
                                    //if (!applyLimit.ContainsKey((int)gradeYear))
                                    //    passscore = 60;
                                    //else
                                    //    passscore = applyLimit[(int)gradeYear];
                                    // 新寫及格標準
                                    passscore = 100;
                                    if (studentPassScoreDict.ContainsKey(var.StudentID))
                                    {
                                        if (studentPassScoreDict[var.StudentID].ContainsKey(key))
                                        {
                                            passscore = studentPassScoreDict[var.StudentID][key];
                                        }
                                        else
                                        {
                                            if (!applyLimit.ContainsKey((int)gradeYear))
                                                passscore = 60;
                                            else
                                                passscore = applyLimit[(int)gradeYear];
                                        }
                                    }


                                    #endregion
                                    newScoreInfo.SetAttribute("是否取得學分", (sacRecord.NotIncludedInCalc || maxScore >= passscore) ? "是" : "否");

                                    // 2024/7/5 會議決議，需要計算學分使用成績判斷是否取得學分
                                    if (sacRecord.NotIncludedInCredit == false)
                                    {
                                        newScoreInfo.SetAttribute("是否取得學分", maxScore >= passscore ? "是" : "否");
                                    }

                                    #endregion
                                    if (!insertSemesterSubjectScoreList.ContainsKey(sy)) insertSemesterSubjectScoreList.Add(sy, new Dictionary<int, Dictionary<string, XmlElement>>());
                                    if (!insertSemesterSubjectScoreList[sy].ContainsKey(se)) insertSemesterSubjectScoreList[sy].Add(se, new Dictionary<string, XmlElement>());
                                    insertSemesterSubjectScoreList[sy][se].Add(key, newScoreInfo);
                                }
                                #endregion
                            }
                        }


                    }

                    #endregion
                    #region 從新增跟修改清單中產生變動資料
                    //抓取暗藏的學生學期成績編號資料
                    Dictionary<int, Dictionary<int, string>> semeScoreID = (Dictionary<int, Dictionary<int, string>>)var.Fields["SemesterSubjectScoreID"];
                    foreach (int sy in updateSemesterSubjectScoreList.Keys)
                    {
                        foreach (int se in updateSemesterSubjectScoreList[sy].Keys)
                        {
                            List<string> appendedKeys = new List<string>();
                            XmlElement parentNode;
                            parentNode = doc.CreateElement("UpdateSemesterScore");
                            parentNode.SetAttribute("ID", semeScoreID[sy][se]);
                            if (sy == schoolyear && se == semester)
                                parentNode.SetAttribute("GradeYear", "" + gradeYear);
                            //加入修改過的成績資料
                            foreach (string key in updateSemesterSubjectScoreList[sy][se].Keys)
                            {
                                appendedKeys.Add(key);
                                parentNode.AppendChild(doc.ImportNode(updateSemesterSubjectScoreList[sy][se][key], true));
                            }
                            if (insertSemesterSubjectScoreList.ContainsKey(sy) && insertSemesterSubjectScoreList[sy].ContainsKey(se))
                            {
                                //加入新增的成績資料
                                foreach (string key in insertSemesterSubjectScoreList[sy][se].Keys)
                                {
                                    appendedKeys.Add(key);
                                    parentNode.AppendChild(doc.ImportNode(insertSemesterSubjectScoreList[sy][se][key], true));
                                }
                                insertSemesterSubjectScoreList[sy].Remove(se);
                                if (insertSemesterSubjectScoreList[sy].Count == 0)
                                    insertSemesterSubjectScoreList.Remove(sy);
                            }
                            if (semesterSubjectScoreList.ContainsKey(sy) && semesterSubjectScoreList[sy].ContainsKey(se))
                            {
                                //加入此學期沒有變動的成績資料
                                foreach (string key in semesterSubjectScoreList[sy][se].Keys)
                                {
                                    if (!appendedKeys.Contains(key))
                                        parentNode.AppendChild(doc.ImportNode(semesterSubjectScoreList[sy][se][key].Detail, true));
                                }
                            }
                            semesterSubjectCalcScoreElement.AppendChild(parentNode);
                        }
                    }
                    //如果還有新增成績，必定為此學期之成績，且此學期尚無任何成績紀錄
                    foreach (int sy in insertSemesterSubjectScoreList.Keys)
                    {
                        foreach (int se in insertSemesterSubjectScoreList[sy].Keys)
                        {
                            if (insertSemesterSubjectScoreList[sy][se].Count > 0)
                            {
                                XmlElement parentNode;
                                if (semeScoreID.ContainsKey(sy) && semeScoreID[sy].ContainsKey(se))
                                {
                                    parentNode = doc.CreateElement("UpdateSemesterScore");
                                    parentNode.SetAttribute("ID", semeScoreID[sy][se]);
                                    if (sy == schoolyear && se == semester)
                                        parentNode.SetAttribute("GradeYear", "" + gradeYear);
                                }
                                else
                                {
                                    parentNode = doc.CreateElement("InsertSemesterScore");
                                    parentNode.SetAttribute("GradeYear", "" + gradeYear);
                                    parentNode.SetAttribute("SchoolYear", "" + sy);
                                    parentNode.SetAttribute("Semester", "" + se);
                                }
                                foreach (XmlElement ele in insertSemesterSubjectScoreList[sy][se].Values)
                                {
                                    parentNode.AppendChild(doc.ImportNode(ele, true));
                                }
                                //如果此學期有新增成績卻沒有更新成績，需將原無變動之成績填回
                                if ((!updateSemesterSubjectScoreList.ContainsKey(sy) || !updateSemesterSubjectScoreList[sy].ContainsKey(se)) && semesterSubjectScoreList.ContainsKey(sy) && semesterSubjectScoreList[sy].ContainsKey(se))
                                {
                                    //加入此學期沒有變動的成績資料
                                    foreach (string key in semesterSubjectScoreList[sy][se].Keys)
                                    {
                                        parentNode.AppendChild(doc.ImportNode(semesterSubjectScoreList[sy][se][key].Detail, true));
                                    }
                                }
                                semesterSubjectCalcScoreElement.AppendChild(parentNode);
                            }
                        }
                    }
                    #endregion

                }
                var.Fields.Add("SemesterSubjectCalcScore", semesterSubjectCalcScoreElement);
            }

            // 補修成績寫入學期歷程
            if (makeUpScoreList.Count > 0)
            {
                LearningHistoryDataAccess learningHistoryDataAccess = new LearningHistoryDataAccess();
                learningHistoryDataAccess.SaveScores43(makeUpScoreList, schoolyear, semester);
            }

            // 重修學期寫入學期歷程
            if (restudyScoreList.Count > 0)
            {
                LearningHistoryDataAccess learningHistoryDataAccess = new LearningHistoryDataAccess();
                learningHistoryDataAccess.SaveScores52(restudyScoreList, schoolyear, semester);
            }

            // 轉學轉科寫入學期歷程
            if (transferScoreList.Count > 0)
            {
                LearningHistoryDataAccess learningHistoryDataAccess = new LearningHistoryDataAccess();
                learningHistoryDataAccess.SaveScores44(transferScoreList, schoolyear, semester);
            }

            // 重讀學期寫入學期歷程
            if (repeatScoreList.Count > 0)
            {
                LearningHistoryDataAccess learningHistoryDataAccess = new LearningHistoryDataAccess();
                learningHistoryDataAccess.SaveScores53(repeatScoreList, schoolyear, semester);
            }

            try
            {
                // 呼叫學期歷程同步
                FISCA.Features.Invoke("StudentLearningHistoryDetailContent");
            }
            catch (Exception ex)
            {
                Console.WriteLine("StudentLearningHistoryDetailContent 無法呼叫：" + ex.Message);
            }
            return _ErrorList;
        }

        /// <summary>
        /// 計算學期分項成績
        /// </summary>
        /// <param name="schoolyear">學年度</param>
        /// <param name="semester">學期</param>
        /// <param name="accesshelper"></param>
        /// <param name="students">學生物件列表</param>
        /// <returns></returns>
        public Dictionary<StudentRecord, List<string>> FillSemesterEntryCalcScore(int schoolyear, int semester, AccessHelper accesshelper, List<StudentRecord> students)
        {
            Dictionary<StudentRecord, List<string>> _ErrorList = new Dictionary<StudentRecord, List<string>>();

            accesshelper.StudentHelper.FillSemesterSubjectScore(false, students);
            foreach (StudentRecord var in students)
            {
                Dictionary<string, decimal> entryCreditCount = new Dictionary<string, decimal>();
                Dictionary<string, List<decimal>> entrySubjectScores = new Dictionary<string, List<decimal>>();
                Dictionary<string, decimal> entryDividend = new Dictionary<string, decimal>();
                Dictionary<string, bool> calcEntry = new Dictionary<string, bool>();
                Dictionary<string, bool> calcInStudy = new Dictionary<string, bool>();
                List<string> takeScore = new List<string>();
                bool takeRepairScore = false;
                //精準位數
                int decimals = 2;
                //進位模式
                RoundMode mode = RoundMode.四捨五入;
                //成績年級及計算規則皆存在，允許計算成績
                bool canCalc = true;
                #region 取得成績年級跟計算規則
                {
                    #region 處理計算規則
                    XmlElement scoreCalcRule = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID).ScoreCalcRuleElement;
                    if (scoreCalcRule == null)
                    {
                        if (!_ErrorList.ContainsKey(var))
                            _ErrorList.Add(var, new List<string>());
                        _ErrorList[var].Add("沒有設定成績計算規則。");
                        canCalc &= false;
                    }
                    else
                    {
                        DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                        bool tryParsebool;
                        int tryParseint;
                        decimal tryParseDecimal;

                        #region 精準位數
                        if (scoreCalcRule.SelectSingleNode("各項成績計算位數/學期分項成績計算位數") != null)
                        {
                            if (int.TryParse(helper.GetText("各項成績計算位數/學期分項成績計算位數/@位數"), out tryParseint))
                                decimals = tryParseint;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/學期分項成績計算位數/@四捨五入"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.四捨五入;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/學期分項成績計算位數/@無條件捨去"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.無條件捨去;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/學期分項成績計算位數/@無條件進位"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.無條件進位;
                        }
                        #endregion
                        #region 計算類別
                        foreach (string entry in new string[] { "體育", "學業", "國防通識", "健康與護理", "實習科目", "專業科目" })
                        {
                            if (scoreCalcRule.SelectSingleNode("分項成績計算項目") == null || scoreCalcRule.SelectSingleNode("分項成績計算項目/" + entry) == null || ((XmlElement)scoreCalcRule.SelectSingleNode("分項成績計算項目/" + entry)).GetAttribute("計算成績") == "True")
                                calcEntry.Add(entry, true);
                            else
                                calcEntry.Add(entry, false);

                            // 2014/3/20，修改當沒有勾選 預設併入學期學業成績 ChenCT
                            if (scoreCalcRule.SelectSingleNode("分項成績計算項目") == null || scoreCalcRule.SelectSingleNode("分項成績計算項目/" + entry) == null || ((XmlElement)scoreCalcRule.SelectSingleNode("分項成績計算項目/" + entry)).GetAttribute("併入學期學業成績") == "True")
                                calcInStudy.Add(entry, true);
                            else
                                calcInStudy.Add(entry, false);
                        }
                        #endregion
                        #region 採計成績欄位
                        // 處理補修成績
                        bool.TryParse(helper.GetText("分項成績計算採計成績欄位/@補修成績"), out takeRepairScore);

                        foreach (string item in new string[] { "原始成績", "補考成績", "重修成績", "擇優採計成績", "學年調整成績" })
                        {
                            if (!bool.TryParse(helper.GetText("分項成績計算採計成績欄位/@" + item), out tryParsebool) || tryParsebool)
                            {//沒有設定這項成績設定規則(預設true)或者設定值是true
                                takeScore.Add(item);
                            }
                        }

                        #endregion
                    }
                    #endregion
                }
                #endregion
                Dictionary<string, decimal> entryScores = new Dictionary<string, decimal>();
                if (canCalc)
                {
                    #region 將成績分到各分項類別中
                    foreach (SemesterSubjectScoreInfo subjectNode in var.SemesterSubjectScoreList)
                    {
                        if (subjectNode.SchoolYear == schoolyear && subjectNode.Semester == semester)
                        {
                            // 2024/3/4 討論結果，不需評分 ="是"，不計算
                            if (subjectNode.Detail.GetAttribute("不需評分") == "是")
                                continue;

                            ////不計學分或不需評分不用算
                            //if (subjectNode.Detail.GetAttribute("不需評分") == "是" || subjectNode.Detail.GetAttribute("不計學分") == "是")
                            //    continue;

                            // 若為補修成績且補修成績不採計，則不計算
                            if (takeRepairScore == false)
                            {
                                if (subjectNode.Detail.GetAttribute("是否補修成績") == "是")
                                    continue;
                            }

                            //string subjectCode = subjectNode.Detail.GetAttribute("修課科目代碼");
                            //if (subjectCode.Length >= 23) //共23碼
                            //{
                            //    if (subjectCode[16].ToString() + subjectCode[18].ToString() == "9D" || subjectCode[16].ToString() + subjectCode[18].ToString() == "9d")
                            //        continue;
                            //}

                            #region 分項類別跟學分數
                            string entry = subjectNode.Detail.GetAttribute("開課分項類別");
                            decimal credit = subjectNode.CreditDec();
                            #endregion
                            decimal tryParseDecimal;
                            decimal maxScore = 0;
                            decimal original = 0;
                            if (decimal.TryParse(subjectNode.Detail.GetAttribute("原始成績"), out tryParseDecimal))
                                original = tryParseDecimal;
                            #region 取得最高分數
                            foreach (var item in takeScore)
                            {
                                if (decimal.TryParse(subjectNode.Detail.GetAttribute(item), out tryParseDecimal) && maxScore < tryParseDecimal)
                                    maxScore = tryParseDecimal;
                            }
                            //if (decimal.TryParse(subjectNode.Detail.GetAttribute("原始成績"), out tryParseDecimal))
                            //    maxScore = tryParseDecimal;
                            //if (decimal.TryParse(subjectNode.Detail.GetAttribute("學年調整成績"), out tryParseDecimal) && maxScore < tryParseDecimal)
                            //    maxScore = tryParseDecimal;
                            //if (decimal.TryParse(subjectNode.Detail.GetAttribute("擇優採計成績"), out tryParseDecimal) && maxScore < tryParseDecimal)
                            //    maxScore = tryParseDecimal;
                            //if (decimal.TryParse(subjectNode.Detail.GetAttribute("補考成績"), out tryParseDecimal) && maxScore < tryParseDecimal)
                            //    maxScore = tryParseDecimal;
                            //if (decimal.TryParse(subjectNode.Detail.GetAttribute("重修成績"), out tryParseDecimal) && maxScore < tryParseDecimal)
                            //    maxScore = tryParseDecimal;
                            #endregion
                            switch (entry)
                            {
                                case "體育":
                                case "國防通識":
                                case "健康與護理":
                                case "實習科目":
                                case "專業科目":
                                    //計算分項成績
                                    if (calcEntry[entry])
                                    {
                                        #region original
                                        //加總學分數
                                        if (!entryCreditCount.ContainsKey(entry + "(原始)"))
                                            entryCreditCount.Add(entry + "(原始)", credit);
                                        else
                                            entryCreditCount[entry + "(原始)"] += credit;
                                        //加入將成績資料分項
                                        if (!entrySubjectScores.ContainsKey(entry + "(原始)")) entrySubjectScores.Add(entry + "(原始)", new List<decimal>());
                                        entrySubjectScores[entry + "(原始)"].Add(original);
                                        //加權總計
                                        if (!entryDividend.ContainsKey(entry + "(原始)"))
                                            entryDividend.Add(entry + "(原始)", original * credit);
                                        else
                                            entryDividend[entry + "(原始)"] += (original * credit);
                                        #endregion
                                        #region maxScore
                                        //加總學分數
                                        if (!entryCreditCount.ContainsKey(entry))
                                            entryCreditCount.Add(entry, credit);
                                        else
                                            entryCreditCount[entry] += credit;
                                        //加入將成績資料分項
                                        if (!entrySubjectScores.ContainsKey(entry)) entrySubjectScores.Add(entry, new List<decimal>());
                                        entrySubjectScores[entry].Add(maxScore);
                                        //加權總計
                                        if (!entryDividend.ContainsKey(entry))
                                            entryDividend.Add(entry, maxScore * credit);
                                        else
                                            entryDividend[entry] += (maxScore * credit);
                                        #endregion
                                    }
                                    //將科目成績與學業成績一併計算
                                    if (calcInStudy[entry])
                                    {
                                        #region original
                                        //加總學分數
                                        if (!entryCreditCount.ContainsKey("學業" + "(原始)"))
                                            entryCreditCount.Add("學業" + "(原始)", credit);
                                        else
                                            entryCreditCount["學業" + "(原始)"] += credit;
                                        //加入將成績資料分項
                                        if (!entrySubjectScores.ContainsKey("學業" + "(原始)")) entrySubjectScores.Add("學業" + "(原始)", new List<decimal>());
                                        entrySubjectScores["學業" + "(原始)"].Add(original);
                                        //加權總計
                                        if (!entryDividend.ContainsKey("學業" + "(原始)"))
                                            entryDividend.Add("學業" + "(原始)", original * credit);
                                        else
                                            entryDividend["學業" + "(原始)"] += (original * credit);
                                        #endregion
                                        #region maxScore
                                        //加總學分數
                                        if (!entryCreditCount.ContainsKey("學業"))
                                            entryCreditCount.Add("學業", credit);
                                        else
                                            entryCreditCount["學業"] += credit;
                                        //加入將成績資料分項
                                        if (!entrySubjectScores.ContainsKey("學業")) entrySubjectScores.Add("學業", new List<decimal>());
                                        entrySubjectScores["學業"].Add(maxScore);
                                        //加權總計
                                        if (!entryDividend.ContainsKey("學業"))
                                            entryDividend.Add("學業", maxScore * credit);
                                        else
                                            entryDividend["學業"] += (maxScore * credit);
                                        #endregion
                                    }
                                    break;

                                case "學業":
                                default:
                                    #region original
                                    //加總學分數
                                    if (!entryCreditCount.ContainsKey("學業" + "(原始)"))
                                        entryCreditCount.Add("學業" + "(原始)", credit);
                                    else
                                        entryCreditCount["學業" + "(原始)"] += credit;
                                    //加入將成績資料分項
                                    if (!entrySubjectScores.ContainsKey("學業" + "(原始)")) entrySubjectScores.Add("學業" + "(原始)", new List<decimal>());
                                    entrySubjectScores["學業" + "(原始)"].Add(original);
                                    //加權總計
                                    if (!entryDividend.ContainsKey("學業" + "(原始)"))
                                        entryDividend.Add("學業" + "(原始)", original * credit);
                                    else
                                        entryDividend["學業" + "(原始)"] += (original * credit);
                                    #endregion
                                    #region maxScore
                                    //加總學分數
                                    if (!entryCreditCount.ContainsKey("學業"))
                                        entryCreditCount.Add("學業", credit);
                                    else
                                        entryCreditCount["學業"] += credit;
                                    //加入將成績資料分項
                                    if (!entrySubjectScores.ContainsKey("學業")) entrySubjectScores.Add("學業", new List<decimal>());
                                    entrySubjectScores["學業"].Add(maxScore);
                                    //加權總計
                                    if (!entryDividend.ContainsKey("學業"))
                                        entryDividend.Add("學業", maxScore * credit);
                                    else
                                        entryDividend["學業"] += (maxScore * credit);
                                    #endregion
                                    break;
                            }
                        }
                    }
                    #endregion
                    #region 處理計算各分項類別的成績
                    foreach (string entry in entryCreditCount.Keys)
                    {
                        decimal entryScore = 0;
                        #region 計算entryScore
                        if (entryCreditCount[entry] == 0)
                        {
                            foreach (decimal score in entrySubjectScores[entry])
                            {
                                entryScore += score;
                            }
                            entryScore = (entryScore / entrySubjectScores[entry].Count);
                        }
                        else
                        {
                            //用加權總分除學分數
                            entryScore = (entryDividend[entry] / entryCreditCount[entry]);
                        }
                        #endregion
                        //精準位數處理
                        entryScore = GetRoundScore(entryScore, decimals, mode);
                        #region 填入EntryScores
                        entryScores.Add(entry, entryScore);
                        #endregion
                    }
                    #endregion
                }
                if (var.Fields.ContainsKey("CalcEntryScores"))
                    var.Fields["CalcEntryScores"] = entryScores;
                else
                    var.Fields.Add("CalcEntryScores", entryScores);
            }
            return _ErrorList;
        }

        /// <summary>
        /// 計算學年分項成績（依據學期科目成績計算）
        /// </summary>
        /// <param name="schoolyear">學年度</param>
        /// <param name="accesshelper"></param>
        /// <param name="students">學生物件列表</param>
        /// <returns></returns>
        public Dictionary<StudentRecord, List<string>> FillSchoolYearEntryCalcScoreBySchoolYearSubjectScore(int schoolyear, AccessHelper accesshelper, List<StudentRecord> students)
        {
            Dictionary<StudentRecord, List<string>> _ErrorList = new Dictionary<StudentRecord, List<string>>();
            //取得學期科目成績
            accesshelper.StudentHelper.FillSemesterSubjectScore(false, students);

            //取得學年科目成績
            accesshelper.StudentHelper.FillSchoolYearSubjectScore(false, students);

            foreach (StudentRecord var in students)
            {
                //計算結果
                Dictionary<string, decimal> entryCalcScores = new Dictionary<string, decimal>();
                //精準位數
                int decimals = 2;
                //進位模式
                RoundMode mode = RoundMode.四捨五入;
                //成績年級及計算規則皆存在，允許計算成績
                bool canCalc = true;
                Dictionary<string, bool> calcEntry = new Dictionary<string, bool>();
                Dictionary<string, bool> calcInStudy = new Dictionary<string, bool>();

                #region 取得成績年級跟計算規則
                {
                    #region 處理計算規則
                    XmlElement scoreCalcRule = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID).ScoreCalcRuleElement;
                    if (scoreCalcRule == null)
                    {
                        if (!_ErrorList.ContainsKey(var))
                            _ErrorList.Add(var, new List<string>());
                        _ErrorList[var].Add("沒有設定成績計算規則。");
                        canCalc &= false;
                    }
                    else
                    {
                        DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                        bool tryParsebool;
                        int tryParseint;

                        #region 精準位數
                        if (scoreCalcRule.SelectSingleNode("各項成績計算位數/學年分項成績計算位數") != null)
                        {
                            if (int.TryParse(helper.GetText("各項成績計算位數/學年分項成績計算位數/@位數"), out tryParseint))
                                decimals = tryParseint;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/學年分項成績計算位數/@四捨五入"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.四捨五入;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/學年分項成績計算位數/@無條件捨去"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.無條件捨去;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/學年分項成績計算位數/@無條件進位"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.無條件進位;
                        }
                        #endregion
                        #region 計算類別
                        //學年科目成績沒有分項屬性，此部分暫不考量
                        //foreach (string entry in new string[] { "體育", "學業", "國防通識", "健康與護理", "實習科目" })
                        //{
                        //    if (scoreCalcRule.SelectSingleNode("分項成績計算項目") == null || scoreCalcRule.SelectSingleNode("分項成績計算項目/" + entry) == null || ((XmlElement)scoreCalcRule.SelectSingleNode("分項成績計算項目/" + entry)).GetAttribute("計算成績") == "True")
                        //        calcEntry.Add(entry, true);
                        //    else
                        //        calcEntry.Add(entry, false);

                        //    if (scoreCalcRule.SelectSingleNode("分項成績計算項目") == null || scoreCalcRule.SelectSingleNode("分項成績計算項目/" + entry) == null || ((XmlElement)scoreCalcRule.SelectSingleNode("分項成績計算項目/" + entry)).GetAttribute("併入學期學業成績") != "True")
                        //        calcInStudy.Add(entry, false);
                        //    else
                        //        calcInStudy.Add(entry, true);
                        //}
                        #endregion
                    }
                    #endregion
                }
                #endregion
                if (canCalc)
                {
                    //根據學年度取得學年科目成績的成績年級
                    int? gradeyear = var.SchoolYearSubjectScoreList.GetGradeYear(schoolyear);

                    if (gradeyear != null)
                    {
                        //根據學年度及成績年級刪除不需要學年科目成績
                        var.SchoolYearSubjectScoreList.FilterSchoolYearSubjectScore(gradeyear, schoolyear);

                        //根據學年度及成績年級刪除不需要學期科目成績
                        var.SemesterSubjectScoreList.FilterSemesterSubjectScore(gradeyear, schoolyear);

                        entryCalcScores = var.CalculateSchoolYearEntryScore(calcEntry, calcInStudy, schoolyear, gradeyear.Value, mode, decimals);
                    }
                }
                if (var.Fields.ContainsKey("CalcSchoolYearEntryScores"))
                    var.Fields["CalcSchoolYearEntryScores"] = entryCalcScores;
                else
                    var.Fields.Add("CalcSchoolYearEntryScores", entryCalcScores);
            }
            return _ErrorList;
        }

        /// <summary>
        /// 計算學年分項成績
        /// </summary>
        /// <param name="schoolyear">學年度</param>
        /// <param name="accesshelper"></param>
        /// <param name="students">學生物件列表</param>
        /// <returns></returns>
        public Dictionary<StudentRecord, List<string>> FillSchoolYearEntryCalcScore(int schoolyear, AccessHelper accesshelper, List<StudentRecord> students)
        {
            Dictionary<StudentRecord, List<string>> _ErrorList = new Dictionary<StudentRecord, List<string>>();
            //取得學期分項成績
            accesshelper.StudentHelper.FillSemesterEntryScore(false, students);
            foreach (StudentRecord var in students)
            {
                //計算結果
                Dictionary<string, decimal> entryCalcScores = new Dictionary<string, decimal>();
                //精準位數
                int decimals = 2;
                //進位模式
                RoundMode mode = RoundMode.四捨五入;
                //成績年級及計算規則皆存在，允許計算成績
                bool canCalc = true;
                Dictionary<string, bool> calcEntry = new Dictionary<string, bool>();
                #region 取得成績年級跟計算規則
                {
                    #region 處理計算規則
                    XmlElement scoreCalcRule = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID).ScoreCalcRuleElement;
                    if (scoreCalcRule == null)
                    {
                        if (!_ErrorList.ContainsKey(var))
                            _ErrorList.Add(var, new List<string>());
                        _ErrorList[var].Add("沒有設定成績計算規則。");
                        canCalc &= false;
                    }
                    else
                    {
                        DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                        bool tryParsebool;
                        int tryParseint;
                        decimal tryParseDecimal;

                        #region 精準位數
                        if (scoreCalcRule.SelectSingleNode("各項成績計算位數/學年分項成績計算位數") != null)
                        {
                            if (int.TryParse(helper.GetText("各項成績計算位數/學年分項成績計算位數/@位數"), out tryParseint))
                                decimals = tryParseint;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/學年分項成績計算位數/@四捨五入"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.四捨五入;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/學年分項成績計算位數/@無條件捨去"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.無條件捨去;
                            if (bool.TryParse(helper.GetText("各項成績計算位數/學年分項成績計算位數/@無條件進位"), out tryParsebool) && tryParsebool)
                                mode = RoundMode.無條件進位;
                        }
                        #endregion
                        #region 計算類別
                        foreach (string entry in new string[] { "體育", "學業", "國防通識", "健康與護理", "實習科目", "專業科目" })
                        {
                            if (scoreCalcRule.SelectSingleNode("分項成績計算項目") == null || scoreCalcRule.SelectSingleNode("分項成績計算項目/" + entry) == null || ((XmlElement)scoreCalcRule.SelectSingleNode("分項成績計算項目/" + entry)).GetAttribute("計算成績") == "True")
                                calcEntry.Add(entry, true);
                            else
                                calcEntry.Add(entry, false);
                        }
                        #endregion
                    }
                    #endregion
                }
                #endregion
                if (canCalc)
                {
                    int? gradeyear = null;
                    #region 抓年級
                    foreach (SemesterEntryScoreInfo score in var.SemesterEntryScoreList)
                    {
                        if (calcEntry.ContainsKey(score.Entry) && score.SchoolYear == schoolyear)
                        {
                            if (gradeyear == null || score.GradeYear > gradeyear)
                                gradeyear = score.GradeYear;
                        }
                    }
                    #endregion
                    if (gradeyear != null)
                    {
                        #region 移除不需要成績
                        Dictionary<int, int> ApplySemesterSchoolYear = new Dictionary<int, int>();
                        //先掃一遍抓出該年級最高的學年度
                        foreach (SemesterEntryScoreInfo scoreInfo in var.SemesterEntryScoreList)
                        {
                            if (scoreInfo.SchoolYear <= schoolyear && scoreInfo.GradeYear == gradeyear)
                            {
                                if (!ApplySemesterSchoolYear.ContainsKey(scoreInfo.Semester))
                                    ApplySemesterSchoolYear.Add(scoreInfo.Semester, scoreInfo.SchoolYear);
                                else
                                {
                                    if (ApplySemesterSchoolYear[scoreInfo.Semester] < scoreInfo.SchoolYear)
                                        ApplySemesterSchoolYear[scoreInfo.Semester] = scoreInfo.SchoolYear;
                                }
                            }
                        }
                        //如果成績資料的年級學年度不在清單中就移掉
                        List<SemesterEntryScoreInfo> removeList = new List<SemesterEntryScoreInfo>();
                        foreach (SemesterEntryScoreInfo scoreInfo in var.SemesterEntryScoreList)
                        {
                            if (!ApplySemesterSchoolYear.ContainsKey(scoreInfo.Semester) || ApplySemesterSchoolYear[scoreInfo.Semester] != scoreInfo.SchoolYear)
                                removeList.Add(scoreInfo);
                        }
                        foreach (SemesterEntryScoreInfo scoreInfo in removeList)
                        {
                            var.SemesterEntryScoreList.Remove(scoreInfo);
                        }
                        #endregion
                        #region 計算該年級的分項成績
                        Dictionary<string, List<decimal>> entryScores = new Dictionary<string, List<decimal>>();
                        foreach (SemesterEntryScoreInfo score in var.SemesterEntryScoreList)
                        {
                            if (calcEntry.ContainsKey(score.Entry) && score.SchoolYear <= schoolyear && score.GradeYear == gradeyear)
                            {
                                if (!entryScores.ContainsKey(score.Entry))
                                    entryScores.Add(score.Entry, new List<decimal>());
                                entryScores[score.Entry].Add(score.Score);
                            }
                        }
                        foreach (string key in entryScores.Keys)
                        {
                            decimal sum = 0;
                            decimal count = 0;
                            foreach (decimal sc in entryScores[key])
                            {
                                sum += sc;
                                count += 1;
                            }
                            if (count > 0)
                                entryCalcScores.Add(key, GetRoundScore(sum / count, decimals, mode));
                        }
                        #endregion
                    }
                }
                if (var.Fields.ContainsKey("CalcSchoolYearEntryScores"))
                    var.Fields["CalcSchoolYearEntryScores"] = entryCalcScores;
                else
                    var.Fields.Add("CalcSchoolYearEntryScores", entryCalcScores);
            }
            return _ErrorList;
        }

        /// <summary>
        /// 計算學年科目成績
        /// </summary>
        /// <param name="schoolyear">要計算的學年度</param>
        /// <param name="accesshelper"></param>
        /// <param name="students">學生物件列表</param>
        /// <returns></returns>
        public Dictionary<StudentRecord, List<string>> FillSchoolYearSubjectCalcScore(int schoolyear, AccessHelper accesshelper, List<StudentRecord> students)
        {
            Dictionary<StudentRecord, List<string>> _ErrorList = new Dictionary<StudentRecord, List<string>>();
            //取得學生學期科目成績資料
            accesshelper.StudentHelper.FillSemesterSubjectScore(false, students);

            //針對每位學生計算學期科目成績
            foreach (StudentRecord var in students)
            {
                #region 處理CalcSchoolYearSubjectScores
                //計算結果
                Dictionary<string, decimal> subjectCalcScores = new Dictionary<string, decimal>();
                //精準位數
                int decimals = 2;
                //進位模式
                RoundMode mode = RoundMode.四捨五入;
                //成績年級及計算規則皆存在，允許計算成績
                bool canCalc = true;
                Dictionary<string, bool> calcEntry = new Dictionary<string, bool>();


                List<string> scoreTypeList = new List<string>();
                bool takeRepairScore = false;

                #region 取得成績年級跟計算規則
                #region 處理計算規則

                //取得某位學生的成績計算規則
                XmlElement scoreCalcRule = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID).ScoreCalcRuleElement;
                if (scoreCalcRule == null)
                {
                    if (!_ErrorList.ContainsKey(var))
                        _ErrorList.Add(var, new List<string>());
                    _ErrorList[var].Add("沒有設定成績計算規則。");
                    canCalc &= false;
                }
                else
                {
                    DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                    bool tryParsebool;
                    int tryParseint;
                    decimal tryParseDecimal;

                    #region 精準位數
                    if (scoreCalcRule.SelectSingleNode("各項成績計算位數/學年科目成績計算位數") != null)
                    {
                        if (int.TryParse(helper.GetText("各項成績計算位數/學年科目成績計算位數/@位數"), out tryParseint))
                            decimals = tryParseint;
                        if (bool.TryParse(helper.GetText("各項成績計算位數/學年科目成績計算位數/@四捨五入"), out tryParsebool) && tryParsebool)
                            mode = RoundMode.四捨五入;
                        if (bool.TryParse(helper.GetText("各項成績計算位數/學年科目成績計算位數/@無條件捨去"), out tryParsebool) && tryParsebool)
                            mode = RoundMode.無條件捨去;
                        if (bool.TryParse(helper.GetText("各項成績計算位數/學年科目成績計算位數/@無條件進位"), out tryParsebool) && tryParsebool)
                            mode = RoundMode.無條件進位;
                    }
                    #endregion

                    #region 學年成績計算採計成績欄位
                    if (scoreCalcRule.SelectSingleNode("學年成績計算採計成績欄位") != null)
                    {
                        // 處理補修成績
                        bool.TryParse(helper.GetText("學年成績計算採計成績欄位/@補修成績"), out takeRepairScore);

                        if (bool.TryParse(helper.GetText("學年成績計算採計成績欄位/@原始成績"), out tryParsebool) && tryParsebool)
                        {
                            scoreTypeList.Add("原始成績");
                        }
                        if (bool.TryParse(helper.GetText("學年成績計算採計成績欄位/@補考成績"), out tryParsebool) && tryParsebool)
                        {
                            scoreTypeList.Add("補考成績");
                        }
                        if (bool.TryParse(helper.GetText("學年成績計算採計成績欄位/@重修成績"), out tryParsebool) && tryParsebool)
                        {
                            scoreTypeList.Add("重修成績");
                        }
                        if (bool.TryParse(helper.GetText("學年成績計算採計成績欄位/@擇優採計成績"), out tryParsebool) && tryParsebool)
                        {
                            scoreTypeList.Add("擇優採計成績");
                        }
                        if (bool.TryParse(helper.GetText("學年成績計算採計成績欄位/@學年調整成績"), out tryParsebool) && tryParsebool)
                        {
                            scoreTypeList.Add("學年調整成績");
                        }


                    }
                    // 2017/1/13 穎驊 與恩正討論 過後， 需要 再加上一個 若使用者都沒有 儲存設定 的預設， 此預設項目 與舊模式 相同，
                    // 也與  SmartSchool.Evaluation.Configuration.ScoreCalcRuleEditor 內的 使用者設定介面預設  一樣
                    else
                    {
                        scoreTypeList.Add("原始成績");
                        scoreTypeList.Add("補考成績");
                        scoreTypeList.Add("重修成績");
                        scoreTypeList.Add("擇優採計成績");
                        //scoreTypeList.Add("學年調整成績");                                        
                    }

                    #endregion


                }
                #endregion
                #endregion
                int? gradeyear = null;
                if (canCalc)
                {
                    #region 根據學期科目成績中的學年度取得成績計算的年級，年級以高的為準
                    foreach (SemesterSubjectScoreInfo score in var.SemesterSubjectScoreList)
                    {
                        //判斷學期科目成績的學年度
                        if (score.SchoolYear == schoolyear)
                        {
                            //假設年級為null或是目前學期科目成績的年級大於目前成績
                            if (gradeyear == null || score.GradeYear > gradeyear)
                                gradeyear = score.GradeYear;
                        }
                    }
                    #endregion

                    if (gradeyear != null)
                    {
                        #region 移除不需要成績

                        //此為實際計算成績的學年度及學期。Key是學期，Value是學年度；其中的Key應該只有1及2
                        Dictionary<int, int> ApplySemesterSchoolYear = new Dictionary<int, int>();

                        //先掃一遍抓出該年級最高的學年度
                        foreach (SemesterSubjectScoreInfo scoreInfo in var.SemesterSubjectScoreList)
                        {
                            if (scoreInfo.SchoolYear <= schoolyear && scoreInfo.GradeYear == gradeyear)
                            {
                                if (!ApplySemesterSchoolYear.ContainsKey(scoreInfo.Semester))
                                    ApplySemesterSchoolYear.Add(scoreInfo.Semester, scoreInfo.SchoolYear);
                                else
                                {
                                    if (ApplySemesterSchoolYear[scoreInfo.Semester] < scoreInfo.SchoolYear)
                                        ApplySemesterSchoolYear[scoreInfo.Semester] = scoreInfo.SchoolYear;
                                }
                            }
                        }
                        //如果成績資料的年級學年度不在清單中（指定的學年度及學期）就移掉
                        // 如果課程代碼17碼為「9」，19碼為「D」的成績先寫死這個規則並移除 -2023.5.9 俊緯
                        List<SemesterSubjectScoreInfo> removeList = new List<SemesterSubjectScoreInfo>();
                        foreach (SemesterSubjectScoreInfo scoreInfo in var.SemesterSubjectScoreList)
                        {
                            if (!ApplySemesterSchoolYear.ContainsKey(scoreInfo.Semester) || ApplySemesterSchoolYear[scoreInfo.Semester] != scoreInfo.SchoolYear)
                                removeList.Add(scoreInfo);

                            // 2024/3/4 討論結果，不需評分 ="是"，不計算
                            if (scoreInfo.Detail.GetAttribute("不需評分") == "是")
                            {
                                removeList.Add(scoreInfo);
                            }

                            //string subjectCode = scoreInfo.Detail.GetAttribute("修課科目代碼");
                            //if (subjectCode.Length == 23)
                            //{
                            //    if (subjectCode[16].ToString() + subjectCode[18].ToString() == "9D" || subjectCode[16].ToString() + subjectCode[18].ToString() == "9d")
                            //    {
                            //        removeList.Add(scoreInfo);
                            //    }
                            //}

                            // 若為補修成績，且成績計算不採計，則移除
                            if (takeRepairScore == false && scoreInfo.Detail.GetAttribute("是否補修成績") == "是")
                            {
                                removeList.Add(scoreInfo);
                            }
                        }
                        foreach (SemesterSubjectScoreInfo scoreInfo in removeList)
                        {
                            var.SemesterSubjectScoreList.Remove(scoreInfo);
                        }
                        #endregion
                        #region 計算該年級的科目成績
                        Dictionary<string, Dictionary<SemesterSubjectScoreInfo, decimal>> subjectScores = new Dictionary<string, Dictionary<SemesterSubjectScoreInfo, decimal>>();

                        //針對每筆學期科目成績，將學期科目先放到subjectScores
                        foreach (SemesterSubjectScoreInfo score in var.SemesterSubjectScoreList)
                        {
                            if (score.SchoolYear <= schoolyear && score.GradeYear == gradeyear)
                            {
                                //先判斷這筆科目成績能不能計算
                                decimal maxscore = decimal.MinValue, tryParsedecimal;

                                bool hasScore = false;

                                // 2016/12/22 穎驊因應 客戶學校反映調整， 重修成績只能取得學分，不能納入計算分數， 恩正指示調整此處，移除 scoreType 重修成績。                                
                                //將學期科目成績擇優
                                //foreach (string scoreType in new string[] { "原始成績", "補考成績", "擇優採計成績" })
                                //{
                                //    if (decimal.TryParse(score.Detail.GetAttribute(scoreType), out tryParsedecimal))
                                //    {
                                //        hasScore |= true;
                                //        if (tryParsedecimal > maxscore)
                                //            maxscore = tryParsedecimal;
                                //    }
                                //}


                                // 2017/1/12 穎驊因應學校調整， 將上次的(上面)的更動 與恩正、韻如討論後， 決定 將"學年成績計算" 會有哪幾項成績納入計算
                                //一併放在 教務作業 設定 中的 "成績計算規則" 一起設定

                                foreach (string scoreType in scoreTypeList)
                                {
                                    if (decimal.TryParse(score.Detail.GetAttribute(scoreType), out tryParsedecimal))
                                    {
                                        hasScore |= true;
                                        if (tryParsedecimal > maxscore)
                                            maxscore = tryParsedecimal;
                                    }
                                }

                                //可以被計算
                                if (hasScore)
                                {
                                    string key = score.Subject;
                                    if (score.Detail.GetAttribute("指定學年科目名稱") != "")
                                        key = score.Detail.GetAttribute("指定學年科目名稱");

                                    if (!subjectScores.ContainsKey(key))
                                        subjectScores.Add(key, new Dictionary<SemesterSubjectScoreInfo, decimal>());
                                    subjectScores[key].Add(score, maxscore);
                                }
                            }
                        }

                        //計算年級科目成績
                        foreach (string key in subjectScores.Keys)
                        {
                            decimal sum = 0;
                            decimal count = 0;
                            foreach (decimal sc in subjectScores[key].Values)
                            {
                                sum += sc;
                                count += 1;
                            }
                            if (count > 0)
                            {
                                decimal schoolYearSubjectScore = GetRoundScore(sum / count, decimals, mode);

                                if (schoolYearSubjectScore < 60 && subjectScores[key].Count >= 2)
                                {
                                    XmlElement element = scoreCalcRule.SelectSingleNode("學期科目成績計算至學年科目成績規則") as XmlElement;
                                    if (element != null && element.GetAttribute("進校上下學期及格規則").Equals("True"))
                                    {
                                        bool upPass = false, downPass = false;
                                        foreach (var sssi in subjectScores[key].Keys)
                                        {
                                            if (sssi.Semester == 1 && subjectScores[key][sssi] >= 50) upPass = true;
                                            if (sssi.Semester == 2 && subjectScores[key][sssi] >= 60) downPass = true;
                                        }
                                        if (upPass && downPass) { schoolYearSubjectScore = 60; }
                                    }
                                }
                                subjectCalcScores.Add(key, schoolYearSubjectScore);
                            }
                        }
                        #endregion
                    }
                }
                if (var.Fields.ContainsKey("CalcSchoolYearSubjectScores"))
                    var.Fields["CalcSchoolYearSubjectScores"] = subjectCalcScores;
                else
                    var.Fields.Add("CalcSchoolYearSubjectScores", subjectCalcScores);
                #endregion
                //#region 處理SchoolYearApplyScores
                //canCalc = true;
                ////及格標準<年及,及格標準>
                //Dictionary<int, decimal> applyLimit = new Dictionary<int, decimal>();
                ////0:不登錄，1:60分登錄，2:及格標準登錄
                //int regWay = 0;
                //#region 處理計算規則
                //if (scoreCalcRule == null)
                //{
                //    canCalc &= false;
                //}
                //else
                //{
                //    DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                //    #region 及格標準
                //    foreach (XmlElement element in helper.GetElements("及格標準/學生類別"))
                //    {
                //        string cat = element.GetAttribute("類別");
                //        bool useful = false;
                //        //掃描學生的類別作比對
                //        foreach (CategoryInfo catinfo in var.StudentCategorys)
                //        {
                //            if (catinfo.Name == cat || catinfo.FullName == cat)
                //                useful = true;
                //        }
                //        //學生是指定的類別或類別為"預設"
                //        if (cat == "預設" || useful)
                //        {
                //            decimal tryParseDecimal;
                //            for (int gyear = 1; gyear <= 4; gyear++)
                //            {
                //                switch (gyear)
                //                {
                //                    case 1:
                //                        if (decimal.TryParse(element.GetAttribute("一年級及格標準"), out tryParseDecimal))
                //                        {
                //                            if (!applyLimit.ContainsKey(gyear))
                //                                applyLimit.Add(gyear, tryParseDecimal);
                //                            if (applyLimit[gyear] > tryParseDecimal)
                //                                applyLimit[gyear] = tryParseDecimal;
                //                        }
                //                        break;
                //                    case 2:
                //                        if (decimal.TryParse(element.GetAttribute("二年級及格標準"), out tryParseDecimal))
                //                        {
                //                            if (!applyLimit.ContainsKey(gyear))
                //                                applyLimit.Add(gyear, tryParseDecimal);
                //                            if (applyLimit[gyear] > tryParseDecimal)
                //                                applyLimit[gyear] = tryParseDecimal;
                //                        }
                //                        break;
                //                    case 3:
                //                        if (decimal.TryParse(element.GetAttribute("三年級及格標準"), out tryParseDecimal))
                //                        {
                //                            if (!applyLimit.ContainsKey(gyear))
                //                                applyLimit.Add(gyear, tryParseDecimal);
                //                            if (applyLimit[gyear] > tryParseDecimal)
                //                                applyLimit[gyear] = tryParseDecimal;
                //                        }
                //                        break;
                //                    case 4:
                //                        if (decimal.TryParse(element.GetAttribute("四年級及格標準"), out tryParseDecimal))
                //                        {
                //                            if (!applyLimit.ContainsKey(gyear))
                //                                applyLimit.Add(gyear, tryParseDecimal);
                //                            if (applyLimit[gyear] > tryParseDecimal)
                //                                applyLimit[gyear] = tryParseDecimal;
                //                        }
                //                        break;
                //                    default:
                //                        break;
                //                }
                //            }
                //        }
                //    }
                //    #endregion
                //    bool tryParseBool = false;
                //    if (bool.TryParse(helper.GetText("學年調整成績/@不登錄學年調整成績"), out tryParseBool) && tryParseBool)
                //        regWay = 0;
                //    if (bool.TryParse(helper.GetText("學年調整成績/@以六十分登錄"), out tryParseBool) && tryParseBool)
                //        regWay = 1;
                //    if (bool.TryParse(helper.GetText("學年調整成績/@以學生及格標準登錄"), out tryParseBool) && tryParseBool)
                //        regWay = 2;
                //    if (bool.TryParse(helper.GetText("學年調整成績/@不使用學年調整成績"), out tryParseBool) && tryParseBool)
                //        regWay = 3;
                //}
                //#endregion
                //#region 處理學年調整成績
                //List<SemesterSubjectScoreInfo> updateScores = new List<SemesterSubjectScoreInfo>();
                //if (gradeyear != null)
                //{
                //    decimal applylimit = 40;
                //    if (applyLimit.ContainsKey((int)gradeyear))
                //        applylimit = applyLimit[(int)gradeyear];
                //    foreach (SemesterSubjectScoreInfo score in var.SemesterSubjectScoreList)
                //    {
                //        if (!score.Pass && subjectCalcScores.ContainsKey(score.Subject) && subjectCalcScores[score.Subject] >= applylimit)
                //        {
                //            if (regWay != 3)//不使用學年調整成績
                //            {
                //                switch (regWay)
                //                {
                //                    default:
                //                    case 0:
                //                        break;
                //                    case 1:
                //                        score.Detail.SetAttribute("學年調整成績", "60");
                //                        break;
                //                    case 2:
                //                        score.Detail.SetAttribute("學年調整成績", "" + applylimit);
                //                        break;

                //                }
                //                score.Detail.SetAttribute("是否取得學分", "是");
                //                updateScores.Add(score);
                //            }
                //        }
                //    }
                //}
                //if (var.Fields.ContainsKey("SchoolYearApplyScores"))
                //    var.Fields["SchoolYearApplyScores"] = updateScores;
                //else
                //    var.Fields.Add("SchoolYearApplyScores", updateScores);
                //#endregion
                //#endregion
            }
            return _ErrorList;
        }


        /// <summary>
        /// 計算學年調整成績
        /// </summary>
        /// <param name="schoolyear">要計算的學年度</param>
        /// <param name="accesshelper"></param>
        /// <param name="students">學生物件列表</param>
        /// <returns></returns>
        public Dictionary<StudentRecord, List<string>> FillSchoolYearApplyScoresCalcScore(int schoolyear, AccessHelper accesshelper, List<StudentRecord> students)
        {
            Dictionary<StudentRecord, List<string>> _ErrorList = new Dictionary<StudentRecord, List<string>>();
            //取得學生學期科目成績資料
            accesshelper.StudentHelper.FillSemesterSubjectScore(false, students);
            accesshelper.StudentHelper.FillSchoolYearSubjectScore(false, students);

            foreach (StudentRecord var in students)
            {
                bool canCalc = true;
                //成績年級
                int? gradeyear = null;
                //及格標準<年及,及格標準>
                Dictionary<int, decimal> applyLimit = new Dictionary<int, decimal>();
                //0:不登錄，1:60分登錄，2:及格標準登錄，3.不使用學年調整成績
                int regWay = 3;  // 2014/3/26 ChenCT，因使用者需求調整預設不使用學年調整成績
                #region 取得成績年級跟計算規則
                #region 處理計算規則
                //取得某位學生的成績計算規則
                XmlElement scoreCalcRule = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID).ScoreCalcRuleElement;
                if (scoreCalcRule == null)
                {
                    canCalc &= false;
                }
                else
                {
                    DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                    #region 及格標準
                    foreach (XmlElement element in helper.GetElements("及格標準/學生類別"))
                    {
                        string cat = element.GetAttribute("類別");
                        bool useful = false;
                        //掃描學生的類別作比對
                        foreach (CategoryInfo catinfo in var.StudentCategorys)
                        {
                            if (catinfo.Name == cat || catinfo.FullName == cat)
                                useful = true;
                        }
                        //學生是指定的類別或類別為"預設"
                        if (cat == "預設" || useful)
                        {
                            decimal tryParseDecimal;
                            for (int gyear = 1; gyear <= 4; gyear++)
                            {
                                switch (gyear)
                                {
                                    case 1:
                                        if (decimal.TryParse(element.GetAttribute("一年級及格標準"), out tryParseDecimal))
                                        {
                                            if (!applyLimit.ContainsKey(gyear))
                                                applyLimit.Add(gyear, tryParseDecimal);
                                            if (applyLimit[gyear] > tryParseDecimal)
                                                applyLimit[gyear] = tryParseDecimal;
                                        }
                                        break;
                                    case 2:
                                        if (decimal.TryParse(element.GetAttribute("二年級及格標準"), out tryParseDecimal))
                                        {
                                            if (!applyLimit.ContainsKey(gyear))
                                                applyLimit.Add(gyear, tryParseDecimal);
                                            if (applyLimit[gyear] > tryParseDecimal)
                                                applyLimit[gyear] = tryParseDecimal;
                                        }
                                        break;
                                    case 3:
                                        if (decimal.TryParse(element.GetAttribute("三年級及格標準"), out tryParseDecimal))
                                        {
                                            if (!applyLimit.ContainsKey(gyear))
                                                applyLimit.Add(gyear, tryParseDecimal);
                                            if (applyLimit[gyear] > tryParseDecimal)
                                                applyLimit[gyear] = tryParseDecimal;
                                        }
                                        break;
                                    case 4:
                                        if (decimal.TryParse(element.GetAttribute("四年級及格標準"), out tryParseDecimal))
                                        {
                                            if (!applyLimit.ContainsKey(gyear))
                                                applyLimit.Add(gyear, tryParseDecimal);
                                            if (applyLimit[gyear] > tryParseDecimal)
                                                applyLimit[gyear] = tryParseDecimal;
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                    }
                    #endregion
                    bool tryParseBool = false;
                    if (bool.TryParse(helper.GetText("學年調整成績/@不登錄學年調整成績"), out tryParseBool) && tryParseBool)
                        regWay = 0;
                    if (bool.TryParse(helper.GetText("學年調整成績/@以六十分登錄"), out tryParseBool) && tryParseBool)
                        regWay = 1;
                    if (bool.TryParse(helper.GetText("學年調整成績/@以學生及格標準登錄"), out tryParseBool) && tryParseBool)
                        regWay = 2;
                    if (bool.TryParse(helper.GetText("學年調整成績/@不使用學年調整成績"), out tryParseBool) && tryParseBool)
                        regWay = 3;
                }
                #endregion
                #endregion
                #region 根據學期科目成績中的學年度取得成績計算的年級，年級以高的為準
                foreach (SemesterSubjectScoreInfo score in var.SemesterSubjectScoreList)
                {
                    //判斷學期科目成績的學年度
                    if (score.SchoolYear == schoolyear)
                    {
                        //假設年級為null或是目前學期科目成績的年級大於目前成績
                        if (gradeyear == null || score.GradeYear > gradeyear)
                            gradeyear = score.GradeYear;
                    }
                }
                #endregion
                #region 處理學年調整成績
                List<SemesterSubjectScoreInfo> updateScores = new List<SemesterSubjectScoreInfo>();
                if (gradeyear != null)
                {
                    decimal applylimit = 40;
                    if (applyLimit.ContainsKey((int)gradeyear))
                        applylimit = applyLimit[(int)gradeyear];
                    foreach (SemesterSubjectScoreInfo score in var.SemesterSubjectScoreList)
                    {
                        // CT,2024/10/15，當學期科目成績 是否補修 = "是"，略過不處理
                        // 因工單：系統功能修改 事由[多校]計算學年調整成績時未排除補修成績判斷
                        //https://3.basecamp.com/4399967/buckets/15765350/todos/7916166121

                        if (score.Detail.GetAttribute("是否補修成績") == "是")
                            continue;

                        if (!score.Pass && score.SchoolYear == schoolyear)
                        {//&& subjectCalcScores.ContainsKey(score.Subject) && subjectCalcScores[score.Subject] >= applylimit
                            foreach (var schoolYearSubjectScore in var.SchoolYearSubjectScoreList)
                            {

                                string semesterSubject = score.Subject;
                                if (score.Detail.GetAttribute("指定學年科目名稱") != "")
                                    semesterSubject = score.Detail.GetAttribute("指定學年科目名稱");

                                if (schoolYearSubjectScore.SchoolYear == schoolyear && schoolYearSubjectScore.Subject == semesterSubject
                                    && schoolYearSubjectScore.Score >= applylimit)
                                {
                                    //subjectCalcScores.Add(schoolYearSubjectScore.Subject, schoolYearSubjectScore.Score);//[score.Subject]
                                    decimal tryParseDecimal, topScore = decimal.MinValue;
                                    string tip = score.Detail.GetAttribute("註記");
                                    foreach (string field in new string[] { "結算成績", "補考成績", "重修成績" })
                                    {
                                        tip = tip.Replace("學年" + field + "及格。", "");
                                        if (decimal.TryParse(schoolYearSubjectScore.Detail.GetAttribute(field), out tryParseDecimal) && tryParseDecimal >= applylimit && tryParseDecimal > topScore)
                                        {
                                            tip += "學年" + field + "及格。";
                                            topScore = tryParseDecimal;
                                        }
                                    }
                                    if (regWay != 3)//不使用學年調整成績
                                    {
                                        switch (regWay)
                                        {
                                            default:
                                            case 0:
                                                break;
                                            case 1:
                                                score.Detail.SetAttribute("學年調整成績", "60");
                                                break;
                                            case 2:
                                                // 讀取學期科目成績修課及格標準，當有設定使用修課
                                                decimal d;
                                                if (decimal.TryParse(score.Detail.GetAttribute("修課及格標準"), out d))
                                                {
                                                    applylimit = d;
                                                }
                                                score.Detail.SetAttribute("學年調整成績", "" + applylimit);
                                                break;
                                        }
                                        score.Detail.SetAttribute("是否取得學分", "是");
                                        score.Detail.SetAttribute("註記", tip);
                                        updateScores.Add(score);
                                    }
                                }
                            }
                        }
                    }
                }
                if (var.Fields.ContainsKey("SchoolYearApplyScores"))
                    var.Fields["SchoolYearApplyScores"] = updateScores;
                else
                    var.Fields.Add("SchoolYearApplyScores", updateScores);
                #endregion
            }
            return _ErrorList;
        }

        public enum GradScoreCalcMode
        {
            SubjectWeighted,      // 學期科目成績加權
            SemesterEntryAverage, // 學期分項成績平均
            SchoolYearEntryAverage // 學年分項成績平均
        }

        /// <summary>
        /// 計算學生畢業成績
        /// </summary>
        /// <param name="accesshelper"></param>
        /// <param name="students"></param>
        /// <returns></returns>
        public Dictionary<StudentRecord, List<string>> FillStudentGradCalcScore(AccessHelper accesshelper, List<StudentRecord> students)
        {
            Dictionary<StudentRecord, List<string>> _ErrorList = new Dictionary<StudentRecord, List<string>>();
            //抓成績資料
            accesshelper.StudentHelper.FillSemesterSubjectScore(true, students);
            accesshelper.StudentHelper.FillSemesterEntryScore(true, students);
            accesshelper.StudentHelper.FillSchoolYearEntryScore(true, students); // 新增學年分項成績提取

            foreach (StudentRecord var in students)
            {
                //可以計算(表示需要的資料都有)
                bool canCalc = true;
                //精準位數(預設小數點後2位)
                int decimals = 2;
                //進位模式(預設四捨五入)
                RoundMode mode = RoundMode.四捨五入;
                //計算資料採用課程規劃(預設採用課程規劃)
                bool useGPlan = true;
                ////畢業成績使用所有科目成績加權計算(預設為否=>使用分項成績平均 )
                //bool useSubjectAdv = false;
                GradScoreCalcMode calcMode = GradScoreCalcMode.SubjectWeighted; // 預設為科目加權

                XmlElement scoreCalcRule = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID).ScoreCalcRuleElement;
                if (scoreCalcRule == null)
                {
                    if (!_ErrorList.ContainsKey(var))
                        _ErrorList.Add(var, new List<string>());
                    _ErrorList[var].Add("沒有設定成績計算規則。");
                    canCalc &= false;
                }
                else
                {
                    DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                    #region 處理精準位數
                    bool tryParsebool;
                    int tryParseint;
                    decimal tryParseDecimal;

                    if (scoreCalcRule.SelectSingleNode("各項成績計算位數/畢業成績計算位數") != null)
                    {
                        if (int.TryParse(helper.GetText("各項成績計算位數/畢業成績計算位數/@位數"), out tryParseint))
                            decimals = tryParseint;
                        if (bool.TryParse(helper.GetText("各項成績計算位數/畢業成績計算位數/@四捨五入"), out tryParsebool) && tryParsebool)
                            mode = RoundMode.四捨五入;
                        if (bool.TryParse(helper.GetText("各項成績計算位數/畢業成績計算位數/@無條件捨去"), out tryParsebool) && tryParsebool)
                            mode = RoundMode.無條件捨去;
                        if (bool.TryParse(helper.GetText("各項成績計算位數/畢業成績計算位數/@無條件進位"), out tryParsebool) && tryParsebool)
                            mode = RoundMode.無條件進位;
                    }
                    #endregion
                    #region 處理採計方式
                    if (scoreCalcRule.SelectSingleNode("學期科目成績屬性採計方式") != null)
                    {
                        if (scoreCalcRule.SelectSingleNode("學期科目成績屬性採計方式").InnerText == "以實際學期科目成績內容為準")
                        {
                            useGPlan = false;
                        }
                    }
                    if (useGPlan && GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(var.StudentID) == null)
                    {
                        if (!_ErrorList.ContainsKey(var))
                            _ErrorList.Add(var, new List<string>());
                        _ErrorList[var].Add("沒有設定課程規劃表，當\"學期科目成績屬性採計方式\"設定使用\"以課程規劃表內容為準\"時，學生必須要有課程規劃表以做參考。");
                        canCalc &= false;
                    }
                    #endregion
                    // CT 修正 XML 文字錯誤導致設定無法比對
                    #region 處理畢業成績計算規則
                    if (scoreCalcRule.SelectSingleNode("畢業成績計算規則") != null)
                    {
                        //useSubjectAdv = scoreCalcRule.SelectSingleNode("畢業成績計算規則").InnerText == "學期科目成績加權";

                        string modeText = scoreCalcRule.SelectSingleNode("畢業成績計算規則").InnerText;
                        switch (modeText)
                        {
                            case "學期科目成績加權":
                                calcMode = GradScoreCalcMode.SubjectWeighted;
                                break;
                            case "學期分項成績平均":
                                calcMode = GradScoreCalcMode.SemesterEntryAverage;
                                break;
                            case "學年分項成績平均":
                                calcMode = GradScoreCalcMode.SchoolYearEntryAverage;
                                break;
                            default:
                                if (!_ErrorList.ContainsKey(var))
                                    _ErrorList.Add(var, new List<string>());
                                _ErrorList[var].Add("畢業成績計算模式設定無效。");
                                canCalc &= false;
                                break;
                        }
                    }
                    #endregion
                }


                XmlDocument doc = new XmlDocument();
                XmlElement gradeCalcScoreElement = doc.CreateElement("GradScore");
                if (canCalc)
                {
                    Dictionary<string, decimal> entryCount = new Dictionary<string, decimal>();
                    Dictionary<string, decimal> entrySum = new Dictionary<string, decimal>();
                    //使用所有科目成績加權計算畢業成績(學業)
                    if (calcMode == GradScoreCalcMode.SubjectWeighted)
                    {
                        #region 使用所有科目成績加權計算畢業成績(總分及總數)
                        decimal creditCount = 0;
                        decimal scoreSum = 0;

                        // 專業
                        decimal PcreditCount = 0;
                        decimal PscoreSum = 0;

                        // 實習
                        decimal IcreditCount = 0;
                        decimal IscoreSum = 0;


                        foreach (SemesterSubjectScoreInfo subjectScoreInfo in var.SemesterSubjectScoreList)
                        {
                            if (useGPlan)//相關屬性採用課程規劃
                            {
                                //不計學分或不需評分不用算
                                if (GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(var.StudentID).GetSubjectInfo(subjectScoreInfo.Subject, subjectScoreInfo.Level).NotIncludedInCredit ||
                                     GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(var.StudentID).GetSubjectInfo(subjectScoreInfo.Subject, subjectScoreInfo.Level).NotIncludedInCalc
                                    )
                                    continue;
                                decimal realCredit = 0;
                                decimal.TryParse(GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(var.StudentID).GetSubjectInfo(subjectScoreInfo.Subject, subjectScoreInfo.Level).Credit, out realCredit);


                                // 2024/3/4 討論結果，不需評分 ="是"，不計算
                                if (subjectScoreInfo.Detail.GetAttribute("不需評分") == "是")
                                    continue;

                                scoreSum += subjectScoreInfo.Score * realCredit;
                                creditCount += realCredit;
                            }
                            else//相關屬性直接採用成績內容
                            {
                                //不計學分不用算
                                if (subjectScoreInfo.Detail.GetAttribute("不計學分") == "是" || subjectScoreInfo.Detail.GetAttribute("不需評分") == "是")
                                    continue;
                                scoreSum += subjectScoreInfo.Score * subjectScoreInfo.Credit;
                                creditCount += subjectScoreInfo.Credit;
                            }

                            if (subjectScoreInfo.Detail.GetAttribute("開課分項類別") == "專業科目")
                            {
                                PscoreSum += subjectScoreInfo.Score * subjectScoreInfo.Credit;
                                PcreditCount += subjectScoreInfo.Credit;
                            }

                            if (subjectScoreInfo.Detail.GetAttribute("開課分項類別") == "實習科目")
                            {
                                IscoreSum += subjectScoreInfo.Score * subjectScoreInfo.Credit;
                                IcreditCount += subjectScoreInfo.Credit;
                            }

                        }
                        if (creditCount != 0)
                        {
                            entryCount.Add("學業", creditCount);
                            entrySum.Add("學業", scoreSum);
                        }

                        if (PcreditCount != 0)
                        {
                            entryCount.Add("專業科目", PcreditCount);
                            entrySum.Add("專業科目", PscoreSum);
                        }

                        if (IcreditCount != 0)
                        {
                            entryCount.Add("實習科目", IcreditCount);
                            entrySum.Add("實習科目", IscoreSum);
                        }
                        #endregion
                    }
                    else if (calcMode == GradScoreCalcMode.SemesterEntryAverage)
                    {
                        //計算各分項畢業成績(總分及總數)
                        foreach (SemesterEntryScoreInfo entryScore in var.SemesterEntryScoreList)
                        {
                            #region 計算各分項畢業成績
                            //if (entryScore.Entry != "學業" || !useSubjectAdv)//其它分項或者是學業分項但不使用科目成績加權

                            if (!entryCount.ContainsKey(entryScore.Entry))
                                entryCount.Add(entryScore.Entry, 0);
                            if (!entrySum.ContainsKey(entryScore.Entry))
                                entrySum.Add(entryScore.Entry, 0);

                            entryCount[entryScore.Entry]++;
                            entrySum[entryScore.Entry] += entryScore.Score;

                            #endregion
                        }
                    }
                    else if (calcMode == GradScoreCalcMode.SchoolYearEntryAverage)
                    {
                        #region 學年分項成績平均
                        foreach (SchoolYearEntryScoreInfo schoolYearScore in var.SchoolYearEntryScoreList)
                        {
                            if (!entryCount.ContainsKey(schoolYearScore.Entry))
                                entryCount.Add(schoolYearScore.Entry, 0);
                            if (!entrySum.ContainsKey(schoolYearScore.Entry))
                                entrySum.Add(schoolYearScore.Entry, 0);

                            entryCount[schoolYearScore.Entry]++;
                            entrySum[schoolYearScore.Entry] += schoolYearScore.Score;
                        }
                        #endregion
                    }


                    //計算&填入分項畢業成績
                    foreach (string entry in entryCount.Keys)
                    {
                        XmlElement entryScoreElement = doc.CreateElement("EntryScore");
                        entryScoreElement.SetAttribute("Entry", entry);
                        entryScoreElement.SetAttribute("Score", "" + GetRoundScore(entrySum[entry] / entryCount[entry], decimals, mode));
                        gradeCalcScoreElement.AppendChild(entryScoreElement);
                    }
                }
                if (var.Fields.ContainsKey("GradCalcScore"))
                    var.Fields.Remove("GradCalcScore");
                var.Fields.Add("GradCalcScore", gradeCalcScoreElement);
            }
            return _ErrorList;
        }

        /// <summary>
        /// 判斷學生是否符合畢業資格
        /// </summary>
        /// <param name="accesshelper"></param>
        /// <param name="students"></param>
        /// <returns></returns>
        public Dictionary<StudentRecord, List<string>> FillStudentGradCheck(AccessHelper accesshelper, List<StudentRecord> students)
        {
            return new SmartSchool.Evaluation.WearyDogComputerHelper.GraduateEvaluator(accesshelper, students).Evaluate();
        }

        public Dictionary<StudentRecord, List<string>> FillStudentFulfilledProgram(AccessHelper accesshelper, List<StudentRecord> students)
        {
            Dictionary<StudentRecord, List<string>> _ErrorList = new Dictionary<StudentRecord, List<string>>();
            //抓成績資料
            accesshelper.StudentHelper.FillSemesterSubjectScore(true, students);
            foreach (StudentRecord var in students)
            {
                bool canCalc = true;
                //計算資料採用課程規劃(預設採用課程規劃)
                bool useGPlan = true;

                XmlElement scoreCalcRule = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID).ScoreCalcRuleElement;
                if (scoreCalcRule == null)
                {
                    if (!_ErrorList.ContainsKey(var))
                        _ErrorList.Add(var, new List<string>());
                    _ErrorList[var].Add("沒有設定成績計算規則。");
                    canCalc &= false;
                }
                else
                {
                    DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                    #region 學期科目成績屬性採計方式
                    if (scoreCalcRule.SelectSingleNode("學期科目成績屬性採計方式") != null)
                    {
                        if (scoreCalcRule.SelectSingleNode("學期科目成績屬性採計方式").InnerText == "以實際學期科目成績內容為準")
                        {
                            useGPlan = false;
                        }
                    }
                    if (useGPlan && GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(var.StudentID) == null)
                    {
                        if (!_ErrorList.ContainsKey(var))
                            _ErrorList.Add(var, new List<string>());
                        _ErrorList[var].Add("沒有設定課程規劃表，當\"學期科目成績屬性採計方式\"設定使用\"以課程規劃表內容為準\"時，學生必須要有課程規劃表以做參考。");
                        canCalc &= false;
                    }
                    #endregion
                }

                XmlDocument doc = new XmlDocument();
                XmlElement fulfilledProgramElement = doc.CreateElement("FulfilledProgram");//可以計算(表示需要的資料都有)
                if (canCalc)
                {
                    foreach (SubjectTableItem subjectTable in SubjectTable.Items["學程科目表"])
                    {
                        XmlElement contentElement = (XmlElement)subjectTable.Content.SelectSingleNode("SubjectTableContent");

                        decimal passLimit;
                        decimal.TryParse(contentElement.GetAttribute("CreditCount"), out passLimit);

                        decimal coreLimit;
                        decimal.TryParse(contentElement.GetAttribute("CoreCount"), out coreLimit);

                        List<string> subjectLevelsInTable = new List<string>();
                        List<string> subjectLevelsInCore = new List<string>();
                        #region 整理在科目表中所有的科目級別(科目+^_^+級別)
                        foreach (XmlElement snode in contentElement.SelectNodes("Subject"))
                        {
                            string name = ((XmlElement)snode).GetAttribute("Name");
                            bool iscore = false;
                            bool.TryParse(snode.GetAttribute("IsCore"), out iscore);

                            if (snode.SelectNodes("Level").Count == 0)
                            {
                                subjectLevelsInTable.Add(name + "^_^");
                                if (iscore)
                                    subjectLevelsInCore.Add(name + "^_^");
                            }
                            else
                            {
                                foreach (XmlNode lnode in snode.SelectNodes("Level"))
                                {
                                    subjectLevelsInTable.Add(name + "^_^" + lnode.InnerText);
                                    if (iscore)
                                        subjectLevelsInCore.Add(name + "^_^" + lnode.InnerText);
                                }
                            }
                        }
                        #endregion

                        decimal credits = 0;
                        decimal coreCredits = 0;
                        foreach (SemesterSubjectScoreInfo subjectScore in var.SemesterSubjectScoreList)
                        {
                            //總學分數
                            if (subjectScore.Pass && subjectLevelsInTable.Contains(subjectScore.Subject + "^_^" + subjectScore.Level))
                            {
                                if (useGPlan)
                                {
                                    #region 從課程規劃表取得這個科目的相關屬性
                                    GraduationPlan.GraduationPlanSubject gPlanSubject = GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(var.StudentID).GetSubjectInfo(subjectScore.Subject, subjectScore.Level);
                                    //不計學分不用算
                                    if (gPlanSubject.NotIncludedInCredit)
                                        continue;

                                    decimal credit = 0;
                                    decimal.TryParse(gPlanSubject.Credit, out credit);
                                    credits += credit;
                                    #endregion
                                }
                                else
                                {
                                    #region 直接使用科目成績上的屬性
                                    //不計學分不用算
                                    if (subjectScore.Detail.GetAttribute("不計學分") == "是")
                                        continue;
                                    credits += subjectScore.Credit;
                                    #endregion
                                }
                            }
                            //核心學分數
                            if (subjectScore.Pass && subjectLevelsInCore.Contains(subjectScore.Subject + "^_^" + subjectScore.Level))
                            {
                                if (useGPlan)
                                {
                                    #region 從課程規劃表取得這個科目的相關屬性
                                    GraduationPlan.GraduationPlanSubject gPlanSubject = GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(var.StudentID).GetSubjectInfo(subjectScore.Subject, subjectScore.Level);
                                    //不計學分不用算
                                    if (gPlanSubject.NotIncludedInCredit)
                                        continue;

                                    decimal credit = 0;
                                    decimal.TryParse(gPlanSubject.Credit, out credit);
                                    coreCredits += credit;
                                    #endregion
                                }
                                else
                                {
                                    #region 直接使用科目成績上的屬性
                                    //不計學分不用算
                                    if (subjectScore.Detail.GetAttribute("不計學分") == "是")
                                        continue;
                                    coreCredits += subjectScore.Credit;
                                    #endregion
                                }
                            }
                        }
                        if (credits >= passLimit && coreCredits >= coreLimit)
                        {
                            XmlElement programElement = doc.CreateElement("Program");
                            programElement.InnerText = subjectTable.Name;
                            fulfilledProgramElement.AppendChild(programElement);
                        }
                    }
                }
                if (var.Fields.ContainsKey("FulfilledProgram"))
                    var.Fields.Remove("FulfilledProgram");
                var.Fields.Add("FulfilledProgram", fulfilledProgramElement);

            }
            return _ErrorList;
        }

        public void FillSemesterSubjectScoreInfoWithResit(AccessHelper accesshelper, bool filterRepeat, List<StudentRecord> students)
        {
            //抓科目成績
            accesshelper.StudentHelper.FillSemesterSubjectScore(filterRepeat, students);
            foreach (StudentRecord var in students)
            {
                //補考標準<年及,及格標準>
                Dictionary<int, decimal> resitLimit = new Dictionary<int, decimal>();
                //resitLimit.Add(1, 40);
                //resitLimit.Add(2, 40);
                //resitLimit.Add(3, 40);
                //resitLimit.Add(4, 40);

                #region 處理計算規則
                XmlElement scoreCalcRule = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID) == null ? null : ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(var.StudentID).ScoreCalcRuleElement;
                if (scoreCalcRule == null)
                {
                }
                else
                {
                    DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                    foreach (XmlElement element in helper.GetElements("及格標準/學生類別"))
                    {
                        string cat = element.GetAttribute("類別");
                        bool useful = false;
                        //掃描學生的類別作比對
                        foreach (CategoryInfo catinfo in var.StudentCategorys)
                        {
                            if (catinfo.Name == cat || catinfo.FullName == cat)
                                useful = true;
                        }
                        //學生是指定的類別或類別為"預設"
                        if (cat == "預設" || useful)
                        {
                            decimal tryParseDecimal;
                            for (int gyear = 1; gyear <= 4; gyear++)
                            {
                                switch (gyear)
                                {
                                    case 1:
                                        if (decimal.TryParse(element.GetAttribute("一年級補考標準"), out tryParseDecimal))
                                        {
                                            if (!resitLimit.ContainsKey(gyear))
                                                resitLimit.Add(gyear, tryParseDecimal);
                                            if (resitLimit[gyear] > tryParseDecimal)
                                                resitLimit[gyear] = tryParseDecimal;
                                        }
                                        break;
                                    case 2:
                                        if (decimal.TryParse(element.GetAttribute("二年級補考標準"), out tryParseDecimal))
                                        {
                                            if (!resitLimit.ContainsKey(gyear))
                                                resitLimit.Add(gyear, tryParseDecimal);
                                            if (resitLimit[gyear] > tryParseDecimal)
                                                resitLimit[gyear] = tryParseDecimal;
                                        }
                                        break;
                                    case 3:
                                        if (decimal.TryParse(element.GetAttribute("三年級補考標準"), out tryParseDecimal))
                                        {
                                            if (!resitLimit.ContainsKey(gyear))
                                                resitLimit.Add(gyear, tryParseDecimal);
                                            if (resitLimit[gyear] > tryParseDecimal)
                                                resitLimit[gyear] = tryParseDecimal;
                                        }
                                        break;
                                    case 4:
                                        if (decimal.TryParse(element.GetAttribute("四年級補考標準"), out tryParseDecimal))
                                        {
                                            if (!resitLimit.ContainsKey(gyear))
                                                resitLimit.Add(gyear, tryParseDecimal);
                                            if (resitLimit[gyear] > tryParseDecimal)
                                                resitLimit[gyear] = tryParseDecimal;
                                        }
                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                    }
                }
                #endregion
                foreach (SemesterSubjectScoreInfo score in var.SemesterSubjectScoreList)
                {
                    bool canResit = false;
                    decimal s = 0;
                    //  decimal limit = 40;
                    decimal limit = 0;
                    if (decimal.TryParse(score.Detail.GetAttribute("原始成績"), out s))
                    {
                        //if (resitLimit.ContainsKey(score.GradeYear)) limit = resitLimit[score.GradeYear];
                        //canResit = (s >= limit);

                        decimal rs = 0;
                        if (decimal.TryParse(score.Detail.GetAttribute("修課補考標準"), out rs))
                        {
                            canResit = (s >= rs);
                        }

                    }
                    score.Detail.SetAttribute("達補考標準", canResit ? "是" : "否");
                    score.Detail.SetAttribute("補考標準", limit.ToString());
                }
            }
        }

        /// <summary>
        /// 刪除重讀成績
        /// </summary>
        /// <param name="subjectScoreList"></param>
        private void CleanUpRepeat(Dictionary<SemesterSubjectScoreInfo, string> subjectScoreList)
        {
            #region 刪除重讀成績
            Dictionary<int, Dictionary<int, int>> ApplySemesterSchoolYear = new Dictionary<int, Dictionary<int, int>>();
            //先掃一遍抓出每個年級最高的學年度
            foreach (SemesterSubjectScoreInfo scoreInfo in subjectScoreList.Keys)
            {
                if (!ApplySemesterSchoolYear.ContainsKey(scoreInfo.GradeYear))
                    ApplySemesterSchoolYear.Add(scoreInfo.GradeYear, new Dictionary<int, int>());
                if (!ApplySemesterSchoolYear[scoreInfo.GradeYear].ContainsKey(scoreInfo.Semester))
                    ApplySemesterSchoolYear[scoreInfo.GradeYear].Add(scoreInfo.Semester, scoreInfo.SchoolYear);
                if (scoreInfo.SchoolYear > ApplySemesterSchoolYear[scoreInfo.GradeYear][scoreInfo.Semester])
                    ApplySemesterSchoolYear[scoreInfo.GradeYear][scoreInfo.Semester] = scoreInfo.SchoolYear;
            }
            //如果成績資料的年級學年度不在清單中就移掉
            List<SemesterSubjectScoreInfo> removeList = new List<SemesterSubjectScoreInfo>();
            foreach (SemesterSubjectScoreInfo scoreInfo in subjectScoreList.Keys)
            {
                if (ApplySemesterSchoolYear[scoreInfo.GradeYear][scoreInfo.Semester] != scoreInfo.SchoolYear)
                    removeList.Add(scoreInfo);
            }
            foreach (SemesterSubjectScoreInfo scoreInfo in removeList)
            {
                subjectScoreList.Remove(scoreInfo);
            }
            #endregion
        }
    }
}