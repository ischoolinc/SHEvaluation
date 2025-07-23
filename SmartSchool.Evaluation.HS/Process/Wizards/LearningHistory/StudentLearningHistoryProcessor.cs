using FISCA.Data;
using SHSchool.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SmartSchool.Evaluation.Process.Wizards.LearningHistory
{
    public class StudentLearningHistoryProcessor
    {
        private LearningHistoryDataAccess _learningHistoryDataAccess;

        private int _SchoolYear = 0, _Semester = 0;

        public StudentLearningHistoryProcessor()
        {
            _learningHistoryDataAccess = new LearningHistoryDataAccess();
        }

        public void ProcessLearningHistory(SmartSchool.Customization.Data.AccessHelper accHelper, List<SmartSchool.Customization.Data.StudentRecord> StudentRecList, int schoolYear, int semester, BackgroundWorker bgWorker)
        {
            _SchoolYear = schoolYear;
            _Semester = semester;

            List<SubjectScoreRec108> SubjectScoreRec108List = new List<SubjectScoreRec108>();
            List<SubjectScoreRec108> SubjectScoreRec108OtherList = new List<SubjectScoreRec108>();
            //List<SubjectScoreRec108> SubjectScoreRec108ReScoreList = new List<SubjectScoreRec108>();

            //List<SubjectScoreRec108> SubjectScoreRec108List1 = new List<SubjectScoreRec108>();
            //List<SubjectScoreRec108> SubjectScoreRec108List2 = new List<SubjectScoreRec108>();

            List<SubjectScoreRec108> SubjectScoreRec108ListN = new List<SubjectScoreRec108>();

            // 科目補考
            List<SubjectYearScoreRec108N> SubjectReScoreRec108ListN = new List<SubjectYearScoreRec108N>();

            List<SubjectScoreRec108> SubjectScoreRec108OtherListN = new List<SubjectScoreRec108>();
            // 學生學年成績
            Dictionary<string, Dictionary<string, decimal>> StudentYearScoreDict = new Dictionary<string, Dictionary<string, decimal>>();



            //     SmartSchool.Customization.Data.AccessHelper accHelper = new SmartSchool.Customization.Data.AccessHelper();

            // 取得學生ID
            List<string> studentIDList = StudentRecList.Select(x => x.StudentID).ToList();

            //// 取得所選學生資料
            //List<SmartSchool.Customization.Data.StudentRecord> StudentRecList = accHelper.StudentHelper.GetStudents(studentIDList);

            // 依年級批
            Dictionary<string, List<string>> ClassStudentDict = new Dictionary<string, List<string>>();

            Dictionary<string, string> StudGradYearDict = new Dictionary<string, string>();

            Dictionary<string, string> StudGDCCodeDict = new Dictionary<string, string>();

            List<string> StudentIDList = new List<string>();

            foreach (SmartSchool.Customization.Data.StudentRecord rec in StudentRecList)
            {
                StudentIDList.Add(rec.StudentID);

                if (!StudGradYearDict.ContainsKey(rec.StudentID))
                    if (rec.RefClass != null)
                        StudGradYearDict.Add(rec.StudentID, rec.RefClass.GradeYear);

                string cla = "n";
                if (rec.RefClass != null)
                    cla = rec.RefClass.ClassName;

                if (!ClassStudentDict.ContainsKey(cla))
                    ClassStudentDict.Add(cla, new List<string>());

                ClassStudentDict[cla].Add(rec.StudentID);

            }


            // 取得異動與身分別對照
            Dictionary<string, string> UpdateCodeMappingDict = Utility.GetUpdateCodeMappingDict();

            // 取得有符合對照學生
            Dictionary<string, string> StudentHasUpdateCodeDict = Utility.GetStudentHasUpdateCodeDict(_SchoolYear, _Semester, studentIDList, UpdateCodeMappingDict.Keys.ToList());


            //// 取得補修資料學生
            //Dictionary<string, Dictionary<string, SubjectScoreXML>> StudentSubjectReScoreDict = new Dictionary<string, Dictionary<string, SubjectScoreXML>>();

            //// 取得畫面上學年度學期補修資料
            //QueryHelper qhStudSemsReScore = new QueryHelper();
            //string queryReScore = string.Format(@"
            //WITH sems_subj_score AS(
            //    SELECT
            //        sems_subj_score_ext.ref_student_id AS student_id,
            //        sems_subj_score_ext.grade_year,
            //        sems_subj_score_ext.semester,
            //        sems_subj_score_ext.school_year,
            //        sems_subj_score_ext.score_info,
            //        array_to_string(xpath('//Subject/@是否補修成績', subj_score_ele), '') :: text AS 是否補修成績,
            //        array_to_string(xpath('//Subject/@補修學年度', subj_score_ele), '') :: text AS 補修學年度,
            //        array_to_string(xpath('//Subject/@補修學期', subj_score_ele), '') :: text AS 補修學期,
            //        array_to_string(xpath('//Subject/@修課科目代碼', subj_score_ele), '') :: text AS 課程代碼
            //    FROM
            //        (
            //            SELECT
            //                sems_subj_score.*,
            //                unnest(
            //                    xpath(
            //                        '//SemesterSubjectScoreInfo/Subject',
            //                        xmlparse(content score_info)
            //                    )
            //                ) AS subj_score_ele
            //            FROM
            //                sems_subj_score
            //            WHERE
            //                ref_student_id IN ({0})
            //        ) AS sems_subj_score_ext
            //)
            //SELECT
            //    *
            //FROM
            //    sems_subj_score
            //WHERE
            //    是否補修成績 = '是'
            //    AND 補修學年度 = '{1}'
            //    AND 補修學期 = '{2}';
            //", string.Join(",", StudentIDList.ToArray()), _SchoolYear, _Semester);

            //DataTable dtReScore = qhStudSemsReScore.Select(queryReScore);
            //foreach (DataRow dr in dtReScore.Rows)
            //{
            //    string student_id = dr["student_id"] + "";

            //    SubjectScoreXML ssx = new SubjectScoreXML();
            //    ssx.StudentID = student_id;
            //    ssx.SchoolYear = dr["school_year"] + "";
            //    ssx.Semester = dr["semester"] + "";
            //    ssx.GradeYear = dr["grade_year"] + "";


            //    string key = ssx.SchoolYear + "_" + ssx.Semester;

            //    XElement elm = null;
            //    try
            //    {
            //        elm = XElement.Parse(dr["score_info"] + "");
            //        ssx.ScoreXML = elm;
            //    }
            //    catch (Exception ex)
            //    {
            //        Console.WriteLine(ex.Message);
            //    }
            //    if (!StudentSubjectReScoreDict.ContainsKey(student_id))
            //        StudentSubjectReScoreDict.Add(student_id, new Dictionary<string, SubjectScoreXML>());

            //    if (!StudentSubjectReScoreDict[student_id].ContainsKey(key))
            //        StudentSubjectReScoreDict[student_id].Add(key, ssx);
            //}

            // 有符合異動學生非當學年度學期資料
            Dictionary<string, Dictionary<string, SubjectScoreXML>> StudentSubjectScoreOtherDict = new Dictionary<string, Dictionary<string, SubjectScoreXML>>();

            if (StudentHasUpdateCodeDict.Count > 0)
            {
                QueryHelper qhStudSemsScoreOther = new QueryHelper();
                foreach (string className in ClassStudentDict.Keys)
                {
                    List<string> studIDList = ClassStudentDict[className];
                    if (studIDList.Count > 0)
                    {
                        string query = string.Format(@"
                        SELECT
                            ref_student_id AS student_id,
                            school_year,
                            semester,
                            grade_year,
                            score_info
                        FROM
                            sems_subj_score
                        WHERE
                            ref_student_id IN({0}) 
                        ", string.Join(",", StudentHasUpdateCodeDict.Keys.ToArray()));


                        DataTable dtSemsScore = qhStudSemsScoreOther.Select(query);
                        foreach (DataRow dr in dtSemsScore.Rows)
                        {
                            string student_id = dr["student_id"] + "";

                            SubjectScoreXML ssx = new SubjectScoreXML();
                            ssx.StudentID = student_id;
                            ssx.SchoolYear = dr["school_year"] + "";
                            ssx.Semester = dr["semester"] + "";
                            ssx.GradeYear = dr["grade_year"] + "";

                            int xx = int.Parse(ssx.SchoolYear) * 10 + int.Parse(ssx.Semester);

                            if (xx >= (_SchoolYear * 10 + _Semester))
                                continue;

                            string key = ssx.SchoolYear + "_" + ssx.Semester;

                            XElement elm = null;
                            try
                            {
                                elm = XElement.Parse(dr["score_info"].ToString());
                                ssx.ScoreXML = elm;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                            if (!StudentSubjectScoreOtherDict.ContainsKey(student_id))
                                StudentSubjectScoreOtherDict.Add(student_id, new Dictionary<string, SubjectScoreXML>());

                            if (!StudentSubjectScoreOtherDict[student_id].ContainsKey(key))
                                StudentSubjectScoreOtherDict[student_id].Add(key, ssx);

                        }
                    }
                }
            }



            //// 取得學生學期科目成績
            // 因為這寫法在學生人數超過1200人，Client Out off memory,需要用SQL 直接取資料
            //accHelper.StudentHelper.FillSemesterSubjectScore(true, StudentRecList);
            Dictionary<string, Dictionary<string, SubjectScoreXML>> StudentSubjectScoreDict = new Dictionary<string, Dictionary<string, SubjectScoreXML>>();

            // 依年級分批取得學生學期科目成績
            QueryHelper qhStudSemsScore = new QueryHelper();
            foreach (string className in ClassStudentDict.Keys)
            {
                List<string> studIDList = ClassStudentDict[className];
                if (studIDList.Count > 0)
                {
                    string query = string.Format(@"
                    SELECT
                        ref_student_id AS student_id,
                        school_year,
                        semester,
                        grade_year,
                        score_info
                    FROM
                        sems_subj_score
                    WHERE
                        school_year = {0}
                        AND semester = {1}
                        AND ref_student_id IN({2});
                    ", _SchoolYear, _Semester, string.Join(",", studIDList.ToArray()));

                    DataTable dtSemsScore = qhStudSemsScore.Select(query);
                    foreach (DataRow dr in dtSemsScore.Rows)
                    {
                        string student_id = "" + dr["student_id"];

                        SubjectScoreXML ssx = new SubjectScoreXML();
                        ssx.StudentID = student_id;
                        ssx.SchoolYear = dr["school_year"] + "";
                        ssx.Semester = dr["semester"] + "";
                        ssx.GradeYear = dr["grade_year"] + "";

                        string key = ssx.SchoolYear + "_" + ssx.Semester;

                        XElement elm = null;
                        try
                        {
                            elm = XElement.Parse(dr["score_info"] + "");
                            ssx.ScoreXML = elm;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        if (!StudentSubjectScoreDict.ContainsKey(student_id))
                            StudentSubjectScoreDict.Add(student_id, new Dictionary<string, SubjectScoreXML>());

                        if (!StudentSubjectScoreDict[student_id].ContainsKey(key))
                            StudentSubjectScoreDict[student_id].Add(key, ssx);

                    }
                }
            }

            // 讀取補修資料
            Dictionary<string, List<SubjectScoreRec108>> dataValue43 = GetLearningHistoryReScoreDataAsDictionary43(_SchoolYear, _Semester, studentIDList);

            // 取得有補修資料的學期成績
            Dictionary<string, List<SubjectScoreRec108>> semsScore43 = GetStudentScoreDataAsDictionary(dataValue43.Keys.ToList());

            // 比對並填入補考成績資料
            foreach (var studentData in dataValue43)
            {
                string studentID = studentData.Key;
                List<SubjectScoreRec108> dataValueList = studentData.Value;

                // 檢查是否有對應的學期成績資料
                if (semsScore43.ContainsKey(studentID))
                {
                    List<SubjectScoreRec108> semsScoreList = semsScore43[studentID];

                    // 比對每個補修記錄
                    foreach (var dataValueRecord in dataValueList)
                    {
                        // 在學期成績中尋找對應的記錄
                        var matchingSemsScore = semsScoreList.FirstOrDefault(s => 
                            s.StudentID == dataValueRecord.StudentID &&
                            s.SubjectName == dataValueRecord.SubjectName &&
                            s.SubjectLevel == dataValueRecord.SubjectLevel);

                        if (matchingSemsScore != null)
                        {
                            // 填入補考成績
                            dataValueRecord.ReScore = matchingSemsScore.ReScore;

                            // 比較補考成績與及格標準
                            if (!string.IsNullOrWhiteSpace(matchingSemsScore.ReScore) && 
                                !string.IsNullOrWhiteSpace(matchingSemsScore.ScoreP))
                            {
                                // 嘗試轉換為數值進行比較
                                if (decimal.TryParse(matchingSemsScore.ReScore, out decimal reScore) &&
                                    decimal.TryParse(matchingSemsScore.ScoreP, out decimal scoreP))
                                {
                                    if (reScore >= scoreP)
                                    {
                                        dataValueRecord.ReScoreP = "1"; // 及格
                                    }
                                    else
                                    {
                                        dataValueRecord.ReScoreP = "0"; // 不及格
                                    }
                                }
                                else
                                {
                                    dataValueRecord.ReScoreP = "-1"; // 無法比較
                                }
                            }
                            else
                            {
                                dataValueRecord.ReScoreP = "-1"; // 預設值
                            }
                        }
                        else
                        {
                            // 沒有找到對應的學期成績記錄
                            dataValueRecord.ReScoreP = "-1"; // 預設值
                        }
                    }
                }
            }

            // 處理重讀回資料比對補考後回寫學習歷程重讀資料
            // 讀取重讀資料
            Dictionary<string, List<SubjectScoreRec108>> dataValue53 = GetLearningHistoryRetakeDataAsDictionary53(_SchoolYear, _Semester, studentIDList);

            // 取得有重讀資料的學期成績
            Dictionary<string, List<SubjectScoreRec108>> semsScore53 = GetStudentScoreDataAsDictionary(dataValue53.Keys.ToList());

            // 比對並填入補考成績資料
            foreach (var studentData in dataValue53)
            {
                string studentID = studentData.Key;
                List<SubjectScoreRec108> dataValueList = studentData.Value;

                // 檢查是否有對應的學期成績資料
                if (semsScore53.ContainsKey(studentID))
                {
                    List<SubjectScoreRec108> semsScoreList = semsScore53[studentID];

                    // 比對每個重讀記錄
                    foreach (var dataValueRecord in dataValueList)
                    {
                        // 在學期成績中尋找對應的記錄
                        var matchingSemsScore = semsScoreList.FirstOrDefault(s =>
                            s.StudentID == dataValueRecord.StudentID &&
                            s.SubjectName == dataValueRecord.SubjectName &&
                            s.SubjectLevel == dataValueRecord.SubjectLevel);

                        if (matchingSemsScore != null)
                        {
                            // 填入補考成績
                            dataValueRecord.ReScore = matchingSemsScore.ReScore;

                            // 比較補考成績與及格標準
                            if (!string.IsNullOrWhiteSpace(matchingSemsScore.ReScore) &&
                                !string.IsNullOrWhiteSpace(matchingSemsScore.ScoreP))
                            {
                                // 嘗試轉換為數值進行比較
                                if (decimal.TryParse(matchingSemsScore.ReScore, out decimal reScore) &&
                                    decimal.TryParse(matchingSemsScore.ScoreP, out decimal scoreP))
                                {
                                    if (reScore >= scoreP)
                                    {
                                        dataValueRecord.ReScoreP = "1"; // 及格
                                    }
                                    else
                                    {
                                        dataValueRecord.ReScoreP = "0"; // 不及格
                                    }
                                }
                                else
                                {
                                    dataValueRecord.ReScoreP = "-1"; // 無法比較
                                }
                            }
                            else
                            {
                                dataValueRecord.ReScoreP = "-1"; // 預設值
                            }
                        }
                        else
                        {
                            // 沒有找到對應的學期成績記錄
                            dataValueRecord.ReScoreP = "-1"; // 預設值
                        }
                    }
                }
            }


            // 取得學期對照年為主
            List<SHSemesterHistoryRecord> SemsH = SHSemesterHistory.SelectByStudentIDs(studentIDList);
            Dictionary<string, List<K12.Data.SemesterHistoryItem>> SemsHDict = new Dictionary<string, List<K12.Data.SemesterHistoryItem>>();
            foreach (SHSemesterHistoryRecord rec in SemsH)
            {
                if (!SemsHDict.ContainsKey(rec.RefStudentID))
                    SemsHDict.Add(rec.RefStudentID, rec.SemesterHistoryItems);

                foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                {
                    if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                    {
                        if (!StudGradYearDict.ContainsKey(item.RefStudentID))
                            StudGradYearDict.Add(item.RefStudentID, item.GradeYear.ToString());
                        else
                            StudGradYearDict[item.RefStudentID] = item.GradeYear.ToString();

                        // 取得學期對照群科班代碼
                        if (!StudGDCCodeDict.ContainsKey(item.RefStudentID))
                            StudGDCCodeDict.Add(item.RefStudentID, item.GDCCode);

                    }
                }
            }

            // 取得修課設定再次修習
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
ORDER BY courseName,className, seatNo ASC", _SchoolYear, _Semester, string.Join(",", studentIDList.ToArray()));

            QueryHelper qh1 = new QueryHelper();

            DataTable dt_SCAttend = qh1.Select(query2);

            Dictionary<string, string> duplicateSubjectLevelMethodDict = new Dictionary<string, string>();

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
            }

            Dictionary<string, string> duplicateSubjectLevelMethodDict_Afterfilter = new Dictionary<string, string>(); // 真正過濾後，有重覆科目級別的項目

            // 過濾取得重覆科目級別的計算處理方式
            foreach (var kv in duplicateSubjectLevelMethodDict)
            {
                if (!string.IsNullOrWhiteSpace(kv.Value) && !duplicateSubjectLevelMethodDict_Afterfilter.ContainsKey(kv.Key))
                {
                    duplicateSubjectLevelMethodDict_Afterfilter.Add(kv.Key, kv.Value);
                }
            }




            foreach (SmartSchool.Customization.Data.StudentRecord studRec in StudentRecList)
            {
                string IDNumber = studRec.IDNumber.ToUpper();
                string BirthDate = "";
                DateTime dt;
                if (DateTime.TryParse(studRec.Birthday, out dt))
                    BirthDate = Utility.ConvertChDateString(dt);

                if (StudentSubjectScoreDict.ContainsKey(studRec.StudentID))
                {
                    string smsKey = _SchoolYear + "_" + _Semester;

                    #region 一般與補修

                    if (StudentSubjectScoreDict[studRec.StudentID].ContainsKey(smsKey))
                    {


                        XElement elmRoot = StudentSubjectScoreDict[studRec.StudentID][smsKey].ScoreXML;

                        foreach (XElement elmScore in elmRoot.Elements("Subject"))
                        {

                            SubjectScoreRec108 ssr = new SubjectScoreRec108();
                            ssr.IDNumber = IDNumber.ToUpper();
                            ssr.StudentID = studRec.StudentID;
                            ssr.Birthday = BirthDate;
                            ssr.GradeYear = StudentSubjectScoreDict[studRec.StudentID][smsKey].GradeYear;
                            ssr.SchoolYear = StudentSubjectScoreDict[studRec.StudentID][smsKey].SchoolYear;
                            ssr.Semester = StudentSubjectScoreDict[studRec.StudentID][smsKey].Semester;
                            ssr.SubjectName = Utility.GetAttribute(elmScore, "科目");

                            ssr.SubjectLevel = Utility.GetAttribute(elmScore, "科目級別");

                            ssr.Name = studRec.StudentName;
                            //ssr.ClassName = studRec.RefClass.ClassName;
                            //ssr.SeatNo = studRec.SeatNo;
                            //ssr.StudentNumber = studRec.StudentNumber;

                            string subjectKey = ssr.SubjectName.Trim() + "_" + ssr.SubjectLevel.Trim();
                            string studentSubjectKey = ssr.StudentID + "_" + subjectKey;

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
                                // 再次修習不放入學期成績工作頁
                                continue;
                            }

                            // 不計學分=是，不列入學習成績
                            if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                continue;


                            foreach (SHSemesterHistoryRecord rec in SemsH)
                            {
                                foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                                {
                                    if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                                    {
                                        if (studRec.StudentID == item.RefStudentID)
                                        {
                                            ssr.HisClassName = item.ClassName;
                                            ssr.HisSeatNo = item.SeatNo;
                                            ssr.HisStudentNumber = item.StudentNumber;
                                        }
                                    }
                                }
                            }



                            ssr.CourseCode = Utility.GetAttribute(elmScore, "修課科目代碼");

                            //當課程類別為8(團體活動時間)及9(彈性活動時間)，且科目屬性不為D(充實(增廣)、補強性教學 [全學期、授予學分])時，不允許提交成績。
                            //課程代碼為23碼
                            ssr.CodePass = true;
                            int startIndex1 = 16;
                            int endIndex = 1;
                            int startIndex2 = 18;

                            if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                            {
                                if (ssr.CourseCode.Length > 22)
                                {
                                    string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                                    string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                                    if ((sub1 == "8" || sub1 == "9") && sub2 != "D")
                                    {
                                        ssr.CodePass = false;   //不可提交
                                    }
                                }
                            }


                            ssr.Credit = Utility.GetAttribute(elmScore, "開課學分數");
                            // 預設值 -1
                            ssr.Score = ssr.ScScore = ssr.ReScore = ssr.useCredit = ssr.ScoreP = ssr.ScScoreP = ssr.ReAScoreP = ssr.ReScoreP = "-1";

                            string GrStr = "";
                            if (StudGradYearDict.ContainsKey(studRec.StudentID))
                                GrStr = StudGradYearDict[studRec.StudentID] + "_及";

                            decimal ds, dsre, passScore = 60;

                            if (decimal.TryParse(Utility.GetAttribute(elmScore, "原始成績"), out ds))
                            {
                                // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                                //ds = Math.Round(ds, 0, MidpointRounding.AwayFromZero);

                                decimal dsp;
                                if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsp))
                                    passScore = dsp;

                                // ssr.Score = string.Format("{0:##0}", ds);
                                ssr.Score = ds.ToString();                                
                                
                                if (ds < passScore)
                                    ssr.ScoreP = "0";
                                else
                                    ssr.ScoreP = "1";
                            }
                            
                            ssr.useCredit = "1";

                            if (Utility.GetAttribute(elmScore, "抵免") == "是")
                                ssr.useCredit = "2";

                            if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                            {
                                if (ssr.CourseCode.Length > 22)
                                {
                                    string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                                    string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                                    if (sub1 == "9" && sub2 == "D")
                                    {
                                        ssr.useCredit = "3";
                                    }
                                }
                            }


                            if (decimal.TryParse(Utility.GetAttribute(elmScore, "補考成績"), out dsre))
                            {
                                decimal dsreP;

                                // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                                //dsre = Math.Round(dsre, 0, MidpointRounding.AwayFromZero);

                                if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsreP))
                                {
                                    passScore = dsreP;
                                }

                                // ssr.ReScore = string.Format("{0:##0}", dsre);
                                ssr.ReScore = dsre.ToString();

                                if (dsre < passScore)
                                    ssr.ReScoreP = "0";
                                else
                                    ssr.ReScoreP = "1";
                            }

                            ssr.isScScore = false;
                            ssr.ScScoreType = "3";  // 補考方式預設 3，專班辦理。

                            if (Utility.GetAttribute(elmScore, "是否補修成績") != null)
                            {
                                if (Utility.GetAttribute(elmScore, "是否補修成績") == "是")
                                {
                                    ssr.isScScore = true;
                                }
                            }

                            if (ssr.ScoreP == "-1")
                            {
                                if (Utility.GetAttribute(elmScore, "是否取得學分") != null)
                                {
                                    if (Utility.GetAttribute(elmScore, "是否取得學分") == "是")
                                    {
                                        ssr.ScoreP = "1";
                                    }
                                    else
                                    {
                                        ssr.ScoreP = "0";
                                    }
                                }
                            }

                            // 非補修成績才寫入
                            if (ssr.isScScore == false)
                                SubjectScoreRec108List.Add(ssr);
                        }
                    }

                    #endregion
                }


                if (StudentSubjectScoreOtherDict.ContainsKey(studRec.StudentID))
                {
                    // 轉學轉科
                    foreach (string key in StudentSubjectScoreOtherDict[studRec.StudentID].Keys)
                    {
                        XElement elmRoot = StudentSubjectScoreOtherDict[studRec.StudentID][key].ScoreXML;
                        foreach (XElement elmScore in elmRoot.Elements("Subject"))
                        {
                            SubjectScoreRec108 ssr = new SubjectScoreRec108();
                            ssr.IDNumber = IDNumber.ToUpper();
                            ssr.StudentID = studRec.StudentID;
                            ssr.Birthday = BirthDate;
                            ssr.GradeYear = StudentSubjectScoreOtherDict[studRec.StudentID][key].GradeYear;
                            ssr.SchoolYear = StudentSubjectScoreOtherDict[studRec.StudentID][key].SchoolYear;
                            ssr.Semester = StudentSubjectScoreOtherDict[studRec.StudentID][key].Semester;
                            ssr.SubjectName = Utility.GetAttribute(elmScore, "科目");

                            ssr.SubjectLevel = Utility.GetAttribute(elmScore, "科目級別");

                            ssr.Name = studRec.StudentName;
                            //ssr.ClassName = studRec.RefClass.ClassName;
                            //ssr.SeatNo = studRec.SeatNo;
                            //ssr.StudentNumber = studRec.StudentNumber;

                            foreach (SHSemesterHistoryRecord rec in SemsH)
                            {
                                foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                                {
                                    if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                                    {
                                        if (studRec.StudentID == item.RefStudentID)
                                        {
                                            ssr.HisClassName = item.ClassName;
                                            ssr.HisSeatNo = item.SeatNo;
                                            ssr.HisStudentNumber = item.StudentNumber;
                                        }
                                    }
                                }
                            }

                            ssr.CourseCode = Utility.GetAttribute(elmScore, "修課科目代碼");

                            //當課程類別為8(團體活動時間)及9(彈性活動時間)，且科目屬性不為D(充實(增廣)、補強性教學 [全學期、授予學分])時，不允許提交成績。
                            //課程代碼為23碼
                            ssr.CodePass = true;
                            int startIndex1 = 16;
                            int endIndex = 1;
                            int startIndex2 = 18;

                            if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                            {
                                if (ssr.CourseCode.Length > 22)
                                {
                                    string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                                    string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                                    if ((sub1 == "8" || sub1 == "9") && sub2 != "D")
                                    {
                                        ssr.CodePass = false;   //不可提交
                                    }
                                }
                            }

                            ssr.Credit = Utility.GetAttribute(elmScore, "開課學分數");
                            // 預設值 -1
                            ssr.Score = ssr.ScScore = ssr.ReScore = ssr.useCredit = ssr.ScoreP = ssr.ScScoreP = ssr.ReAScoreP = ssr.ReScoreP = "-1";

                            string GrStr = "";
                            if (StudGradYearDict.ContainsKey(studRec.StudentID))
                                GrStr = StudGradYearDict[studRec.StudentID] + "_及";

                            decimal ds, dsre, passScore = 60;

                            if (decimal.TryParse(Utility.GetAttribute(elmScore, "原始成績"), out ds))
                            {
                                // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                                //ds = Math.Round(ds, 0, MidpointRounding.AwayFromZero);

                                decimal dsp;
                                if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsp))
                                    passScore = dsp;

                                //ssr.Score = string.Format("{0:##0}", ds);
                                ssr.Score = ds.ToString();

                                if (ds < passScore)
                                    ssr.ScoreP = "0";
                                else
                                    ssr.ScoreP = "1";
                            }

                            ssr.useCredit = "1";
                            if (Utility.GetAttribute(elmScore, "抵免") == "否")
                                ssr.useCredit = "2";

                            if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                            {
                                if (ssr.CourseCode.Length > 22)
                                {
                                    string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                                    string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                                    if (sub1 == "9" && sub2 == "D")
                                    {
                                        ssr.useCredit = "3";
                                    }
                                }
                            }



                            if (decimal.TryParse(Utility.GetAttribute(elmScore, "補考成績"), out dsre))
                            {
                                // 四捨五入到整數位--2021年3月 取消處理四捨五入
                                //dsre = Math.Round(dsre, 0, MidpointRounding.AwayFromZero);

                                decimal dsreP;

                                if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsreP))
                                {
                                    passScore = dsreP;
                                }

                                // ssr.ReScore = string.Format("{0:##0}", dsre);
                                ssr.ReScore = dsre.ToString();

                                if (dsre < passScore)
                                    ssr.ReScoreP = "0";
                                else
                                    ssr.ReScoreP = "1";
                            }

                            ssr.isScScore = false;
                            ssr.ScScoreType = "3";  // 補考方式預設 3，專班辦理。

                            if (Utility.GetAttribute(elmScore, "是否補修成績") != null)
                            {
                                if (Utility.GetAttribute(elmScore, "是否補修成績") == "是")
                                {
                                    ssr.isScScore = true;
                                }
                            }

                            // 對應學生身分別
                            if (StudentHasUpdateCodeDict.ContainsKey(studRec.StudentID))
                            {
                                if (UpdateCodeMappingDict.ContainsKey(StudentHasUpdateCodeDict[studRec.StudentID]))
                                {
                                    ssr.StudType = UpdateCodeMappingDict[StudentHasUpdateCodeDict[studRec.StudentID]];
                                }
                            }

                            if (ssr.ScoreP == "-1")
                            {
                                if (Utility.GetAttribute(elmScore, "是否取得學分") != null)
                                {
                                    if (Utility.GetAttribute(elmScore, "是否取得學分") == "是")
                                    {
                                        ssr.ScoreP = "1";
                                    }
                                    else
                                    {
                                        ssr.ScoreP = "0";
                                    }
                                }
                            }

                            // 不計學分 = 是，不列入學習成績
                            if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                continue;

                            // 如果有勾是否補修成績，那補修學年度、學期和畫面上「不同」才會填入，相同的話會在補修成績工作頁 
                            // 2021-11 
                            // https://3.basecamp.com/4399967/buckets/15765350/todos/4305803619
                            if (ssr.isScScore)
                            {
                                if (!(Utility.GetAttribute(elmScore, "補修學年度") == _SchoolYear.ToString() && Utility.GetAttribute(elmScore, "補修學期") == _Semester.ToString()))
                                {
                                    SubjectScoreRec108OtherList.Add(ssr);
                                }
                            }
                            else
                            {
                                //沒有勾是否補修成績，就直接填入
                                SubjectScoreRec108OtherList.Add(ssr);
                            }

                            //SubjectScoreRec108OtherList.Add(ssr);
                        }
                    }
                }


                //// 補修
                //if (StudentSubjectReScoreDict.ContainsKey(studRec.StudentID))
                //{
                //    #region 補修
                //    foreach (string smsKey in StudentSubjectReScoreDict[studRec.StudentID].Keys)
                //        if (StudentSubjectReScoreDict[studRec.StudentID].ContainsKey(smsKey))
                //        {
                //            XElement elmRoot = StudentSubjectReScoreDict[studRec.StudentID][smsKey].ScoreXML;

                //            foreach (XElement elmScore in elmRoot.Elements("Subject"))
                //            {

                //                SubjectScoreRec108 ssr = new SubjectScoreRec108();
                //                ssr.IDNumber = IDNumber.ToUpper();
                //                ssr.StudentID = studRec.StudentID;
                //                ssr.Birthday = BirthDate;
                //                ssr.GradeYear = StudentSubjectReScoreDict[studRec.StudentID][smsKey].GradeYear;
                //                ssr.SchoolYear = StudentSubjectReScoreDict[studRec.StudentID][smsKey].SchoolYear;
                //                ssr.Semester = StudentSubjectReScoreDict[studRec.StudentID][smsKey].Semester;
                //                ssr.SubjectName = Utility.GetAttribute(elmScore, "科目");

                //                ssr.SubjectLevel = Utility.GetAttribute(elmScore, "科目級別");

                //                ssr.Name = studRec.StudentName;
                //                ssr.ClassName = studRec.RefClass.ClassName;
                //                ssr.SeatNo = studRec.SeatNo;
                //                ssr.StudentNumber = studRec.StudentNumber;

                //                foreach (SHSemesterHistoryRecord rec in SemsH)
                //                {
                //                    foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                //                    {
                //                        if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                //                        {
                //                            if (studRec.StudentID == item.RefStudentID)
                //                            {
                //                                ssr.HisClassName = item.ClassName;
                //                                ssr.HisSeatNo = item.SeatNo;
                //                                ssr.HisStudentNumber = item.StudentNumber;
                //                            }
                //                        }
                //                    }
                //                }



                //                ssr.CourseCode = Utility.GetAttribute(elmScore, "修課科目代碼");

                //                //當課程類別為8(團體活動時間)及9(彈性活動時間)，且科目屬性不為D(充實(增廣)、補強性教學 [全學期、授予學分])時，不允許提交成績。
                //                //課程代碼為23碼
                //                ssr.CodePass = true;
                //                int startIndex1 = 16;
                //                int endIndex = 1;
                //                int startIndex2 = 18;

                //                if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                //                {
                //                    if (ssr.CourseCode.Length > 22)
                //                    {
                //                        string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                //                        string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                //                        if ((sub1 == "8" || sub1 == "9") && sub2 != "D")
                //                        {
                //                            ssr.CodePass = false;   //不可提交
                //                        }
                //                    }
                //                }


                //                ssr.Credit = Utility.GetAttribute(elmScore, "開課學分數");
                //                // 預設值 -1
                //                ssr.Score = ssr.ScScore = ssr.ReScore = ssr.useCredit = ssr.ScoreP = ssr.ScScoreP = ssr.ReAScoreP = ssr.ReScoreP = "-1";

                //                string GrStr = "";
                //                if (StudGradYearDict.ContainsKey(studRec.StudentID))
                //                    GrStr = StudGradYearDict[studRec.StudentID] + "_及";

                //                decimal ds, dsre, passScore = 60;

                //                if (decimal.TryParse(Utility.GetAttribute(elmScore, "原始成績"), out ds))
                //                {
                //                    // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                //                    //ds = Math.Round(ds, 0, MidpointRounding.AwayFromZero);

                //                    decimal dsp;
                //                    if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsp))
                //                        passScore = dsp;

                //                    // ssr.Score = string.Format("{0:##0}", ds);
                //                    ssr.Score = ds.ToString();

                //                    //2021年4月

                //                    //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                //                    // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                //                    if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                //                    {
                //                        ssr.useCredit = "1";

                //                        if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                //                        {
                //                            if (ssr.CourseCode.Length > 22)
                //                            {
                //                                string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                //                                string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                //                                if (sub1 == "9" && sub2 == "D")
                //                                {
                //                                    ssr.useCredit = "3";
                //                                }
                //                            }
                //                        }
                //                        // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                //                        if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                //                            ssr.useCredit = "3";
                //                    }
                //                    else
                //                    {
                //                        if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                //                            ssr.useCredit = "2";
                //                    }

                //                    if (ds < passScore)
                //                        ssr.ScoreP = "0";
                //                    else
                //                        ssr.ScoreP = "1";
                //                }
                //                else
                //                {

                //                    //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                //                    // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                //                    if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                //                    {
                //                        ssr.useCredit = "1";

                //                        if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                //                        {
                //                            if (ssr.CourseCode.Length > 22)
                //                            {
                //                                string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                //                                string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                //                                if (sub1 == "9" && sub2 == "D")
                //                                {
                //                                    ssr.useCredit = "3";
                //                                }
                //                            }
                //                        }

                //                        // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                //                        if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                //                            ssr.useCredit = "3";
                //                    }
                //                    else
                //                    {
                //                        if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                //                            ssr.useCredit = "2";
                //                    }

                //                    //// 判斷是否 不需評分
                //                    //if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                //                    //{
                //                    //    ssr.useCredit = "3";
                //                    //}
                //                }

                //                if (decimal.TryParse(Utility.GetAttribute(elmScore, "補考成績"), out dsre))
                //                {
                //                    decimal dsreP;

                //                    // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                //                    //dsre = Math.Round(dsre, 0, MidpointRounding.AwayFromZero);

                //                    if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsreP))
                //                    {
                //                        passScore = dsreP;
                //                    }

                //                    // ssr.ReScore = string.Format("{0:##0}", dsre);
                //                    ssr.ReScore = dsre.ToString();

                //                    if (dsre < passScore)
                //                        ssr.ReScoreP = "0";
                //                    else
                //                        ssr.ReScoreP = "1";
                //                }

                //                ssr.isScScore = false;
                //                ssr.ScScoreType = "3";  // 補考方式預設 3，專班辦理。

                //                if (Utility.GetAttribute(elmScore, "是否補修成績") != null)
                //                {
                //                    if (Utility.GetAttribute(elmScore, "是否補修成績") == "是")
                //                    {
                //                        ssr.isScScore = true;
                //                    }
                //                }

                //                if (ssr.ScoreP == "-1")
                //                {
                //                    if (Utility.GetAttribute(elmScore, "是否取得學分") != null)
                //                    {
                //                        if (Utility.GetAttribute(elmScore, "是否取得學分") == "是")
                //                        {
                //                            ssr.ScoreP = "1";
                //                        }
                //                        else
                //                        {
                //                            ssr.ScoreP = "0";
                //                        }
                //                    }
                //                }


                //                // 補修學年度、學期和畫面上選相同相同才會填入
                //                if (Utility.GetAttribute(elmScore, "補修學年度") == _SchoolYear.ToString() && Utility.GetAttribute(elmScore, "補修學期") == _Semester.ToString())
                //                {
                //                    SubjectScoreRec108ReScoreList.Add(ssr);
                //                }

                //            }
                //        }

                //    #endregion
                //}

            }


            // 處理重修重讀名冊  ------
            // For 處理重修學生 群科班對照
            Dictionary<string, Dictionary<string, string>> StudHistoryGDCCodeDict = new Dictionary<string, Dictionary<string, string>>();

            // 2021-11-04 尋找需要抵免的異動 (復學/轉科/重讀)
            Dictionary<string, string> CreditUpdateCodeMappingDict = Utility.GetUpdateCodeMappingDict3();
            //找最後一個學年度學期 (Dictionary student/1091) 
            Dictionary<string, string> StudentCreditUpdateCodeDict = Utility.GetStudentHasUpdateCodeDict(studentIDList, CreditUpdateCodeMappingDict.Keys.ToList());

            // 取得異動與身分別對照
            UpdateCodeMappingDict = Utility.GetUpdateCodeMappingDict2();

            // 取得有符合對照學生
            StudentHasUpdateCodeDict = Utility.GetStudentHasUpdateCodeDict(_SchoolYear, _Semester, studentIDList, UpdateCodeMappingDict.Keys.ToList());

            //// 重修
            //Dictionary<string, Dictionary<string, SubjectScoreXML>> StudentSubjectScore1Dict = new Dictionary<string, Dictionary<string, SubjectScoreXML>>();

            // 重讀
            Dictionary<string, Dictionary<string, SubjectScoreXML>> StudentSubjectScore2Dict = new Dictionary<string, Dictionary<string, SubjectScoreXML>>();

            if (StudentHasUpdateCodeDict.Count > 0)
            {
                // 重讀
                QueryHelper qhStudSemsScore2 = new QueryHelper();
                foreach (string className in ClassStudentDict.Keys)
                {
                    List<string> studIDList = ClassStudentDict[className];
                    if (studIDList.Count > 0)
                    {
                        string query = string.Format(@"
                        SELECT
                            ref_student_id AS student_id,
                            school_year,
                            semester,
                            grade_year,
                            score_info
                        FROM
                            sems_subj_score
                        WHERE
                            school_year = {0}
                            AND semester = {1}
                            AND ref_student_id IN({2});
                        ", _SchoolYear, _Semester, string.Join(",", StudentHasUpdateCodeDict.Keys.ToArray()));

                        DataTable dtSemsScore = qhStudSemsScore2.Select(query);
                        foreach (DataRow dr in dtSemsScore.Rows)
                        {
                            string student_id = dr["student_id"] + "";

                            SubjectScoreXML ssx = new SubjectScoreXML();
                            ssx.StudentID = student_id;
                            ssx.SchoolYear = dr["school_year"] + "";
                            ssx.Semester = dr["semester"] + "";
                            ssx.GradeYear = dr["grade_year"] + "";

                            string key = ssx.SchoolYear + "_" + ssx.Semester;

                            XElement elm = null;
                            try
                            {
                                elm = XElement.Parse(dr["score_info"] + "");
                                ssx.ScoreXML = elm;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                            if (!StudentSubjectScore2Dict.ContainsKey(student_id))
                                StudentSubjectScore2Dict.Add(student_id, new Dictionary<string, SubjectScoreXML>());

                            if (!StudentSubjectScore2Dict[student_id].ContainsKey(key))
                                StudentSubjectScore2Dict[student_id][key] = ssx;

                        }
                    }
                }
            }

            //// 依年級分批取得學生學期科目成績 比對重修
            //foreach (string className in ClassStudentDict.Keys)
            //{
            //    List<string> studIDList = ClassStudentDict[className];
            //    if (studIDList.Count > 0)
            //    {
            //        string query = string.Format(@"
            //        SELECT
            //            ref_student_id AS student_id,
            //            school_year,
            //            semester,
            //            grade_year,
            //            score_info
            //        FROM
            //            sems_subj_score
            //        WHERE
            //            ref_student_id IN({0});
            //        ", string.Join(",", studIDList.ToArray()));

            //        DataTable dtSemsScore = qhStudSemsScore.Select(query);
            //        foreach (DataRow dr in dtSemsScore.Rows)
            //        {
            //            string student_id = dr["student_id"] + "";

            //            SubjectScoreXML ssx = new SubjectScoreXML();
            //            ssx.StudentID = student_id;
            //            ssx.SchoolYear = dr["school_year"] + "";
            //            ssx.Semester = dr["semester"] + "";
            //            ssx.GradeYear = dr["grade_year"] + "";

            //            string key = ssx.SchoolYear + "_" + ssx.Semester;

            //            XElement elm = null;
            //            try
            //            {
            //                elm = XElement.Parse(dr["score_info"] + "");
            //                ssx.ScoreXML = elm;
            //            }
            //            catch (Exception ex)
            //            {
            //                Console.WriteLine(ex.Message);
            //            }
            //            if (!StudentSubjectScore1Dict.ContainsKey(student_id))
            //                StudentSubjectScore1Dict.Add(student_id, new Dictionary<string, SubjectScoreXML>());

            //            if (!StudentSubjectScore1Dict[student_id].ContainsKey(key))
            //                StudentSubjectScore1Dict[student_id][key] = ssx;

            //        }
            //    }
            //}


            foreach (SmartSchool.Customization.Data.StudentRecord studRec in StudentRecList)
            {
                string IDNumber = studRec.IDNumber.ToUpper();
                string BirthDate = "";
                DateTime dt;
                if (DateTime.TryParse(studRec.Birthday, out dt))
                    BirthDate = Utility.ConvertChDateString(dt);

                //if (StudentSubjectScore1Dict.ContainsKey(studRec.StudentID))
                //{
                //    // string smsKey = _SchoolYear + "_" + _Semester;

                //    #region 重修

                //    foreach (string smsKey in StudentSubjectScore1Dict[studRec.StudentID].Keys)
                //    {
                //        if (StudentSubjectScore1Dict[studRec.StudentID].ContainsKey(smsKey))
                //        {
                //            XElement elmRoot = StudentSubjectScore1Dict[studRec.StudentID][smsKey].ScoreXML;

                //            foreach (XElement elmScore in elmRoot.Elements("Subject"))
                //            {
                //                SubjectScoreRec108 ssr = new SubjectScoreRec108();
                //                ssr.IDNumber = IDNumber;
                //                ssr.StudentID = studRec.StudentID;
                //                ssr.Birthday = BirthDate;
                //                ssr.GradeYear = StudentSubjectScore1Dict[studRec.StudentID][smsKey].GradeYear;
                //                ssr.SchoolYear = StudentSubjectScore1Dict[studRec.StudentID][smsKey].SchoolYear;  //原修課學年度
                //                ssr.Semester = StudentSubjectScore1Dict[studRec.StudentID][smsKey].Semester;  //原修課學期
                //                ssr.SubjectName = Utility.GetAttribute(elmScore, "科目");
                //                ssr.SubjectLevel = Utility.GetAttribute(elmScore, "科目級別");
                //                ssr.Name = studRec.StudentName;
                //                ssr.ClassName = studRec.RefClass.ClassName;
                //                ssr.SeatNo = studRec.SeatNo;
                //                ssr.StudentNumber = studRec.StudentNumber;


                //                foreach (SHSemesterHistoryRecord rec in SemsH)
                //                {
                //                    foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                //                    {
                //                        //if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                //                        //{
                //                        //    if (studRec.StudentID == item.RefStudentID)
                //                        //    {
                //                        //        ssr.HisClassName = item.ClassName;
                //                        //        ssr.HisSeatNo = item.SeatNo;
                //                        //        ssr.HisStudentNumber = item.StudentNumber;
                //                        //    }
                //                        //}

                //                        if (item.SchoolYear.ToString() == ssr.SchoolYear && item.Semester.ToString() == ssr.Semester)
                //                        {
                //                            if (studRec.StudentID == item.RefStudentID)
                //                            {
                //                                ssr.HisClassName = item.ClassName;
                //                                ssr.HisSeatNo = item.SeatNo;
                //                                ssr.HisStudentNumber = item.StudentNumber;


                //                            }
                //                        }

                //                        /// 2021-11-05 因國教署系統判斷問題，故呈報重修時，課程代碼要填入原習修學年期在使用的課程代碼
                //                        // 將每個學年期的學期對照表GDCCode存入Dictionary
                //                        string key = item.RefStudentID + "_" + item.SchoolYear.ToString() + "_" + item.Semester.ToString();

                //                        if (!StudHistoryGDCCodeDict.ContainsKey(key))
                //                        {
                //                            // 要傳入 StudHistoryGDCCodeDict 的 Dictionary
                //                            Dictionary<string, string> paramForStudHistoryGDCCodeDict = new Dictionary<string, string>();
                //                            if (!paramForStudHistoryGDCCodeDict.ContainsKey(item.RefStudentID))
                //                            {
                //                                paramForStudHistoryGDCCodeDict.Add(item.RefStudentID, item.GDCCode);
                //                            }
                //                            StudHistoryGDCCodeDict.Add(key, paramForStudHistoryGDCCodeDict);
                //                        }
                //                    }
                //                }

                //                string dicKey = studRec.StudentID + "_" + ssr.SchoolYear + "_" + ssr.Semester;
                //                string subjectCode = "";

                //                if (Utility.GetAttribute(elmScore, "重修學年度") == _SchoolYear.ToString() && Utility.GetAttribute(elmScore, "重修學期") == _Semester.ToString())
                //                {
                //                    /// 2021-11-05 如果有遇到復學/重讀/轉入，會將過去的課程代碼抵免成新的課程代碼，
                //                    if (StudentCreditUpdateCodeDict.ContainsKey(studRec.StudentID))
                //                    {
                //                        // 最後一筆異動學年期
                //                        int intUpdateYear_Sems = int.Parse(StudentCreditUpdateCodeDict[studRec.StudentID]);
                //                        string strtUpdateYear_Sems = StudentCreditUpdateCodeDict[studRec.StudentID];
                //                        //重修 原習修學年期
                //                        int strSsrYear_Sems = int.Parse(ssr.SchoolYear + ssr.Semester);

                //                        string updateYear = strtUpdateYear_Sems.Substring(0, 3);
                //                        string updateSems = strtUpdateYear_Sems.Substring(3, 1);

                //                        //https://3.basecamp.com/4399967/buckets/15765350/todos/4314350115#__recording_4323437970
                //                        /// 所以 當 原習修學年期 <= 異動學年期，2022/09/21 且 異動學年期<=選擇的學年期，這裡的課程代碼已經抵免了，就填寫異動學年期的課程代碼
                //                        if (strSsrYear_Sems <= intUpdateYear_Sems && intUpdateYear_Sems <= int.Parse(_SchoolYear.ToString() + _Semester.ToString()))
                //                        {
                //                            dicKey = studRec.StudentID + "_" + updateYear + "_" + updateSems;
                //                        }
                //                    }
                //                    ssr.CourseCode = Utility.GetAttribute(elmScore, "修課科目代碼");



                //                }


                //                //當課程類別為8(團體活動時間)及9(彈性活動時間)，且科目屬性不為D(充實(增廣)、補強性教學 [全學期、授予學分])時，不允許提交成績。
                //                //課程代碼為23碼
                //                ssr.CodePass = true;
                //                int startIndex1 = 16;
                //                int endIndex = 1;
                //                int startIndex2 = 18;

                //                if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                //                {
                //                    if (ssr.CourseCode.Length > 22)
                //                    {
                //                        string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                //                        string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                //                        if ((sub1 == "8" || sub1 == "9") && sub2 != "D")
                //                        {
                //                            ssr.CodePass = false;   //不可提交
                //                        }
                //                    }
                //                }

                //                ssr.Credit = Utility.GetAttribute(elmScore, "開課學分數");
                //                // 預設值 -1
                //                ssr.Score = ssr.ScScore = ssr.ReScore = ssr.useCredit = ssr.ScoreP = ssr.ScScoreP = ssr.ReAScoreP = ssr.ReScoreP = "-1";

                //                string GrStr = "";
                //                if (StudGradYearDict.ContainsKey(studRec.StudentID))
                //                    GrStr = StudGradYearDict[studRec.StudentID] + "_及";

                //                decimal ds, dsre, passScore = 60;

                //                if (decimal.TryParse(Utility.GetAttribute(elmScore, "重修成績"), out ds))
                //                {
                //                    // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                //                    //ds = Math.Round(ds, 0, MidpointRounding.AwayFromZero);

                //                    decimal dsp;
                //                    if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsp))
                //                        passScore = dsp;

                //                    //ssr.ReAScore = string.Format("{0:##0}", ds);
                //                    ssr.ReAScore = ds.ToString();

                //                    // if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                //                    // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                //                    if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                //                    {
                //                        ssr.useCredit = "1";

                //                        if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                //                        {
                //                            if (ssr.CourseCode.Length > 22)
                //                            {
                //                                string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                //                                string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                //                                if (sub1 == "9" && sub2 == "D")
                //                                {
                //                                    ssr.useCredit = "3";
                //                                }
                //                            }
                //                        }
                //                        // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                //                        if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                //                            ssr.useCredit = "3";
                //                    }
                //                    else
                //                    {
                //                        if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                //                            ssr.useCredit = "2";
                //                    }

                //                    if (ds < passScore)
                //                        ssr.ReAScoreP = "0";
                //                    else
                //                    {
                //                        ssr.ReAScoreP = "1";

                //                        //2021-11-09 重修成績最高分為及格標準，若沒有設定及格標準則以60分計算。
                //                        ssr.ReAScore = passScore.ToString();
                //                    }

                //                }
                //                else
                //                {
                //                    //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                //                    // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                //                    if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                //                    {
                //                        ssr.useCredit = "1";

                //                        if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                //                        {
                //                            if (ssr.CourseCode.Length > 22)
                //                            {
                //                                string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                //                                string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                //                                if (sub1 == "9" && sub2 == "D")
                //                                {
                //                                    ssr.useCredit = "3";
                //                                }
                //                            }
                //                        }
                //                        // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                //                        if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                //                            ssr.useCredit = "3";
                //                    }
                //                    else
                //                    {
                //                        if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                //                            ssr.useCredit = "2";
                //                    }

                //                    // 判斷是否 不需評分
                //                    //if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                //                    //{
                //                    //    ssr.useCredit = "3";
                //                    //}
                //                }

                //                if (decimal.TryParse(Utility.GetAttribute(elmScore, "補考成績"), out dsre))
                //                {
                //                    // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                //                    //dsre = Math.Round(dsre, 0, MidpointRounding.AwayFromZero);

                //                    decimal dsreP;

                //                    if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsreP))
                //                    {
                //                        passScore = dsreP;
                //                    }

                //                    // ssr.ReScore = string.Format("{0:##0}", dsre);
                //                    ssr.ReScore = dsre.ToString();

                //                    if (dsre < passScore)
                //                        ssr.ReScoreP = "0";
                //                    else
                //                        ssr.ReScoreP = "1";
                //                }

                //                ssr.isScScore = false;
                //                ssr.ReAScoreType = "3";  // 重修方式預設 3，專班辦理。

                //                if (ssr.ScoreP == "-1")
                //                {
                //                    if (Utility.GetAttribute(elmScore, "是否取得學分") != null)
                //                    {
                //                        if (Utility.GetAttribute(elmScore, "是否取得學分") == "是")
                //                        {
                //                            ssr.ScoreP = "1";
                //                        }
                //                        else
                //                        {
                //                            ssr.ScoreP = "0";
                //                        }
                //                    }
                //                }

                //                if (ssr.ReAScoreP == "-1")
                //                {
                //                    if (Utility.GetAttribute(elmScore, "是否取得學分") != null)
                //                    {
                //                        if (Utility.GetAttribute(elmScore, "是否取得學分") == "是")
                //                        {
                //                            ssr.ReAScoreP = "1";
                //                        }
                //                        else
                //                        {
                //                            ssr.ReAScoreP = "0";
                //                        }
                //                    }
                //                }

                //                // 有重修
                //                int sy, ss;
                //                if (int.TryParse(Utility.GetAttribute(elmScore, "重修學年度"), out sy) && int.TryParse(Utility.GetAttribute(elmScore, "重修學期"), out ss))
                //                {
                //                    if (sy == _SchoolYear && ss == _Semester)
                //                    {
                //                        SubjectScoreRec108List1.Add(ssr);
                //                    }
                //                }

                //            }

                //        }
                //    }

                //    #endregion
                //}


                //if (StudentSubjectScore2Dict.ContainsKey(studRec.StudentID))
                //{
                //    // 重讀
                //    foreach (string key in StudentSubjectScore2Dict[studRec.StudentID].Keys)
                //    {
                //        XElement elmRoot = StudentSubjectScore2Dict[studRec.StudentID][key].ScoreXML;
                //        foreach (XElement elmScore in elmRoot.Elements("Subject"))
                //        {
                //            SubjectScoreRec108 ssr = new SubjectScoreRec108();
                //            ssr.IDNumber = IDNumber;
                //            ssr.StudentID = studRec.StudentID;
                //            ssr.Birthday = BirthDate;
                //            ssr.GradeYear = StudentSubjectScore2Dict[studRec.StudentID][key].GradeYear;
                //            ssr.SchoolYear = StudentSubjectScore2Dict[studRec.StudentID][key].SchoolYear;
                //            ssr.Semester = StudentSubjectScore2Dict[studRec.StudentID][key].Semester;
                //            ssr.SubjectName = Utility.GetAttribute(elmScore, "科目");

                //            ssr.SubjectLevel = Utility.GetAttribute(elmScore, "科目級別");

                //            ssr.Name = studRec.StudentName;
                //            ssr.ClassName = studRec.RefClass.ClassName;
                //            ssr.SeatNo = studRec.SeatNo;
                //            ssr.StudentNumber = studRec.StudentNumber;

                //            foreach (SHSemesterHistoryRecord rec in SemsH)
                //            {
                //                foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                //                {
                //                    if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                //                    {
                //                        if (studRec.StudentID == item.RefStudentID)
                //                        {
                //                            ssr.HisClassName = item.ClassName;
                //                            ssr.HisSeatNo = item.SeatNo;
                //                            ssr.HisStudentNumber = item.StudentNumber;
                //                        }
                //                    }
                //                }
                //            }

                //            ssr.CourseCode = Utility.GetAttribute(elmScore, "修課科目代碼");

                //            //當課程類別為8(團體活動時間)及9(彈性活動時間)，且科目屬性不為D(充實(增廣)、補強性教學 [全學期、授予學分])時，不允許提交成績。
                //            //課程代碼為23碼
                //            ssr.CodePass = true;
                //            int startIndex1 = 16;
                //            int endIndex = 1;
                //            int startIndex2 = 18;

                //            if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                //            {
                //                if (ssr.CourseCode.Length > 22)
                //                {
                //                    string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                //                    string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                //                    if ((sub1 == "8" || sub1 == "9") && sub2 != "D")
                //                    {
                //                        ssr.CodePass = false;   //不可提交
                //                    }
                //                }
                //            }

                //            ssr.Credit = Utility.GetAttribute(elmScore, "開課學分數");
                //            // 預設值 -1
                //            ssr.Score = ssr.ScScore = ssr.ReScore = ssr.useCredit = ssr.ScoreP = ssr.ScScoreP = ssr.ReAScoreP = ssr.ReScoreP = "-1";

                //            string GrStr = "";
                //            if (StudGradYearDict.ContainsKey(studRec.StudentID))
                //                GrStr = StudGradYearDict[studRec.StudentID] + "_及";

                //            decimal ds, dsre, passScore = 60;

                //            if (decimal.TryParse(Utility.GetAttribute(elmScore, "原始成績"), out ds))
                //            {
                //                // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                //                //ds = Math.Round(ds, 0, MidpointRounding.AwayFromZero);

                //                decimal dsp;
                //                if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsp))
                //                    passScore = dsp;

                //                //ssr.Score = string.Format("{0:##0}", ds);
                //                ssr.Score = ds.ToString();

                //                //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                //                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                //                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                //                {
                //                    ssr.useCredit = "1";

                //                    if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                //                    {
                //                        if (ssr.CourseCode.Length > 22)
                //                        {
                //                            string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                //                            string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                //                            if (sub1 == "9" && sub2 == "D")
                //                            {
                //                                ssr.useCredit = "3";
                //                            }
                //                        }
                //                    }
                //                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                //                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                //                        ssr.useCredit = "3";
                //                }
                //                else
                //                {
                //                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                //                        ssr.useCredit = "2";
                //                }

                //                if (ds < passScore)
                //                    ssr.ScoreP = "0";
                //                else
                //                    ssr.ScoreP = "1";
                //            }
                //            else
                //            {
                //                //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                //                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                //                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                //                {
                //                    ssr.useCredit = "1";

                //                    if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                //                    {
                //                        if (ssr.CourseCode.Length > 22)
                //                        {
                //                            string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                //                            string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                //                            if (sub1 == "9" && sub2 == "D")
                //                            {
                //                                ssr.useCredit = "3";
                //                            }
                //                        }
                //                    }
                //                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                //                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                //                        ssr.useCredit = "3";
                //                }
                //                else
                //                {
                //                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                //                        ssr.useCredit = "2";
                //                }

                //                // 判斷是否 不需評分
                //                //if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                //                //{
                //                //    ssr.useCredit = "3";
                //                //}
                //            }

                //            if (decimal.TryParse(Utility.GetAttribute(elmScore, "補考成績"), out dsre))
                //            {
                //                // 四捨五入到整數位--2021年3月 取消處理四捨五入
                //                //dsre = Math.Round(dsre, 0, MidpointRounding.AwayFromZero);

                //                decimal dsreP;

                //                if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsreP))
                //                {
                //                    passScore = dsreP;
                //                }

                //                // ssr.ReScore = string.Format("{0:##0}", dsre);
                //                ssr.ReScore = dsre.ToString();

                //                if (dsre < passScore)
                //                    ssr.ReScoreP = "0";
                //                else
                //                    ssr.ReScoreP = "1";
                //            }

                //            ssr.isScScore = false;
                //            ssr.ReStudMark = "1";  // 重讀備註預設 1，免修。

                //            //if (Utility.GetAttribute(elmScore, "是否補修成績") != null)
                //            //{
                //            //    if (Utility.GetAttribute(elmScore, "是否補修成績") == "是")
                //            //    {
                //            //        ssr.isScScore = true;
                //            //    }
                //            //}

                //            //// 對應學生身分別
                //            //if (StudentHasUpdateCodeDict.ContainsKey(studRec.StudentID))
                //            //{
                //            //    if (UpdateCodeMappingDict.ContainsKey(StudentHasUpdateCodeDict[studRec.StudentID]))
                //            //    {
                //            //        ssr.StudType = UpdateCodeMappingDict[StudentHasUpdateCodeDict[studRec.StudentID]];
                //            //    }
                //            //}

                //            if (ssr.ScoreP == "-1")
                //            {
                //                if (Utility.GetAttribute(elmScore, "是否取得學分") != null)
                //                {
                //                    if (Utility.GetAttribute(elmScore, "是否取得學分") == "是")
                //                    {
                //                        ssr.ScoreP = "1";
                //                    }
                //                    else
                //                    {
                //                        ssr.ScoreP = "0";
                //                    }
                //                }
                //            }

                //            SubjectScoreRec108List2.Add(ssr);
                //        }
                //    }
                //}

            }


            // 驗證資料
            ValidateScores(SubjectScoreRec108List);
            ValidateScores(SubjectScoreRec108OtherList);
            //ValidateScores(SubjectScoreRec108ReScoreList);
            //ValidateScores(SubjectScoreRec108List1);
            //ValidateScores(SubjectScoreRec108List2);


            // 學生學年補考成績
            Dictionary<string, Dictionary<string, decimal>> StudentYearReScoreDict = new Dictionary<string, Dictionary<string, decimal>>();

            //進校學期成績要取學年
            StudentSubjectScoreDict.Clear();
            foreach (string className in ClassStudentDict.Keys)
            {
                List<string> studIDList = ClassStudentDict[className];
                if (studIDList.Count > 0)
                {
                    string query = string.Format(@"
           SELECT
               ref_student_id AS student_id,
               school_year,
               semester,
               grade_year,
               score_info
           FROM
               sems_subj_score
           WHERE
               school_year = {0}               
               AND ref_student_id IN({1});
           ", _SchoolYear, string.Join(",", studIDList.ToArray()));

                    DataTable dtSemsScore = qhStudSemsScore.Select(query);
                    foreach (DataRow dr in dtSemsScore.Rows)
                    {
                        string student_id = "" + dr["student_id"];

                        SubjectScoreXML ssx = new SubjectScoreXML();
                        ssx.StudentID = student_id;
                        ssx.SchoolYear = dr["school_year"] + "";
                        ssx.Semester = dr["semester"] + "";
                        ssx.GradeYear = dr["grade_year"] + "";

                        string key = ssx.SchoolYear + "_" + ssx.Semester;

                        XElement elm = null;
                        try
                        {
                            elm = XElement.Parse(dr["score_info"] + "");
                            ssx.ScoreXML = elm;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                        if (!StudentSubjectScoreDict.ContainsKey(student_id))
                            StudentSubjectScoreDict.Add(student_id, new Dictionary<string, SubjectScoreXML>());

                        if (!StudentSubjectScoreDict[student_id].ContainsKey(key))
                            StudentSubjectScoreDict[student_id].Add(key, ssx);

                    }
                }
            }



            foreach (SmartSchool.Customization.Data.StudentRecord rec in StudentRecList)
            {
                if (!StudGradYearDict.ContainsKey(rec.StudentID))
                    if (rec.RefClass != null)
                        StudGradYearDict.Add(rec.StudentID, rec.RefClass.GradeYear);

                string cla = "n";
                if (rec.RefClass != null)
                    cla = rec.RefClass.ClassName;

                if (!ClassStudentDict.ContainsKey(cla))
                    ClassStudentDict.Add(cla, new List<string>());

                ClassStudentDict[cla].Add(rec.StudentID);

            }

            // 取得學年科目成績
            StudentYearScoreDict = Utility.GetStudentYearScoreByStudentIDDict(_SchoolYear, StudentIDList);

            // 取得學年科目補考成績
            if (_Semester == 2)
            {
                StudentYearReScoreDict = Utility.GetStudentYearReScoreByStudentIDDict(_SchoolYear, StudentIDList);
            }
            else
            {
                StudentYearReScoreDict = new Dictionary<string, Dictionary<string, decimal>>();
            }


            if (StudentHasUpdateCodeDict.Count > 0)
            {
                QueryHelper qhStudSemsScoreOther = new QueryHelper();
                foreach (string className in ClassStudentDict.Keys)
                {
                    List<string> studIDList = ClassStudentDict[className];
                    if (studIDList.Count > 0)
                    {
                        string query = string.Format(@"
                        SELECT
                            ref_student_id AS student_id,
                            school_year,
                            semester,
                            grade_year,
                            score_info
                        FROM
                            sems_subj_score
                        WHERE
                            ref_student_id IN({0});
                        ", string.Join(",", StudentHasUpdateCodeDict.Keys.ToArray()));

                        DataTable dtSemsScore = qhStudSemsScoreOther.Select(query);
                        foreach (DataRow dr in dtSemsScore.Rows)
                        {
                            string student_id = dr["student_id"] + "";

                            SubjectScoreXML ssx = new SubjectScoreXML();
                            ssx.StudentID = student_id;
                            ssx.SchoolYear = dr["school_year"] + "";
                            ssx.Semester = dr["semester"] + "";
                            ssx.GradeYear = dr["grade_year"] + "";

                            int xx = int.Parse(ssx.SchoolYear) * 10 + int.Parse(ssx.Semester);

                            if (xx >= (_SchoolYear * 10 + _Semester))
                                continue;

                            string key = ssx.SchoolYear + "_" + ssx.Semester;

                            XElement elm = null;
                            try
                            {
                                elm = XElement.Parse(dr["score_info"].ToString());
                                ssx.ScoreXML = elm;
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                            if (!StudentSubjectScoreOtherDict.ContainsKey(student_id))
                                StudentSubjectScoreOtherDict.Add(student_id, new Dictionary<string, SubjectScoreXML>());

                            if (!StudentSubjectScoreOtherDict[student_id].ContainsKey(key))
                                StudentSubjectScoreOtherDict[student_id][key] = ssx;

                        }
                    }
                }
            }


            foreach (SmartSchool.Customization.Data.StudentRecord studRec in StudentRecList)
            {
                string IDNumber = studRec.IDNumber.ToUpper();
                string BirthDate = "";
                DateTime dt;
                if (DateTime.TryParse(studRec.Birthday, out dt))
                    BirthDate = Utility.ConvertChDateString(dt);

                if (StudentSubjectScoreDict.ContainsKey(studRec.StudentID))
                {
                    string smsKey = _SchoolYear + "_" + _Semester;

                    #region 一般

                    if (StudentSubjectScoreDict[studRec.StudentID].ContainsKey(smsKey))
                    {
                        XElement elmRoot = StudentSubjectScoreDict[studRec.StudentID][smsKey].ScoreXML;

                        foreach (XElement elmScore in elmRoot.Elements("Subject"))
                        {
                            SubjectScoreRec108 ssr = new SubjectScoreRec108();
                            ssr.IDNumber = IDNumber;
                            ssr.StudentID = studRec.StudentID;
                            ssr.Birthday = BirthDate;
                            ssr.GradeYear = StudentSubjectScoreDict[studRec.StudentID][smsKey].GradeYear;
                            ssr.SchoolYear = StudentSubjectScoreDict[studRec.StudentID][smsKey].SchoolYear;
                            ssr.Semester = StudentSubjectScoreDict[studRec.StudentID][smsKey].Semester;
                            ssr.SubjectName = Utility.GetAttribute(elmScore, "科目");
                            ssr.SubjectLevel = Utility.GetAttribute(elmScore, "科目級別");
                            ssr.Name = studRec.StudentName;
                            //ssr.ClassName = studRec.RefClass.ClassName;
                            //ssr.SeatNo = studRec.SeatNo;
                            //ssr.StudentNumber = studRec.StudentNumber;

                            foreach (SHSemesterHistoryRecord rec in SemsH)
                            {
                                foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                                {
                                    if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                                    {
                                        if (studRec.StudentID == item.RefStudentID)
                                        {
                                            ssr.HisClassName = item.ClassName;
                                            ssr.HisSeatNo = item.SeatNo;
                                            ssr.HisStudentNumber = item.StudentNumber;
                                        }
                                    }
                                }
                            }

                            ssr.CourseCode = Utility.GetAttribute(elmScore, "修課科目代碼");

                            //當課程類別為8(團體活動時間)及9(彈性活動時間)，且科目屬性不為D(充實(增廣)、補強性教學 [全學期、授予學分])時，不允許提交成績。
                            //課程代碼為23碼
                            ssr.CodePass = true;
                            int startIndex1 = 16;
                            int endIndex = 1;
                            int startIndex2 = 18;

                            if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                            {
                                if (ssr.CourseCode.Length > 22)
                                {
                                    string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                                    string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                                    if ((sub1 == "8" || sub1 == "9") && sub2 != "D")
                                    {
                                        ssr.CodePass = false;   //不可提交
                                    }
                                }
                            }



                            ssr.Credit = Utility.GetAttribute(elmScore, "開課學分數");
                            // 預設值 -1
                            ssr.Score = ssr.ScScore = ssr.ReScore = ssr.useCredit = ssr.ScoreP = ssr.ScScoreP = ssr.ReAScoreP = ssr.ReScoreP = "-1";

                            string GrStr = "";
                            if (StudGradYearDict.ContainsKey(studRec.StudentID))
                                GrStr = StudGradYearDict[studRec.StudentID] + "_及";

                            decimal ds, dsre, passScore = 60;

                            if (decimal.TryParse(Utility.GetAttribute(elmScore, "原始成績"), out ds))
                            {
                                // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                                //ds = Math.Round(ds, 0, MidpointRounding.AwayFromZero);

                                decimal dsp;
                                if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsp))
                                    passScore = dsp;

                                //ssr.Score = string.Format("{0:##0}", ds);
                                ssr.Score = ds.ToString();


                                ssr.ScoreP = "-1";

                                if (ds < passScore)
                                    ssr.ScoreP = "0";
                                else
                                    ssr.ScoreP = "1";
                            }

                            ssr.useCredit = "1";
                            if (Utility.GetAttribute(elmScore, "抵免") == "否")
                                ssr.useCredit = "2";

                            if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                            {
                                if (ssr.CourseCode.Length > 22)
                                {
                                    string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                                    string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                                    if (sub1 == "9" && sub2 == "D")
                                    {
                                        ssr.useCredit = "3";
                                    }
                                }
                            }

                            // 比對學年成績
                            ssr.YearScoreP = ssr.YearScore = "-1";
                            if (StudentYearScoreDict.ContainsKey(studRec.StudentID))
                            {
                                if (StudentYearScoreDict[studRec.StudentID].ContainsKey(ssr.SubjectName))
                                {
                                    decimal ys = StudentYearScoreDict[studRec.StudentID][ssr.SubjectName];
                                    ssr.YearScore = string.Format("{0:##0.0}", ys);
                                    if (ys < passScore)
                                        ssr.YearScoreP = "0";
                                    else
                                        ssr.YearScoreP = "1";
                                }
                            }

                            if (decimal.TryParse(Utility.GetAttribute(elmScore, "補考成績"), out dsre))
                            {
                                // 四捨五入到整數位--2021年3月 取消處理四捨五入
                                //dsre = Math.Round(dsre, 0, MidpointRounding.AwayFromZero);

                                decimal dsreP;

                                if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsreP))
                                {
                                    passScore = dsreP;
                                }

                                //ssr.ReScore = string.Format("{0:##0}", dsre);
                                ssr.ReScore = dsre.ToString();

                                ssr.ReScoreP = "-1";
                                if (dsre < passScore)
                                    ssr.ReScoreP = "0";
                                else
                                    ssr.ReScoreP = "1";
                            }

                            if (ssr.ScoreP == "-1")
                            {
                                if (Utility.GetAttribute(elmScore, "是否取得學分") != null)
                                {
                                    if (Utility.GetAttribute(elmScore, "是否取得學分") == "是")
                                    {
                                        ssr.ScoreP = "1";
                                    }
                                    else
                                    {
                                        ssr.ScoreP = "0";
                                    }
                                }
                            }
                            // 不計學分 = 是，不列入學習成績
                            if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                continue;

                            SubjectScoreRec108ListN.Add(ssr);
                        }
                    }

                    #endregion
                }


                if (StudentSubjectScoreOtherDict.ContainsKey(studRec.StudentID))
                {
                    // 轉學轉科
                    foreach (string key in StudentSubjectScoreOtherDict[studRec.StudentID].Keys)
                    {
                        XElement elmRoot = StudentSubjectScoreOtherDict[studRec.StudentID][key].ScoreXML;
                        foreach (XElement elmScore in elmRoot.Elements("Subject"))
                        {
                            SubjectScoreRec108 ssr = new SubjectScoreRec108();
                            ssr.IDNumber = IDNumber;
                            ssr.StudentID = studRec.StudentID;
                            ssr.Birthday = BirthDate;
                            ssr.GradeYear = StudentSubjectScoreOtherDict[studRec.StudentID][key].GradeYear;
                            ssr.SchoolYear = StudentSubjectScoreOtherDict[studRec.StudentID][key].SchoolYear;
                            ssr.Semester = StudentSubjectScoreOtherDict[studRec.StudentID][key].Semester;
                            ssr.SubjectName = Utility.GetAttribute(elmScore, "科目");
                            ssr.SubjectLevel = Utility.GetAttribute(elmScore, "科目級別");

                            ssr.Name = studRec.StudentName;
                            //ssr.ClassName = studRec.RefClass.ClassName;
                            //ssr.SeatNo = studRec.SeatNo;
                            //ssr.StudentNumber = studRec.StudentNumber;

                            foreach (SHSemesterHistoryRecord rec in SemsH)
                            {
                                foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                                {
                                    if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                                    {
                                        if (studRec.StudentID == item.RefStudentID)
                                        {
                                            ssr.HisClassName = item.ClassName;
                                            ssr.HisSeatNo = item.SeatNo;
                                            ssr.HisStudentNumber = item.StudentNumber;
                                        }
                                    }
                                }
                            }
                            ssr.CourseCode = Utility.GetAttribute(elmScore, "修課科目代碼");

                            //當課程類別為8(團體活動時間)及9(彈性活動時間)，且科目屬性不為D(充實(增廣)、補強性教學 [全學期、授予學分])時，不允許提交成績。
                            //課程代碼為23碼
                            ssr.CodePass = true;
                            int startIndex1 = 16;
                            int endIndex = 1;
                            int startIndex2 = 18;

                            if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                            {
                                if (ssr.CourseCode.Length > 22)
                                {
                                    string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                                    string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                                    if ((sub1 == "8" || sub1 == "9") && sub2 != "D")
                                    {
                                        ssr.CodePass = false;   //不可提交
                                    }
                                }
                            }

                            ssr.Credit = Utility.GetAttribute(elmScore, "開課學分數");
                            // 預設值 -1
                            ssr.Score = ssr.ScScore = ssr.ReScore = ssr.useCredit = ssr.ScoreP = ssr.ScScoreP = ssr.ReAScoreP = ssr.ReScoreP = "-1";

                            string GrStr = "";
                            if (StudGradYearDict.ContainsKey(studRec.StudentID))
                                GrStr = StudGradYearDict[studRec.StudentID] + "_及";

                            decimal ds, dsre, passScore = 60;

                            if (decimal.TryParse(Utility.GetAttribute(elmScore, "原始成績"), out ds))
                            {
                                // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                                //ds = Math.Round(ds, 0, MidpointRounding.AwayFromZero);

                                decimal dsp;
                                if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsp))
                                    passScore = dsp;

                                //ssr.Score = string.Format("{0:##0}", ds);
                                ssr.Score = ds.ToString();


                                if (ds < passScore)
                                    ssr.ScoreP = "0";
                                else
                                    ssr.ScoreP = "1";
                            }

                            ssr.useCredit = "1";
                            if (Utility.GetAttribute(elmScore, "抵免") == "否")
                                ssr.useCredit = "2";

                            if (!string.IsNullOrWhiteSpace(ssr.CourseCode))
                            {
                                if (ssr.CourseCode.Length > 22)
                                {
                                    string sub1 = ssr.CourseCode.Substring(startIndex1, endIndex);
                                    string sub2 = ssr.CourseCode.Substring(startIndex2, endIndex);
                                    if (sub1 == "9" && sub2 == "D")
                                    {
                                        ssr.useCredit = "3";
                                    }
                                }
                            }

                            if (decimal.TryParse(Utility.GetAttribute(elmScore, "補考成績"), out dsre))
                            {
                                // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                                //dsre = Math.Round(dsre, 0, MidpointRounding.AwayFromZero);

                                decimal dsreP;

                                if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsreP))
                                {
                                    passScore = dsreP;
                                }

                                //ssr.ReScore = string.Format("{0:##0}", dsre);
                                ssr.ReScore = dsre.ToString();

                                if (dsre < passScore)
                                    ssr.ReScoreP = "0";
                                else
                                    ssr.ReScoreP = "1";
                            }

                            ssr.isScScore = false;
                            ssr.ScScoreType = "3";  // 補考方式預設 3，專班辦理。

                            if (Utility.GetAttribute(elmScore, "是否補修成績") != null)
                            {
                                if (Utility.GetAttribute(elmScore, "是否補修成績") == "是")
                                {
                                    ssr.isScScore = true;
                                }
                            }

                            // 對應學生身分別
                            if (StudentHasUpdateCodeDict.ContainsKey(studRec.StudentID))
                            {
                                if (UpdateCodeMappingDict.ContainsKey(StudentHasUpdateCodeDict[studRec.StudentID]))
                                {
                                    ssr.StudType = UpdateCodeMappingDict[StudentHasUpdateCodeDict[studRec.StudentID]];
                                }
                            }

                            if (ssr.ScoreP == "-1")
                            {
                                if (Utility.GetAttribute(elmScore, "是否取得學分") != null)
                                {
                                    if (Utility.GetAttribute(elmScore, "是否取得學分") == "是")
                                    {
                                        ssr.ScoreP = "1";
                                    }
                                    else
                                    {
                                        ssr.ScoreP = "0";
                                    }
                                }
                            }

                            // 不計學分 = 是，不列入學習成績
                            if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                continue;

                            // 如果有勾是否補修成績，那補修學年度、學期和畫面上「不同」才會填入，相同的話會在補修成績工作頁 
                            // 2021-11 
                            // https://3.basecamp.com/4399967/buckets/15765350/todos/4305803619
                            if (ssr.isScScore)
                            {
                                if (!(Utility.GetAttribute(elmScore, "補修學年度") == _SchoolYear.ToString() && Utility.GetAttribute(elmScore, "補修學期") == _Semester.ToString()))
                                {
                                    SubjectScoreRec108OtherListN.Add(ssr);
                                }
                            }
                            else
                            {
                                //沒有勾是否補修成績，就直接填入
                                SubjectScoreRec108OtherListN.Add(ssr);
                            }
                            //SubjectScoreRec108OtherListN.Add(ssr);
                        }
                    }
                }


                // 處理學年補考
                // 學生有學年學年補考成績
                if (_Semester == 2)
                {
                    if (StudentYearReScoreDict.ContainsKey(studRec.StudentID))
                    {
                        // 學年補考科目名稱
                        foreach (string subjYearName in StudentYearReScoreDict[studRec.StudentID].Keys)
                        {
                            string k1 = _SchoolYear + "_1";
                            string k2 = _SchoolYear + "_" + _Semester;

                            SubjectYearScoreRec108N sysr = new SubjectYearScoreRec108N();
                            sysr.SchoolYear = _SchoolYear + "";
                            sysr.Semester = _Semester + "";
                            sysr.IDNumber = IDNumber;
                            sysr.StudentID = studRec.StudentID;
                            sysr.Birthday = BirthDate;
                            sysr.GradeYear = StudentSubjectScoreDict[studRec.StudentID][k2].GradeYear;
                            sysr.SubjectName = subjYearName;

                            bool hasScore1 = false, hasScore2 = false;
                            sysr.Name = studRec.StudentName;
                            //sysr.ClassName = studRec.RefClass.ClassName;
                            //sysr.SeatNo = studRec.SeatNo;
                            //sysr.StudentNumber = studRec.StudentNumber;

                            foreach (SHSemesterHistoryRecord rec in SemsH)
                            {
                                foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                                {
                                    if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                                    {
                                        if (studRec.StudentID == item.RefStudentID)
                                        {
                                            sysr.HisClassName = item.ClassName;
                                            sysr.HisSeatNo = item.SeatNo;
                                            sysr.HisStudentNumber = item.StudentNumber;
                                        }
                                    }
                                }
                            }

                            // 學年科目補考成績
                            sysr.ReScore = string.Format("{0:##0}", Math.Round(StudentYearReScoreDict[studRec.StudentID][subjYearName], 0, MidpointRounding.AwayFromZero));

                            // 上學期
                            if (StudentSubjectScoreDict[studRec.StudentID].ContainsKey(k1))
                            {
                                XElement elmRoot = StudentSubjectScoreDict[studRec.StudentID][k1].ScoreXML;

                                foreach (XElement elmScore in elmRoot.Elements("Subject"))
                                {
                                    if (subjYearName == Utility.GetAttribute(elmScore, "科目"))
                                    {
                                        sysr.CourseCode = Utility.GetAttribute(elmScore, "修課科目代碼");
                                        sysr.CodePass = true;
                                        int startIndex1 = 16;
                                        int endIndex = 1;
                                        int startIndex2 = 18;

                                        if (!string.IsNullOrWhiteSpace(sysr.CourseCode))
                                        {
                                            if (sysr.CourseCode.Length > 22)
                                            {
                                                string sub1 = sysr.CourseCode.Substring(startIndex1, endIndex);
                                                string sub2 = sysr.CourseCode.Substring(startIndex2, endIndex);
                                                if ((sub1 == "8" || sub1 == "9") && sub2 != "D")
                                                {
                                                    sysr.CodePass = false;   //不可提交
                                                }
                                            }
                                        }
                                        sysr.Credit1 = Utility.GetAttribute(elmScore, "開課學分數") == "" ? "-1" : Utility.GetAttribute(elmScore, "開課學分數");
                                        sysr.Pass1 = Utility.GetAttribute(elmScore, "是否取得學分");
                                        hasScore1 = true;
                                        // 不計學分 = 是，不列入學習成績
                                        if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                            continue;
                                    }
                                }
                            }

                            // 下學期
                            if (StudentSubjectScoreDict[studRec.StudentID].ContainsKey(k2))
                            {
                                XElement elmRoot = StudentSubjectScoreDict[studRec.StudentID][k2].ScoreXML;

                                foreach (XElement elmScore in elmRoot.Elements("Subject"))
                                {
                                    if (subjYearName == Utility.GetAttribute(elmScore, "科目"))
                                    {
                                        sysr.CourseCode = Utility.GetAttribute(elmScore, "修課科目代碼");
                                        sysr.CodePass = true;
                                        int startIndex1 = 16;
                                        int endIndex = 1;
                                        int startIndex2 = 18;

                                        if (!string.IsNullOrWhiteSpace(sysr.CourseCode))
                                        {
                                            if (sysr.CourseCode.Length > 22)
                                            {
                                                string sub1 = sysr.CourseCode.Substring(startIndex1, endIndex);
                                                string sub2 = sysr.CourseCode.Substring(startIndex2, endIndex);
                                                if ((sub1 == "8" || sub1 == "9") && sub2 != "D")
                                                {
                                                    sysr.CodePass = false;   //不可提交
                                                }
                                            }
                                        }
                                        sysr.Credit2 = Utility.GetAttribute(elmScore, "開課學分數") == "" ? "-1" : Utility.GetAttribute(elmScore, "開課學分數");
                                        sysr.Pass2 = Utility.GetAttribute(elmScore, "是否取得學分");
                                        hasScore2 = true;

                                        // 不計學分 = 是，不列入學習成績
                                        if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                            continue;
                                    }
                                }
                            }

                            sysr.ScoreP = "0";
                            if (hasScore1 == true && hasScore2 == true)
                            {
                                if (sysr.Pass1 == "是" && sysr.Pass2 == "是")
                                {
                                    sysr.ScoreP = "1";
                                }
                            }
                            else if (hasScore1 == true && hasScore2 == false)
                            {
                                if (sysr.Pass1 == "是")
                                {
                                    sysr.ScoreP = "1";
                                }
                            }
                            if (hasScore1 == false && hasScore2 == true)
                            {
                                if (sysr.Pass2 == "是")
                                {
                                    sysr.ScoreP = "1";
                                }
                            }

                            SubjectReScoreRec108ListN.Add(sysr);
                        }
                    }
                }

                // 處理補修學年度
                foreach (SubjectScoreRec108 ssr in SubjectScoreRec108ListN)
                {
                    // 有補修，需要回推應修學年度
                    if (ssr.isScScore)
                    {
                        ssr.SchoolYear = ssr.ScScoreSchoolYear;
                    }
                }

            }

            // 判斷資料身分證與課程代碼都有填寫
            foreach (SubjectScoreRec108 ssr in SubjectScoreRec108ListN)
            {
                ssr.checkPass = true;

                if (string.IsNullOrWhiteSpace(ssr.IDNumber) || string.IsNullOrWhiteSpace(ssr.CourseCode))
                    ssr.checkPass = false;
            }
            foreach (SubjectScoreRec108 ssr in SubjectScoreRec108OtherListN)
            {
                ssr.checkPass = true;

                if (string.IsNullOrWhiteSpace(ssr.IDNumber) || string.IsNullOrWhiteSpace(ssr.CourseCode))
                    ssr.checkPass = false;
            }

            foreach (SubjectYearScoreRec108N sys in SubjectReScoreRec108ListN)
            {
                sys.checkPass = true;
                if (string.IsNullOrWhiteSpace(sys.IDNumber) || string.IsNullOrWhiteSpace(sys.CourseCode))
                    sys.checkPass = false;
            }

            // 寫入學期成績
            _learningHistoryDataAccess.SaveScores42(SubjectScoreRec108List, _SchoolYear, _Semester);

            // 補修成績由計算學期科目成績時寫入，在這先註解。
            //// 寫入補修成績
            //_learningHistoryDataAccess.SaveScores43(SubjectScoreRec108ReScoreList, _SchoolYear, _Semester);
            //bgWorker.ReportProgress(80);

            // 寫入轉學/轉科成績
            _learningHistoryDataAccess.SaveScores44(SubjectScoreRec108OtherList, _SchoolYear, _Semester);

            // 重修成績由計算學期科目成績時寫入，在這先註解。
            // 寫入重修成績
            //_learningHistoryDataAccess.SaveScores52(SubjectScoreRec108List1, _SchoolYear, _Semester);


            //// 寫入重讀成績
            //_learningHistoryDataAccess.SaveScores53(SubjectScoreRec108List2, _SchoolYear, _Semester);
            //bgWorker.ReportProgress(90);

            _learningHistoryDataAccess.SaveScores62(SubjectScoreRec108ListN, _SchoolYear, _Semester);
            _learningHistoryDataAccess.SaveScores63(SubjectReScoreRec108ListN, _SchoolYear, _Semester);
            _learningHistoryDataAccess.SaveScores64(SubjectScoreRec108OtherListN, _SchoolYear, _Semester);

            //FISCA.LogAgent.ApplicationLog.Log("補修成績比對", "處理完成", $"學生數:{dataValue43.Count}, 學期成績學生數:{semsScore43.Count}");

            // 將比對結果回寫資料庫
            List<SubjectScoreRec108> allReScoreData = new List<SubjectScoreRec108>();
            foreach (var studentData in dataValue43)
            {
                allReScoreData.AddRange(studentData.Value);
            }

            // 使用 SaveScores43 方法回寫資料庫
            if (allReScoreData.Count > 0)
            {
                try
                {
                    _learningHistoryDataAccess.SaveScores43(allReScoreData, _SchoolYear, _Semester);                    
                }
                catch (Exception ex)
                {
                    FISCA.LogAgent.ApplicationLog.Log("補修成績回寫", "錯誤", $"回寫失敗:{ex.Message}");
                    
                }
            }
            else
            {
                //FISCA.LogAgent.ApplicationLog.Log("補修成績回寫", "警告", "沒有資料需要回寫");
            }

            // 將比對結果回寫資料庫
            List<SubjectScoreRec108> allReScoreData53 = new List<SubjectScoreRec108>();
            foreach (var studentData in dataValue53)
            {
                allReScoreData53.AddRange(studentData.Value);
            }

            // 使用 SaveScores53 方法回寫資料庫
            if (allReScoreData53.Count > 0)
            {
                try
                {
                    _learningHistoryDataAccess.SaveScores53(allReScoreData53, _SchoolYear, _Semester);
                }
                catch (Exception ex)
                {
                    FISCA.LogAgent.ApplicationLog.Log("重讀成績回寫", "錯誤", $"回寫失敗:{ex.Message}");

                }
            }
            else
            {
                //FISCA.LogAgent.ApplicationLog.Log("補修成績回寫", "警告", "沒有資料需要回寫");
            }


        }

        private void ValidateScores(List<SubjectScoreRec108> scores)
        {
            foreach (var ssr in scores)
            {
                ssr.checkPass = !string.IsNullOrWhiteSpace(ssr.IDNumber) && !string.IsNullOrWhiteSpace(ssr.CourseCode);
            }
        }

        /// <summary>
        /// 從學習歷程資料表取得學生補修資料
        /// </summary>
        /// <param name="schoolYear">學年度</param>
        /// <param name="semester">學期</param>
        /// <param name="studentIDList">學生ID列表</param>
        /// <param name="serialNo">學習歷程工作表編號，預設為"4.3"</param>
        /// <returns>補修資料列表</returns>
        public List<SubjectScoreRec108> GetLearningHistoryReScoreData43(int schoolYear, int semester, List<string> studentIDList, string serialNo = "4.3")
        {
            List<SubjectScoreRec108> reScoreList = new List<SubjectScoreRec108>();
            
            try
            {
                QueryHelper qhLearningHistoryReScore = new QueryHelper();
                
                // 建立欄位名稱列表
                List<string> jsonDataList = new List<string>
                {
                    "身分證號", "出生日期", "應修課學年度", "應修課學期", "課程代碼", "開課年級", "修課學分",
                    "補修成績", "補修及格", "補考成績", "補考及格", "補修方式", "是否採計學分", "質性文字描述",
                    "備註(學生姓名)", "備註(資料當學期班級)", "備註(資料當學期座號)", "備註(資料當學期學號)",
                    "備註(資料次學期班級)", "備註(資料次學期座號)", "備註(資料次學期學號)"
                };

                // 建立欄位名稱當 key 的 SQL 片段
                List<string> tmpList = new List<string>();
                foreach (string key in jsonDataList)
                {
                    tmpList.Add(string.Format("MAX(CASE WHEN detail->>'name' = '{0}' THEN detail->>'value' END) AS \"{0}\"", key));
                }

                string queryLearningHistoryReScore = string.Format(@"
                WITH student_score AS (
                    SELECT
                        ref_student_id AS student_id,
                        serial_no,
                        name,
                        school_year,
                        semester,
                        subject,
                        subj_level,
                        jsonb_array_elements(detail) AS detail
                    FROM
                        student_learning_history
                    WHERE              
                        serial_no = '{0}'
                        AND name = '補修成績'
                        AND school_year = {1}
                        AND semester = {2}
                        AND ref_student_id IN ({3})
                ),
                score_pivot AS (
                    SELECT
                        student_id,
                        serial_no,
                        name,
                        school_year,
                        semester,
                        subject,
                        subj_level,
                        {4}     
                    FROM
                        student_score
                    GROUP BY
                        student_id,
                        serial_no,
                        name,
                        school_year,
                        semester,
                        subject,
                        subj_level
                )
                SELECT * FROM score_pivot
                ORDER BY 
                    身分證號,
                    school_year,
                    semester,
                    subject;", 
                    serialNo, schoolYear, semester, string.Join(",", studentIDList.Select(id => "'" + id + "'")), string.Join(",", tmpList.ToArray()));

                DataTable dtLearningHistoryReScore = qhLearningHistoryReScore.Select(queryLearningHistoryReScore);
                
           
                // 用於暫存每個學生的補修記錄
                Dictionary<string, SubjectScoreRec108> studentReScoreDict = new Dictionary<string, SubjectScoreRec108>();

                // 處理查詢結果
                foreach (DataRow dr in dtLearningHistoryReScore.Rows)
                {
                    string student_id = dr["student_id"] + "";
                    string school_year = dr["school_year"] + "";
                    string semester_value = dr["semester"] + "";
                    string subject = dr["subject"] + "";
                    string subject_level = dr["subj_level"] + "";

                    // 如果學生不存在於字典中，則新增
                    if (!studentReScoreDict.ContainsKey(student_id))
                    {
                        SubjectScoreRec108 newRecord = new SubjectScoreRec108();
                        newRecord.StudentID = student_id; // 設定學生ID
                        newRecord.IDNumber = "";
                        newRecord.Birthday = "";
                        newRecord.SchoolYear = school_year;
                        newRecord.Semester = semester_value;
                        newRecord.CourseCode = "";
                        newRecord.SubjectName = subject;
                        newRecord.SubjectLevel = subject_level;
                        newRecord.Credit = "";
                        newRecord.Score = "";
                        newRecord.ScoreP = "";
                        newRecord.ReScore = "";
                        newRecord.ReScoreP = "";
                        newRecord.ScScoreType = "";
                        newRecord.useCredit = "";
                        newRecord.Text = "";
                        newRecord.Name = "";
                        newRecord.HisClassName = "";
                        newRecord.HisSeatNo = 0;
                        newRecord.HisStudentNumber = "";
                        newRecord.ClassName = "";
                        newRecord.SeatNo = "";
                        newRecord.StudentNumber = "";
                        newRecord.isScScore = true; // 標記為補修成績
                        newRecord.checkPass = false;
                        newRecord.CodePass = true;
                        
                        studentReScoreDict.Add(student_id, newRecord);
                    }

                    SubjectScoreRec108 existingRecord = studentReScoreDict[student_id];

                    // 直接從查詢結果取得各欄位的值（使用中文欄位名稱）
                    existingRecord.IDNumber = dr["身分證號"] + "";
                    existingRecord.Birthday = dr["出生日期"] + "";
                    existingRecord.SchoolYear = dr["應修課學年度"] + "";
                    existingRecord.Semester = dr["應修課學期"] + "";
                    existingRecord.CourseCode = dr["課程代碼"] + "";
                    existingRecord.SubjectName = subject; // 使用從查詢中取得的 subject 欄位
                    existingRecord.GradeYear = dr["開課年級"] + "";
                    existingRecord.Credit = dr["修課學分"] + "";
                    existingRecord.Score = dr["補修成績"] + "";
                    existingRecord.ScoreP = dr["補修及格"] + "";
                    existingRecord.ReScore = dr["補考成績"] + "";
                    existingRecord.ReScoreP = dr["補考及格"] + "";
                    existingRecord.ScScoreType = dr["補修方式"] + "";
                    existingRecord.useCredit = dr["是否採計學分"] + "";
                    existingRecord.Text = dr["質性文字描述"] + "";
                    existingRecord.Name = dr["備註(學生姓名)"] + "";
                    existingRecord.HisClassName = dr["備註(資料當學期班級)"] + "";
                    
                    // 處理座號（需要轉換為整數）
                    string seatNoStr = dr["備註(資料當學期座號)"] + "";
                    if (int.TryParse(seatNoStr, out int seatNo))
                        existingRecord.HisSeatNo = seatNo;
                    
                    existingRecord.HisStudentNumber = dr["備註(資料當學期學號)"] + "";
                    existingRecord.ClassName = dr["備註(資料次學期班級)"] + "";
                    existingRecord.SeatNo = dr["備註(資料次學期座號)"] + "";
                    existingRecord.StudentNumber = dr["備註(資料次學期學號)"] + "";
                }

                // 將處理後的資料加入結果列表
                foreach (var reScoreRecord in studentReScoreDict.Values)
                {
                    // 驗證資料完整性
                    if (!string.IsNullOrWhiteSpace(reScoreRecord.IDNumber) )
                    {
                        reScoreRecord.checkPass = true;
                        reScoreList.Add(reScoreRecord);
                    }
                }

          
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return reScoreList;
        }       

        /// <summary>
        /// 取得學習歷程補修資料並建立以StudentID為key的Dictionary
        /// </summary>
        /// <param name="schoolYear">學年度</param>
        /// <param name="semester">學期</param>
        /// <param name="studentIDList">學生ID列表</param>
        /// <param name="serialNo">學習歷程工作表編號，預設為"4.3"</param>
        /// <returns>包含學生ID和補修資料的字典</returns>
        public Dictionary<string, List<SubjectScoreRec108>> GetLearningHistoryReScoreDataAsDictionary43(int schoolYear, int semester, List<string> studentIDList, string serialNo = "4.3")
        {
            Dictionary<string, List<SubjectScoreRec108>> result = new Dictionary<string, List<SubjectScoreRec108>>();
            
            try
            {
                // 使用 GetLearningHistoryReScoreData 取得補修資料
                List<SubjectScoreRec108> reScoreData = GetLearningHistoryReScoreData43(schoolYear, semester, studentIDList, serialNo);
                
                // 將資料按學生ID分組
                foreach (var reScoreRecord in reScoreData)
                {
                    // 使用 StudentID 作為 Dictionary 的 key
                    string studentID = reScoreRecord.StudentID;
                    
                    if (!result.ContainsKey(studentID))
                    {
                        result[studentID] = new List<SubjectScoreRec108>();
                    }
                    
                    result[studentID].Add(reScoreRecord);
                }

                //FISCA.LogAgent.ApplicationLog.Log("學習歷程補修資料處理", "處理完成", $"學年度:{schoolYear}, 學期:{semester}, 學生數:{result.Count}");
            }
            catch (Exception ex)
            {
                //FISCA.LogAgent.ApplicationLog.Log("學習歷程補修資料處理", "錯誤", ex.Message);
                //throw;
            }

            return result;
        }

        /// <summary>
        /// 取得學生成績資料並建立以StudentID為key的Dictionary（包含補考成績）
        /// </summary>
        /// <param name="studentIDList">學生ID列表</param>
        /// <returns>包含學生ID和成績資料的字典</returns>
        public Dictionary<string, List<SubjectScoreRec108>> GetStudentScoreDataAsDictionary(List<string> studentIDList)
        {
            Dictionary<string, List<SubjectScoreRec108>> result = new Dictionary<string, List<SubjectScoreRec108>>();
            
            try
            {
                QueryHelper qh = new QueryHelper();
                string query = string.Format(@"
                SELECT
                    sems_subj_score_ext.ref_student_id,
                    sems_subj_score_ext.grade_year,
                    sems_subj_score_ext.semester,
                    sems_subj_score_ext.school_year,
                    array_to_string(xpath('//Subject/@科目', subj_score_ele), '')::text AS 科目,
                    array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '')::text AS 科目級別,
                    array_to_string(xpath('//Subject/@原始成績', subj_score_ele), '')::text AS 原始成績,
                    array_to_string(xpath('//Subject/@補考成績', subj_score_ele), '')::text AS 補考成績,
                    array_to_string(xpath('//Subject/@修課及格標準', subj_score_ele), '')::text AS 修課及格標準
                FROM (
                    SELECT 
                        sems_subj_score.*,
                        unnest(xpath('//SemesterSubjectScoreInfo/Subject', xmlparse(content score_info))) as subj_score_ele
                    FROM 
                        sems_subj_score 
                    WHERE ref_student_id IN ({0})
                ) as sems_subj_score_ext
                WHERE array_to_string(xpath('//Subject/@補考成績', subj_score_ele), '')::text IS NOT NULL 
                    AND array_to_string(xpath('//Subject/@補考成績', subj_score_ele), '')::text != ''
                ORDER BY grade_year desc, semester desc, school_year desc", 
                    string.Join(",", studentIDList));

                DataTable dt = qh.Select(query);
                
                foreach (DataRow dr in dt.Rows)
                {
                    string student_id = dr["ref_student_id"] + "";
                    
                    // 如果學生不存在於字典中，則新增
                    if (!result.ContainsKey(student_id))
                    {
                        result[student_id] = new List<SubjectScoreRec108>();
                    }

                    SubjectScoreRec108 scoreRecord = new SubjectScoreRec108();
                    
                    // 設定基本資料
                    scoreRecord.StudentID = student_id;
                    scoreRecord.GradeYear = dr["grade_year"] + "";
                    scoreRecord.Semester = dr["semester"] + "";
                    scoreRecord.SchoolYear = dr["school_year"] + "";
                    scoreRecord.SubjectName = dr["科目"] + "";
                    scoreRecord.SubjectLevel = dr["科目級別"] + "";
                    scoreRecord.Score = dr["原始成績"] + "";
                    scoreRecord.ReScore = dr["補考成績"] + "";
                    scoreRecord.ScoreP = dr["修課及格標準"] + "";
                    
                    // 設定其他必要欄位
                    scoreRecord.IDNumber = "";
                    scoreRecord.Birthday = "";
                    scoreRecord.CourseCode = "";
                    scoreRecord.Credit = "";
                    scoreRecord.ReScoreP = "";
                    scoreRecord.useCredit = "";
                    scoreRecord.Text = "";
                    scoreRecord.Name = "";
                    scoreRecord.HisClassName = "";
                    scoreRecord.HisSeatNo = 0;
                    scoreRecord.HisStudentNumber = "";
                    scoreRecord.ClassName = "";
                    scoreRecord.SeatNo = "";
                    scoreRecord.StudentNumber = "";
                    scoreRecord.isScScore = false; // 標記為一般成績
                    scoreRecord.checkPass = false;
                    scoreRecord.CodePass = true;
                    
                    // 驗證資料完整性
                    if (!string.IsNullOrWhiteSpace(scoreRecord.StudentID) && 
                        !string.IsNullOrWhiteSpace(scoreRecord.SubjectName) &&
                        !string.IsNullOrWhiteSpace(scoreRecord.ReScore))
                    {
                        scoreRecord.checkPass = true;
                        result[student_id].Add(scoreRecord);
                    }
                }

                
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return result;
        }
    

        /// <summary>
        /// 取得學習歷程重讀(5.3)資料並建立以StudentID為key的Dictionary
        /// </summary>
        public Dictionary<string, List<SubjectScoreRec108>> GetLearningHistoryRetakeDataAsDictionary53(int schoolYear, int semester, List<string> studentIDList, string serialNo = "5.3")
        {
            Dictionary<string, List<SubjectScoreRec108>> result = new Dictionary<string, List<SubjectScoreRec108>>();
            try
            {
                QueryHelper qh = new QueryHelper();
                List<string> jsonDataList = new List<string>
                {
                    "身分證號", "出生日期", "課程代碼", "開課年級", "修課學分", "再次修習成績", "再次修習成績及格", "補考成績", "補考及格", "重讀成績", "成績及格", "重讀註記", "是否採計學分", "質性文字描述", "備註(學生姓名)", "備註(資料當學期班級)", "備註(資料當學期座號)", "備註(資料當學期學號)", "備註(資料次學期班級)", "備註(資料次學期座號)", "備註(資料次學期學號)"
                };
                List<string> tmpList = new List<string>();
                foreach (string key in jsonDataList)
                {
                    tmpList.Add(string.Format("MAX(CASE WHEN detail->>'name' = '{0}' THEN detail->>'value' END) AS \"{0}\"", key));
                }
                string query = string.Format(@"
                WITH student_score AS (
                    SELECT
                        ref_student_id AS student_id,
                        serial_no,
                        name,
                        school_year,
                        semester,
                        subject,
                        subj_level,
                        jsonb_array_elements(detail) AS detail
                    FROM
                        student_learning_history
                    WHERE              
                        serial_no = '{0}'
                        AND name = '重讀成績'
                        AND school_year = {1}
                        AND semester = {2}
                        AND ref_student_id IN ({3})
                ),
                score_pivot AS (
                    SELECT
                        student_id,
                        serial_no,
                        name,
                        school_year,
                        semester,
                        subject,
                        subj_level,
                        {4}     
                    FROM
                        student_score
                    GROUP BY
                        student_id,
                        serial_no,
                        name,
                        school_year,
                        semester,
                        subject,
                        subj_level
                )
                SELECT * FROM score_pivot
                ORDER BY 
                    身分證號,
                    school_year,
                    semester,
                    subject;",
                    serialNo, schoolYear, semester, string.Join(",", studentIDList.Select(id => "'" + id + "'")), string.Join(",", tmpList.ToArray()));
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string student_id = dr["student_id"] + "";
                    string school_year = dr["school_year"] + "";
                    string semester_value = dr["semester"] + "";
                    string subject = dr["subject"] + "";
                    string subject_level = dr["subj_level"] + "";
                    var rec = new SubjectScoreRec108();
                    rec.StudentID = student_id;
                    rec.IDNumber = dr["身分證號"] + "";
                    rec.Birthday = dr["出生日期"] + "";
                    rec.SchoolYear = school_year;
                    rec.Semester = semester_value;
                    rec.CourseCode = dr["課程代碼"] + "";
                    rec.GradeYear = dr["開課年級"] + "";
                    rec.Credit = dr["修課學分"] + "";
                    rec.ReAScore = dr["再次修習成績"] + "";
                    rec.ReAScoreP = dr["再次修習成績及格"] + "";
                    rec.ReScore = dr["補考成績"] + "";
                    rec.ReScoreP = dr["補考及格"] + "";
                    rec.Score = dr["重讀成績"] + "";
                    rec.ScoreP = dr["成績及格"] + "";
                    rec.ReStudMark = dr["重讀註記"] + "";
                    rec.useCredit = dr["是否採計學分"] + "";
                    rec.Text = dr["質性文字描述"] + "";
                    rec.Name = dr["備註(學生姓名)"] + "";
                    rec.HisClassName = dr["備註(資料當學期班級)"] + "";
                    string seatNoStr = dr["備註(資料當學期座號)"] + "";
                    if (int.TryParse(seatNoStr, out int seatNo))
                        rec.HisSeatNo = seatNo;
                    rec.HisStudentNumber = dr["備註(資料當學期學號)"] + "";
                    rec.ClassName = dr["備註(資料次學期班級)"] + "";
                    rec.SeatNo = dr["備註(資料次學期座號)"] + "";
                    rec.StudentNumber = dr["備註(資料次學期學號)"] + "";
                    rec.SubjectName = subject;
                    rec.SubjectLevel = subject_level;
                    rec.isScScore = false; // 標記為重讀成績
                    rec.checkPass = !string.IsNullOrWhiteSpace(rec.IDNumber) && !string.IsNullOrWhiteSpace(rec.CourseCode);
                    rec.CodePass = true;
                    if (!result.ContainsKey(student_id))
                        result[student_id] = new List<SubjectScoreRec108>();
                    result[student_id].Add(rec);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return result;
        }
    }
}
