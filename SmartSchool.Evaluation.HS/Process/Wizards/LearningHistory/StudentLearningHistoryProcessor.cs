using FISCA.Data;
using SHSchool.Data;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Linq;
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
            bgWorker.ReportProgress(1);
            _SchoolYear = schoolYear;
            _Semester = semester;

            List<SubjectScoreRec108> SubjectScoreRec108List = new List<SubjectScoreRec108>();
            List<SubjectScoreRec108> SubjectScoreRec108OtherList = new List<SubjectScoreRec108>();
            List<SubjectScoreRec108> SubjectScoreRec108ReScoreList = new List<SubjectScoreRec108>();

            List<SubjectScoreRec108> SubjectScoreRec108List1 = new List<SubjectScoreRec108>();
            List<SubjectScoreRec108> SubjectScoreRec108List2 = new List<SubjectScoreRec108>();

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


            // 取得補修資料學生
            Dictionary<string, Dictionary<string, SubjectScoreXML>> StudentSubjectReScoreDict = new Dictionary<string, Dictionary<string, SubjectScoreXML>>();

            // 取得畫面上學年度學期補修資料
            QueryHelper qhStudSemsReScore = new QueryHelper();
            string queryReScore = string.Format(@"
            WITH sems_subj_score AS(
                SELECT
                    sems_subj_score_ext.ref_student_id AS student_id,
                    sems_subj_score_ext.grade_year,
                    sems_subj_score_ext.semester,
                    sems_subj_score_ext.school_year,
                    sems_subj_score_ext.score_info,
                    array_to_string(xpath('//Subject/@是否補修成績', subj_score_ele), '') :: text AS 是否補修成績,
                    array_to_string(xpath('//Subject/@補修學年度', subj_score_ele), '') :: text AS 補修學年度,
                    array_to_string(xpath('//Subject/@補修學期', subj_score_ele), '') :: text AS 補修學期,
                    array_to_string(xpath('//Subject/@修課科目代碼', subj_score_ele), '') :: text AS 課程代碼
                FROM
                    (
                        SELECT
                            sems_subj_score.*,
                            unnest(
                                xpath(
                                    '//SemesterSubjectScoreInfo/Subject',
                                    xmlparse(content score_info)
                                )
                            ) AS subj_score_ele
                        FROM
                            sems_subj_score
                        WHERE
                            ref_student_id IN ({0})
                    ) AS sems_subj_score_ext
            )
            SELECT
                *
            FROM
                sems_subj_score
            WHERE
                是否補修成績 = '是'
                AND 補修學年度 = '{1}'
                AND 補修學期 = '{2}';
            ", string.Join(",", StudentIDList.ToArray()), _SchoolYear, _Semester);

            DataTable dtReScore = qhStudSemsReScore.Select(queryReScore);
            foreach (DataRow dr in dtReScore.Rows)
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
                if (!StudentSubjectReScoreDict.ContainsKey(student_id))
                    StudentSubjectReScoreDict.Add(student_id, new Dictionary<string, SubjectScoreXML>());

                if (!StudentSubjectReScoreDict[student_id].ContainsKey(key))
                    StudentSubjectReScoreDict[student_id].Add(key, ssx);
            }

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

            bgWorker.ReportProgress(30);

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
                            ssr.ClassName = studRec.RefClass.ClassName;
                            ssr.SeatNo = studRec.SeatNo;
                            ssr.StudentNumber = studRec.StudentNumber;

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

                                //2021年4月

                                //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                {
                                    ssr.useCredit = "1";

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
                                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                        ssr.useCredit = "3";
                                }
                                else
                                {
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                        ssr.useCredit = "2";
                                }

                                if (ds < passScore)
                                    ssr.ScoreP = "0";
                                else
                                    ssr.ScoreP = "1";
                            }
                            else
                            {

                                // if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                {
                                    ssr.useCredit = "1";

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

                                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                        ssr.useCredit = "3";
                                }
                                else
                                {
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                        ssr.useCredit = "2";
                                }

                                //// 判斷是否 不需評分
                                //if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                //{
                                //    ssr.useCredit = "3";
                                //}
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
                            ssr.ClassName = studRec.RefClass.ClassName;
                            ssr.SeatNo = studRec.SeatNo;
                            ssr.StudentNumber = studRec.StudentNumber;

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

                                //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                {
                                    ssr.useCredit = "1";

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

                                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                        ssr.useCredit = "3";
                                }
                                else
                                {
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                        ssr.useCredit = "2";
                                }

                                if (ds < passScore)
                                    ssr.ScoreP = "0";
                                else
                                    ssr.ScoreP = "1";
                            }
                            else
                            {
                                // if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                {
                                    ssr.useCredit = "1";

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

                                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                        ssr.useCredit = "3";
                                }
                                else
                                {
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                        ssr.useCredit = "2";
                                }

                                //// 判斷是否 不需評分
                                //if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                //{
                                //    ssr.useCredit = "3";
                                //}
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


                // 補修
                if (StudentSubjectReScoreDict.ContainsKey(studRec.StudentID))
                {
                    #region 補修
                    foreach (string smsKey in StudentSubjectReScoreDict[studRec.StudentID].Keys)
                        if (StudentSubjectReScoreDict[studRec.StudentID].ContainsKey(smsKey))
                        {
                            XElement elmRoot = StudentSubjectReScoreDict[studRec.StudentID][smsKey].ScoreXML;

                            foreach (XElement elmScore in elmRoot.Elements("Subject"))
                            {

                                SubjectScoreRec108 ssr = new SubjectScoreRec108();
                                ssr.IDNumber = IDNumber.ToUpper();
                                ssr.StudentID = studRec.StudentID;
                                ssr.Birthday = BirthDate;
                                ssr.GradeYear = StudentSubjectReScoreDict[studRec.StudentID][smsKey].GradeYear;
                                ssr.SchoolYear = StudentSubjectReScoreDict[studRec.StudentID][smsKey].SchoolYear;
                                ssr.Semester = StudentSubjectReScoreDict[studRec.StudentID][smsKey].Semester;
                                ssr.SubjectName = Utility.GetAttribute(elmScore, "科目");

                                ssr.SubjectLevel = Utility.GetAttribute(elmScore, "科目級別");

                                ssr.Name = studRec.StudentName;
                                ssr.ClassName = studRec.RefClass.ClassName;
                                ssr.SeatNo = studRec.SeatNo;
                                ssr.StudentNumber = studRec.StudentNumber;

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

                                    //2021年4月

                                    //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                    // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                    {
                                        ssr.useCredit = "1";

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
                                        // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                        if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                            ssr.useCredit = "3";
                                    }
                                    else
                                    {
                                        if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                            ssr.useCredit = "2";
                                    }

                                    if (ds < passScore)
                                        ssr.ScoreP = "0";
                                    else
                                        ssr.ScoreP = "1";
                                }
                                else
                                {

                                    //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                    // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                    {
                                        ssr.useCredit = "1";

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

                                        // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                        if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                            ssr.useCredit = "3";
                                    }
                                    else
                                    {
                                        if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                            ssr.useCredit = "2";
                                    }

                                    //// 判斷是否 不需評分
                                    //if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                    //{
                                    //    ssr.useCredit = "3";
                                    //}
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


                                // 補修學年度、學期和畫面上選相同相同才會填入
                                if (Utility.GetAttribute(elmScore, "補修學年度") == _SchoolYear.ToString() && Utility.GetAttribute(elmScore, "補修學期") == _Semester.ToString())
                                {
                                    SubjectScoreRec108ReScoreList.Add(ssr);
                                }

                            }
                        }

                    #endregion
                }

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

            // 重修
            Dictionary<string, Dictionary<string, SubjectScoreXML>> StudentSubjectScore1Dict = new Dictionary<string, Dictionary<string, SubjectScoreXML>>();

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

            // 依年級分批取得學生學期科目成績 比對重修
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
                    ", string.Join(",", studIDList.ToArray()));

                    DataTable dtSemsScore = qhStudSemsScore.Select(query);
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
                        if (!StudentSubjectScore1Dict.ContainsKey(student_id))
                            StudentSubjectScore1Dict.Add(student_id, new Dictionary<string, SubjectScoreXML>());

                        if (!StudentSubjectScore1Dict[student_id].ContainsKey(key))
                            StudentSubjectScore1Dict[student_id][key] = ssx;

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

                if (StudentSubjectScore1Dict.ContainsKey(studRec.StudentID))
                {
                    // string smsKey = _SchoolYear + "_" + _Semester;

                    #region 重修

                    foreach (string smsKey in StudentSubjectScore1Dict[studRec.StudentID].Keys)
                    {
                        if (StudentSubjectScore1Dict[studRec.StudentID].ContainsKey(smsKey))
                        {
                            XElement elmRoot = StudentSubjectScore1Dict[studRec.StudentID][smsKey].ScoreXML;

                            foreach (XElement elmScore in elmRoot.Elements("Subject"))
                            {
                                SubjectScoreRec108 ssr = new SubjectScoreRec108();
                                ssr.IDNumber = IDNumber;
                                ssr.StudentID = studRec.StudentID;
                                ssr.Birthday = BirthDate;
                                ssr.GradeYear = StudentSubjectScore1Dict[studRec.StudentID][smsKey].GradeYear;
                                ssr.SchoolYear = StudentSubjectScore1Dict[studRec.StudentID][smsKey].SchoolYear;  //原修課學年度
                                ssr.Semester = StudentSubjectScore1Dict[studRec.StudentID][smsKey].Semester;  //原修課學期
                                ssr.SubjectName = Utility.GetAttribute(elmScore, "科目");
                                ssr.SubjectLevel = Utility.GetAttribute(elmScore, "科目級別");
                                ssr.Name = studRec.StudentName;
                                ssr.ClassName = studRec.RefClass.ClassName;
                                ssr.SeatNo = studRec.SeatNo;
                                ssr.StudentNumber = studRec.StudentNumber;


                                foreach (SHSemesterHistoryRecord rec in SemsH)
                                {
                                    foreach (K12.Data.SemesterHistoryItem item in rec.SemesterHistoryItems)
                                    {
                                        //if (item.SchoolYear == _SchoolYear && item.Semester == _Semester)
                                        //{
                                        //    if (studRec.StudentID == item.RefStudentID)
                                        //    {
                                        //        ssr.HisClassName = item.ClassName;
                                        //        ssr.HisSeatNo = item.SeatNo;
                                        //        ssr.HisStudentNumber = item.StudentNumber;
                                        //    }
                                        //}

                                        if (item.SchoolYear.ToString() == ssr.SchoolYear && item.Semester.ToString() == ssr.Semester)
                                        {
                                            if (studRec.StudentID == item.RefStudentID)
                                            {
                                                ssr.HisClassName = item.ClassName;
                                                ssr.HisSeatNo = item.SeatNo;
                                                ssr.HisStudentNumber = item.StudentNumber;


                                            }
                                        }

                                        /// 2021-11-05 因國教署系統判斷問題，故呈報重修時，課程代碼要填入原習修學年期在使用的課程代碼
                                        // 將每個學年期的學期對照表GDCCode存入Dictionary
                                        string key = item.RefStudentID + "_" + item.SchoolYear.ToString() + "_" + item.Semester.ToString();

                                        if (!StudHistoryGDCCodeDict.ContainsKey(key))
                                        {
                                            // 要傳入 StudHistoryGDCCodeDict 的 Dictionary
                                            Dictionary<string, string> paramForStudHistoryGDCCodeDict = new Dictionary<string, string>();
                                            if (!paramForStudHistoryGDCCodeDict.ContainsKey(item.RefStudentID))
                                            {
                                                paramForStudHistoryGDCCodeDict.Add(item.RefStudentID, item.GDCCode);
                                            }
                                            StudHistoryGDCCodeDict.Add(key, paramForStudHistoryGDCCodeDict);
                                        }
                                    }
                                }

                                string dicKey = studRec.StudentID + "_" + ssr.SchoolYear + "_" + ssr.Semester;
                                string subjectCode = "";

                                if (Utility.GetAttribute(elmScore, "重修學年度") == _SchoolYear.ToString() && Utility.GetAttribute(elmScore, "重修學期") == _Semester.ToString())
                                {
                                    /// 2021-11-05 如果有遇到復學/重讀/轉入，會將過去的課程代碼抵免成新的課程代碼，
                                    if (StudentCreditUpdateCodeDict.ContainsKey(studRec.StudentID))
                                    {
                                        // 最後一筆異動學年期
                                        int intUpdateYear_Sems = int.Parse(StudentCreditUpdateCodeDict[studRec.StudentID]);
                                        string strtUpdateYear_Sems = StudentCreditUpdateCodeDict[studRec.StudentID];
                                        //重修 原習修學年期
                                        int strSsrYear_Sems = int.Parse(ssr.SchoolYear + ssr.Semester);

                                        string updateYear = strtUpdateYear_Sems.Substring(0, 3);
                                        string updateSems = strtUpdateYear_Sems.Substring(3, 1);

                                        //https://3.basecamp.com/4399967/buckets/15765350/todos/4314350115#__recording_4323437970
                                        /// 所以 當 原習修學年期 <= 異動學年期，2022/09/21 且 異動學年期<=選擇的學年期，這裡的課程代碼已經抵免了，就填寫異動學年期的課程代碼
                                        if (strSsrYear_Sems <= intUpdateYear_Sems && intUpdateYear_Sems <= int.Parse(_SchoolYear.ToString() + _Semester.ToString()))
                                        {
                                            dicKey = studRec.StudentID + "_" + updateYear + "_" + updateSems;
                                        }
                                    }
                                    ssr.CourseCode = Utility.GetAttribute(elmScore, "修課科目代碼");



                                }


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

                                if (decimal.TryParse(Utility.GetAttribute(elmScore, "重修成績"), out ds))
                                {
                                    // 四捨五入到整數位 --2021年3月 取消處理四捨五入
                                    //ds = Math.Round(ds, 0, MidpointRounding.AwayFromZero);

                                    decimal dsp;
                                    if (decimal.TryParse(Utility.GetAttribute(elmScore, "修課及格標準"), out dsp))
                                        passScore = dsp;

                                    //ssr.ReAScore = string.Format("{0:##0}", ds);
                                    ssr.ReAScore = ds.ToString();

                                    // if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                    // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                    {
                                        ssr.useCredit = "1";

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
                                        // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                        if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                            ssr.useCredit = "3";
                                    }
                                    else
                                    {
                                        if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                            ssr.useCredit = "2";
                                    }

                                    if (ds < passScore)
                                        ssr.ReAScoreP = "0";
                                    else
                                    {
                                        ssr.ReAScoreP = "1";

                                        //2021-11-09 重修成績最高分為及格標準，若沒有設定及格標準則以60分計算。
                                        ssr.ReAScore = passScore.ToString();
                                    }

                                }
                                else
                                {
                                    //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                    // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                    {
                                        ssr.useCredit = "1";

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
                                        // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                        if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                            ssr.useCredit = "3";
                                    }
                                    else
                                    {
                                        if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                            ssr.useCredit = "2";
                                    }

                                    // 判斷是否 不需評分
                                    //if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                    //{
                                    //    ssr.useCredit = "3";
                                    //}
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

                                    // ssr.ReScore = string.Format("{0:##0}", dsre);
                                    ssr.ReScore = dsre.ToString();

                                    if (dsre < passScore)
                                        ssr.ReScoreP = "0";
                                    else
                                        ssr.ReScoreP = "1";
                                }

                                ssr.isScScore = false;
                                ssr.ReAScoreType = "3";  // 重修方式預設 3，專班辦理。

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

                                if (ssr.ReAScoreP == "-1")
                                {
                                    if (Utility.GetAttribute(elmScore, "是否取得學分") != null)
                                    {
                                        if (Utility.GetAttribute(elmScore, "是否取得學分") == "是")
                                        {
                                            ssr.ReAScoreP = "1";
                                        }
                                        else
                                        {
                                            ssr.ReAScoreP = "0";
                                        }
                                    }
                                }

                                // 有重修
                                int sy, ss;
                                if (int.TryParse(Utility.GetAttribute(elmScore, "重修學年度"), out sy) && int.TryParse(Utility.GetAttribute(elmScore, "重修學期"), out ss))
                                {
                                    if (sy == _SchoolYear && ss == _Semester)
                                    {
                                        SubjectScoreRec108List1.Add(ssr);
                                    }
                                }

                            }

                        }
                    }

                    #endregion
                }


                if (StudentSubjectScore2Dict.ContainsKey(studRec.StudentID))
                {
                    // 重讀
                    foreach (string key in StudentSubjectScore2Dict[studRec.StudentID].Keys)
                    {
                        XElement elmRoot = StudentSubjectScore2Dict[studRec.StudentID][key].ScoreXML;
                        foreach (XElement elmScore in elmRoot.Elements("Subject"))
                        {
                            SubjectScoreRec108 ssr = new SubjectScoreRec108();
                            ssr.IDNumber = IDNumber;
                            ssr.StudentID = studRec.StudentID;
                            ssr.Birthday = BirthDate;
                            ssr.GradeYear = StudentSubjectScore2Dict[studRec.StudentID][key].GradeYear;
                            ssr.SchoolYear = StudentSubjectScore2Dict[studRec.StudentID][key].SchoolYear;
                            ssr.Semester = StudentSubjectScore2Dict[studRec.StudentID][key].Semester;
                            ssr.SubjectName = Utility.GetAttribute(elmScore, "科目");

                            ssr.SubjectLevel = Utility.GetAttribute(elmScore, "科目級別");

                            ssr.Name = studRec.StudentName;
                            ssr.ClassName = studRec.RefClass.ClassName;
                            ssr.SeatNo = studRec.SeatNo;
                            ssr.StudentNumber = studRec.StudentNumber;

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

                                //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                {
                                    ssr.useCredit = "1";

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
                                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                        ssr.useCredit = "3";
                                }
                                else
                                {
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                        ssr.useCredit = "2";
                                }

                                if (ds < passScore)
                                    ssr.ScoreP = "0";
                                else
                                    ssr.ScoreP = "1";
                            }
                            else
                            {
                                //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                {
                                    ssr.useCredit = "1";

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
                                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                        ssr.useCredit = "3";
                                }
                                else
                                {
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                        ssr.useCredit = "2";
                                }

                                // 判斷是否 不需評分
                                //if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                //{
                                //    ssr.useCredit = "3";
                                //}
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

                                if (dsre < passScore)
                                    ssr.ReScoreP = "0";
                                else
                                    ssr.ReScoreP = "1";
                            }

                            ssr.isScScore = false;
                            ssr.ReStudMark = "1";  // 重讀備註預設 1，免修。

                            //if (Utility.GetAttribute(elmScore, "是否補修成績") != null)
                            //{
                            //    if (Utility.GetAttribute(elmScore, "是否補修成績") == "是")
                            //    {
                            //        ssr.isScScore = true;
                            //    }
                            //}

                            //// 對應學生身分別
                            //if (StudentHasUpdateCodeDict.ContainsKey(studRec.StudentID))
                            //{
                            //    if (UpdateCodeMappingDict.ContainsKey(StudentHasUpdateCodeDict[studRec.StudentID]))
                            //    {
                            //        ssr.StudType = UpdateCodeMappingDict[StudentHasUpdateCodeDict[studRec.StudentID]];
                            //    }
                            //}

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

                            SubjectScoreRec108List2.Add(ssr);
                        }
                    }
                }

            }


            // 驗證資料
            ValidateScores(SubjectScoreRec108List);
            ValidateScores(SubjectScoreRec108OtherList);
            ValidateScores(SubjectScoreRec108ReScoreList);
            ValidateScores(SubjectScoreRec108List1);
            ValidateScores(SubjectScoreRec108List2);


            bgWorker.ReportProgress(50);

            // 學生學年補考成績
            Dictionary<string, Dictionary<string, decimal>> StudentYearReScoreDict = new Dictionary<string, Dictionary<string, decimal>>();


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
                            ssr.ClassName = studRec.RefClass.ClassName;
                            ssr.SeatNo = studRec.SeatNo;
                            ssr.StudentNumber = studRec.StudentNumber;

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

                                //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                {
                                    ssr.useCredit = "1";

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
                                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                        ssr.useCredit = "3";

                                }
                                else
                                {
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                        ssr.useCredit = "2";
                                }

                                ssr.ScoreP = "-1";

                                if (ds < passScore)
                                    ssr.ScoreP = "0";
                                else
                                    ssr.ScoreP = "1";


                            }
                            else
                            {
                                //if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                {
                                    ssr.useCredit = "1";

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
                                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                        ssr.useCredit = "3";

                                }
                                else
                                {
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                        ssr.useCredit = "2";
                                }

                                // 判斷是否 不需評分
                                //if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                //{
                                //    ssr.useCredit = "3";
                                //}
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
                            ssr.ClassName = studRec.RefClass.ClassName;
                            ssr.SeatNo = studRec.SeatNo;
                            ssr.StudentNumber = studRec.StudentNumber;

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

                                // if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                {
                                    ssr.useCredit = "1";

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
                                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                        ssr.useCredit = "3";
                                }
                                else
                                {
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                        ssr.useCredit = "2";
                                }

                                if (ds < passScore)
                                    ssr.ScoreP = "0";
                                else
                                    ssr.ScoreP = "1";
                            }
                            else
                            {
                                // if (Utility.GetAttribute(elmScore, "不計學分") == "否" && Utility.GetAttribute(elmScore, "不需評分") == "否")
                                // 2023/10/3，因為判斷規則調整，不需評分不需要判斷。
                                if (Utility.GetAttribute(elmScore, "不計學分") == "否")
                                {
                                    ssr.useCredit = "1";

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
                                    // 2023/10/6，學校反應當需要計分又不需評分狀態時，是否採計學分需要填入3
                                    if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                        ssr.useCredit = "3";
                                }
                                else
                                {
                                    if (Utility.GetAttribute(elmScore, "不計學分") == "是")
                                        ssr.useCredit = "2";
                                }

                                // 判斷是否 不需評分
                                //if (Utility.GetAttribute(elmScore, "不需評分") == "是")
                                //{
                                //    ssr.useCredit = "3";
                                //}
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
                            sysr.ClassName = studRec.RefClass.ClassName;
                            sysr.SeatNo = studRec.SeatNo;
                            sysr.StudentNumber = studRec.StudentNumber;

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

            bgWorker.ReportProgress(70);

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
            _learningHistoryDataAccess.SaveScores42(SubjectScoreRec108List,_SchoolYear,_Semester);
            bgWorker.ReportProgress(75);

            // 寫入補修成績
            _learningHistoryDataAccess.SaveScores43(SubjectScoreRec108ReScoreList, _SchoolYear, _Semester);
            bgWorker.ReportProgress(80);

            // 寫入轉學/轉科成績
            _learningHistoryDataAccess.SaveScores44(SubjectScoreRec108OtherList, _SchoolYear, _Semester);
            bgWorker.ReportProgress(85);


            // 寫入重修成績
            _learningHistoryDataAccess.SaveScores52(SubjectScoreRec108List1, _SchoolYear, _Semester);
            // 寫入重讀成績
            _learningHistoryDataAccess.SaveScores53(SubjectScoreRec108List2, _SchoolYear, _Semester);
            bgWorker.ReportProgress(90);

            _learningHistoryDataAccess.SaveScores62(SubjectScoreRec108ListN, _SchoolYear, _Semester);
            _learningHistoryDataAccess.SaveScores63(SubjectReScoreRec108ListN, _SchoolYear, _Semester);
            _learningHistoryDataAccess.SaveScores64(SubjectScoreRec108OtherListN, _SchoolYear, _Semester);
            bgWorker.ReportProgress(100);
        }

        private void ValidateScores(List<SubjectScoreRec108> scores)
        {
            foreach (var ssr in scores)
            {
                ssr.checkPass = !string.IsNullOrWhiteSpace(ssr.IDNumber) && !string.IsNullOrWhiteSpace(ssr.CourseCode);
            }
        }
    }
}
