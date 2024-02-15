using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using FISCA.Data;
using System.Xml.Linq;

namespace StudentCourseScoreCheck100_SH.DAO
{
    public class QueryData
    {
        /// <summary>
        /// 學生評量名稱
        /// </summary>
        public static List<string> _StudentExamName = new List<string>();

        /// <summary>
        /// 取得班級年級
        /// </summary>
        /// <returns></returns>
        public static List<string> GetClassGradeYearList()
        {
            List<string> retVal = new List<string>();
            string query = "select distinct grade_year from class where grade_year is not null order by grade_year;";
            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(query);
            foreach (DataRow dr in dt.Rows)
                retVal.Add(dr["grade_year"].ToString());

            return retVal;
        }

        /// <summary>
        /// 透過學年度、學期、年級取得學生基本課程評量分數高過100或低於100學生資料
        /// </summary>
        /// <param name="gradeYearList"></param>
        /// <returns></returns>
        public static List<StudentCourseScoreBase> GetStudentClassBase(int SchoolYear, int Semester, List<string> gradeYearList)
        {
            _StudentExamName.Clear();

            List<StudentCourseScoreBase> retVal = new List<StudentCourseScoreBase>();
            Dictionary<string, StudentCourseScoreBase> SelectStudentDict = new Dictionary<string, StudentCourseScoreBase>();

            // 學生班級基本資料(學生狀態為一般)
            QueryHelper qh1 = new QueryHelper();
            string query1 = "select student.id,class.grade_year,student_number,class.class_name,seat_no,student.name from student inner join class on student.ref_class_id=class.id where student.status=1 and class.grade_year in(" + string.Join(",", gradeYearList.ToArray()) + ") order by class.grade_year,student_number;";
            DataTable dt1 = qh1.Select(query1);
            foreach (DataRow dr in dt1.Rows)
            {
                string sid = dr["id"].ToString();
                if (!SelectStudentDict.ContainsKey(sid))
                    SelectStudentDict.Add(sid, new StudentCourseScoreBase());

                SelectStudentDict[sid].StudentID = sid;
                SelectStudentDict[sid].GradeYear = dr["grade_year"].ToString();
                SelectStudentDict[sid].ClassName = dr["class_name"].ToString();
                SelectStudentDict[sid].Name = dr["name"].ToString();
                SelectStudentDict[sid].SeatNo = dr["seat_no"].ToString();
                SelectStudentDict[sid].StudentNumber = dr["student_number"].ToString();
            }

            // 取得所選的學生系統編號
            List<string> SelectStudentIDList = SelectStudentDict.Keys.ToList();

            // 課程授課教師資料
            QueryHelper qh2 = new QueryHelper();
            string query2 = "select ref_course_id,teacher.teacher_name from  tc_instruct inner join teacher on tc_instruct.ref_teacher_id=teacher.id inner join course on tc_instruct.ref_course_id=course.id where tc_instruct.sequence=1 and course.school_year=" + SchoolYear + " and course.semester=" + Semester;
            DataTable dt2 = qh2.Select(query2);
            Dictionary<string, string> courseTeacherDict = new Dictionary<string, string>();
            foreach (DataRow dr in dt2.Rows)
            {
                string id = dr["ref_course_id"].ToString();
                if (!courseTeacherDict.ContainsKey(id))
                    courseTeacherDict.Add(id, dr["teacher_name"].ToString());
            }

            // 透過學生系統編號、學年度、學期，取得課程成績小於0或大於100分的學生,過濾只有學生狀態為一般。
            QueryHelper qh3 = new QueryHelper();
            string query3 = "select course.course_name,sc_attend.score,ref_course_id,ref_student_id from sc_attend inner join course on sc_attend.ref_course_id=course.id inner join student on sc_attend.ref_student_id=student.id where student.status=1 and course.school_year=" + SchoolYear + " and semester=" + Semester + " and ref_student_id in(" + string.Join(",", SelectStudentIDList.ToArray()) + ") and (sc_attend.score>100 or sc_attend.score<0);";
            DataTable dt3 = qh3.Select(query3);
            foreach (DataRow dr in dt3.Rows)
            {
                string sid = dr["ref_student_id"].ToString();
                string cid = dr["ref_course_id"].ToString();
                if (SelectStudentDict.ContainsKey(sid))
                {
                    // 填入課程成績
                    if (!SelectStudentDict[sid].CousreScoreDict.ContainsKey(cid))
                    {
                        decimal score;
                        decimal.TryParse(dr["score"].ToString(), out score);
                        SelectStudentDict[sid].CousreScoreDict.Add(cid, score);
                    }

                    // 填入課程名稱
                    if (!SelectStudentDict[sid].CourseNameDict.ContainsKey(cid))
                        SelectStudentDict[sid].CourseNameDict.Add(cid, dr["course_name"].ToString());
                }
            }

            // 透過學生系統編號、學年度、學期，取得評量成績小於或大於100分的學生,過濾只有學生狀態為一般。
            QueryHelper qh4 = new QueryHelper();
            // string query4 = "select course.course_name,exam.exam_name,sce_take.score,ref_course_id,ref_student_id from sc_attend inner join course on sc_attend.ref_course_id=course.id  inner join sce_take on sce_take.ref_sc_attend_id=sc_attend.id inner join exam on sce_take.ref_exam_id=exam.id inner join student on sc_attend.ref_student_id=student.id where student.status=1 and course.school_year=" + SchoolYear + " and semester=" + Semester + " and ref_student_id in(" + string.Join(",", SelectStudentIDList.ToArray()) + ")  and (sce_take.score>100 or sce_take.score<0 );";

            string query4 = string.Format(@"
            SELECT
                course.course_name,
                exam.exam_name,
                sce_take.score,
	              sce_take.extension,
                ref_course_id,
                ref_student_id
            FROM
                sc_attend
                INNER JOIN course ON sc_attend.ref_course_id = course.id
                INNER JOIN sce_take ON sce_take.ref_sc_attend_id = sc_attend.id
                INNER JOIN exam ON sce_take.ref_exam_id = exam.id
                INNER JOIN student ON sc_attend.ref_student_id = student.id
            WHERE
                student.status = 1
                AND course.school_year = {0} 
                AND semester = {1} 
                AND ref_student_id IN({2})
                AND (
                    sce_take.score > 100
                    OR sce_take.score < 0
                );
            ", SchoolYear, Semester, string.Join(",", SelectStudentIDList.ToArray()));



            DataTable dt4 = qh4.Select(query4);
            foreach (DataRow dr in dt4.Rows)
            {
                string sid = dr["ref_student_id"].ToString();
                string cid = dr["ref_course_id"].ToString();
                string examName = dr["exam_name"].ToString();
                if (SelectStudentDict.ContainsKey(sid))
                {
                    // 放入評量成績
                    if (!SelectStudentDict[sid].ExamScoreDict.ContainsKey(cid))
                        SelectStudentDict[sid].ExamScoreDict.Add(cid, new Dictionary<string, decimal>());
                    if (!SelectStudentDict[sid].ExamScoreTextDict.ContainsKey(cid))
                        SelectStudentDict[sid].ExamScoreTextDict.Add(cid, new Dictionary<string, string>());


                    // 試別名稱整理
                    if (!_StudentExamName.Contains(examName))
                        _StudentExamName.Add(examName);

                    if (!SelectStudentDict[sid].ExamScoreDict[cid].ContainsKey(examName))
                    {
                        decimal score;
                        decimal.TryParse(dr["score"].ToString(), out score);
                        SelectStudentDict[sid].ExamScoreDict[cid].Add(examName, score);


                        if (score == -1 || score == -2)
                        {
                            string extension = dr["extension"] + "";
                            try
                            {
                                XElement elm = XElement.Parse(extension);
                                if (elm.Element("UseText") != null)
                                {
                                    SelectStudentDict[sid].ExamScoreTextDict[cid].Add(examName, elm.Element("UseText").Value);
                                }
                            }
                            catch (Exception ex)
                            {
                                Console.WriteLine(ex.Message);
                            }
                        }

                    }
                }

                // 填入課程名稱
                if (!SelectStudentDict[sid].CourseNameDict.ContainsKey(cid))
                    SelectStudentDict[sid].CourseNameDict.Add(cid, dr["course_name"].ToString());
            }

            // 填入課程教師
            foreach (StudentCourseScoreBase scr in SelectStudentDict.Values)
            {
                foreach (KeyValuePair<string, string> data in scr.CourseNameDict)
                {
                    if (courseTeacherDict.ContainsKey(data.Key))
                        if (!scr.CourseTeacherDict.ContainsKey(data.Key))
                            scr.CourseTeacherDict.Add(data.Key, courseTeacherDict[data.Key]);
                }
            }

            // 將成績有疑問回傳
            foreach (StudentCourseScoreBase scr in SelectStudentDict.Values)
            {
                bool add = false;

                if (scr.ExamScoreDict.Count > 0 || scr.CousreScoreDict.Count > 0)
                    add = true;

                if (add)
                    retVal.Add(scr);
            }
            return retVal;
        }


        // 取得缺考設定輸入文字與內容值
        public static Dictionary<string, string> GetExamUseTextReportValue()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            string query = string.Format(@"
            SELECT
                array_to_string(xpath('//UseText/text()', settings), '') AS UseText,
                array_to_string(xpath('//ScoreType/text()', settings), '') AS ScoreType,
                array_to_string(xpath('//ReportValue/text()', settings), '') AS ReportValue,
                array_to_string(xpath('//UseValue/text()', settings), '') AS UseValue
            FROM
                (
                    SELECT
                        unnest(
                            xpath(
                                '//Configurations/Configuration/Settings/Setting',
                                xmlparse(
                                    content REPLACE(REPLACE(content, '&lt;', '<'), '&gt;', '>')
                                )
                            )
                        ) AS settings
                    FROM
                        list
                    WHERE
                        name = '評量成績缺考設定'
                ) AS content
            ");

            DataTable dt = qh.Select(query);
            foreach (DataRow dr in dt.Rows)
            {
                string UseText = dr["usetext"] + "";
                if (!value.ContainsKey(UseText))
                    value.Add(UseText, dr["reportvalue"] + "");
            }

            return value;
        }
    }
}
