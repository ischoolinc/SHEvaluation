using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;
using FISCA.Data;
using SHCourseScoreInputStatus.DAO;

namespace SHCourseScoreInputStatus.DAO
{
    public class QueryData
    {
        /// <summary>
        /// 透過學年度學期，取得課程成績基本資料
        /// </summary>
        /// <param name="CourseIDList"></param>
        /// <returns></returns>
        public static List<CourseScoreBase> GetCourseScoreBaseByCourseSchoolYearSemester(int SchoolYear,int Semester)
        {            
            List<CourseScoreBase> retVal = new List<CourseScoreBase>();
            // 取得課程資料，只需要評分才顯示
            QueryHelper qh = new QueryHelper();
            string query = "select course.id,course.course_name from course  where course.school_year=" + SchoolYear + " and course.semester=" + Semester + " and not_included_in_calc='0' order by course.course_name,course.id;";
            DataTable dt = qh.Select(query);
            foreach (DataRow dr in dt.Rows)
            {
                CourseScoreBase csb = new CourseScoreBase();
                csb.CourseID = dr[0].ToString();
                csb.CourseName = dr[1].ToString();
                retVal.Add(csb);            
            }

            // 取得課程人數
            Dictionary<string, int> data1 = GetCourseStudentCountBySchoolYearSemester(SchoolYear, Semester);

            // 取得有成績人數
            Dictionary<string, int> data2 = GetCourseStudentScoreCountBySchoolYearSemester(SchoolYear, Semester);

            // 取得課程授課教師
            Dictionary<string, string> data3 = GetCousreTeacherNameBySchoolYearSemester(SchoolYear, Semester);

            // 填入人數值
            foreach (CourseScoreBase csb in retVal)
            {
                if (data1.ContainsKey(csb.CourseID))
                    csb.CourseStudentCount = data1[csb.CourseID];

                if (data2.ContainsKey(csb.CourseID))
                    csb.hasScoreCount = data2[csb.CourseID];

                if (data3.ContainsKey(csb.CourseID))
                    csb.TeacherName = data3[csb.CourseID];
            }

            return retVal;
        }

        /// <summary>
        /// 透過學年度學期取得課程人數(過濾學生只有狀態為一般)
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, int> GetCourseStudentCountBySchoolYearSemester(int SchoolYear, int Semester)
        {
            Dictionary<string, int> retVal = new Dictionary<string, int>();
            QueryHelper qh = new QueryHelper();
            string query = "select course.id,count(sc_attend.ref_student_id) from course inner join sc_attend on course.id =sc_attend.ref_course_id inner join student on sc_attend.ref_student_id=student.id where course.school_year="+SchoolYear+" and course.semester="+Semester+" and student.status=1 group by course.id order by course.id;";
            DataTable dt = qh.Select(query);
            foreach (DataRow dr in dt.Rows)
            {
                retVal.Add(dr[0].ToString(), int.Parse(dr[1].ToString()));
            }
            return retVal;
        }

        /// <summary>
        /// 透過學年度學期取得課程有成績人數(過濾學生只有狀態為一般)
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, int> GetCourseStudentScoreCountBySchoolYearSemester(int SchoolYear, int Semester)
        {
            Dictionary<string, int> retVal = new Dictionary<string, int>();
            QueryHelper qh = new QueryHelper();
            string query = "select course.id,count(sc_attend.ref_student_id) from course inner join sc_attend on course.id =sc_attend.ref_course_id inner join student on sc_attend.ref_student_id=student.id where student.status=1 and course.school_year="+SchoolYear+" and course.semester="+Semester+" and score is not null group by course.id order by id;";
            DataTable dt = qh.Select(query);
            foreach (DataRow dr in dt.Rows)
            {
                retVal.Add(dr[0].ToString(), int.Parse(dr[1].ToString()));
            }
            return retVal;
        }

        /// <summary>
        /// 透過學年度學期取得課程授課教師1
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetCousreTeacherNameBySchoolYearSemester(int SchoolYear, int Semester)
        {
            Dictionary<string, string> retVal = new Dictionary<string, string>();
            QueryHelper qh = new QueryHelper();
            string query = "select course.id,teacher.teacher_name from tc_instruct inner join course on tc_instruct.ref_course_id=course.id inner join teacher on tc_instruct.ref_teacher_id=teacher.id where tc_instruct.sequence=1 and course.school_year="+SchoolYear+" and course.semester="+Semester;
            DataTable dt = qh.Select(query);
            foreach (DataRow dr in dt.Rows)
            {
                retVal.Add(dr[0].ToString(), dr[1].ToString());
            }
            return retVal;
        }
    }
}
