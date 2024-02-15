using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StudentCourseScoreCheck100_SH.DAO
{
    /// <summary>
    /// 學生成績基礎資料
    /// </summary>
    public class StudentCourseScoreBase
    {
        public string StudentID { get; set; }

        /// <summary>
        /// 年級
        /// </summary>
        public string GradeYear { get; set; }


        /// <summary>
        /// 學號
        /// </summary>
        public string StudentNumber { get; set; }

        /// <summary>
        /// 班級
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 座號
        /// </summary>
        public string SeatNo { get; set; }

        /// <summary>
        /// 姓名
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 評量成績(試別、成績)(courseid,examName,score)
        /// </summary>
        public Dictionary<string, Dictionary<string, decimal>> ExamScoreDict = new Dictionary<string, Dictionary<string, decimal>>();

        /// <summary>
        /// 課程成績績(courseid)
        /// </summary>
        public Dictionary<string, decimal> CousreScoreDict = new Dictionary<string, decimal>();

        /// <summary>
        /// 課程教師(courseid,name)
        /// </summary>
        public Dictionary<string, string> CourseTeacherDict = new Dictionary<string, string>();

        /// <summary>
        /// 課程編號,名稱(courseid,name)
        /// </summary>
        public Dictionary<string, string> CourseNameDict = new Dictionary<string, string>();

        /// <summary>
        /// 評量成績(試別、使用文字)(courseid,examName,UseText)
        /// </summary>
        public Dictionary<string, Dictionary<string, string>> ExamScoreTextDict = new Dictionary<string, Dictionary<string, string>>();
    }
}
