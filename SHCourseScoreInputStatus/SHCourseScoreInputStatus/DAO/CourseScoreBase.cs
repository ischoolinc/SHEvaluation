using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHCourseScoreInputStatus.DAO
{
    /// <summary>
    /// 課程成績輸入狀態基本
    /// </summary>
    public class CourseScoreBase
    {
        /// <summary>
        /// 課程編號
        /// </summary>
        public string CourseID { get; set; }

        /// <summary>
        /// 課程名稱
        /// </summary>
        public string CourseName { get; set; }

        /// <summary>
        /// 授課教師名稱
        /// </summary>
        public string TeacherName { get; set; }

        /// <summary>
        /// 有成績人數
        /// </summary>
        public int? hasScoreCount { get; set; }
        
        /// <summary>
        /// 課程學生人數
        /// </summary>
        public int? CourseStudentCount { get; set; }        

    }
}
