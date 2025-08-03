using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace StudentDuplicateSubjectCheck
{
    /// <summary>
    /// 學生成績基礎資料
    /// </summary>
    public class SCAttendRecord
    {
        //SCAttend Table 每一筆資料的ID
        public string ID { get; set; }

        //SCAttend Table 每一筆資料的 參考學生ID
        public string RefStudentID { get; set; }

        //SCAttend Table 每一筆資料的 參考課程ID
        public string RefCourseID { get; set; }

        /// <summary>
        /// 學生年級
        /// </summary>
        public string GradeYear { get; set; }


        /// <summary>
        /// 學生學號
        /// </summary>
        public string StudentNumber { get; set; }

        /// <summary>
        /// 學生班級
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 學生座號
        /// </summary>
        public string SeatNo { get; set; }

        /// <summary>
        /// 學生姓名
        /// </summary>
        public string Name { get; set; }


        /// <summary>
        /// 課程名稱
        /// </summary>
        public string CourseName { get; set; }

        /// <summary>
        /// 科目名稱
        /// </summary>
        public string SubjectName { get; set; }


        /// <summary>
        /// 科目級別
        /// </summary>
        public string SubjectLevel { get; set; }


        /// <summary>
        /// 課程參與資料 extensions 欄位
        /// </summary>
        public string Extensions { get; set; }

        // 成績來源
        public List<string> ScoreSource = new List<string>();
    }
}
