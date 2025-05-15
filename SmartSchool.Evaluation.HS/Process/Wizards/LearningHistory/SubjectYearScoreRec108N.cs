using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartSchool.Evaluation.Process.Wizards.LearningHistory
{
    /// <summary>
    /// 進校學年補考使用
    /// </summary>
    public class SubjectYearScoreRec108N
    {
        public string StudentID { get; set; }

        /// <summary>
        /// 身分證號
        /// </summary>
        public string IDNumber { get; set; }

        /// <summary>
        /// 出生日期
        /// </summary>
        public string Birthday { get; set; }

        /// <summary>
        /// 課程代碼
        /// </summary>
        public string CourseCode { get; set; }

        /// <summary>
        /// 科目名稱
        /// </summary>
        public string SubjectName { get; set; }

        /// <summary>
        /// 開課年級
        /// </summary>
        public string GradeYear { get; set; }

        /// <summary>
        /// 第一學期修課學分
        /// </summary>
        public string Credit1 { get; set; }

        /// <summary>
        /// 第二學期修課學分
        /// </summary>
        public string Credit2 { get; set; }

        public string Pass1 { get; set; }

        public string Pass2 { get; set; }

        /// <summary>
        /// 第一次補考成績
        /// </summary>
        public string ReScore { get; set; }

        /// <summary>
        /// 第一次及格
        /// </summary>
        public string ScoreP { get; set; }

        public bool checkPass { get; set; }

        /// <summary>
        /// 檢查課程代碼是否可提交成績
        /// </summary>
        public bool CodePass { get; set; }

        /// <summary>
        /// 備註(學生姓名)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 備註(資料當學期班級) 
        /// 學期歷程班級
        /// </summary>
        public string HisClassName { get; set; }

        /// <summary>
        /// 備註(資料當學期座號)
        /// 學期歷程座號
        /// </summary>
        public int? HisSeatNo { get; set; }

        /// <summary>
        /// 備註(資料當學期學號)
        /// 學期歷程學號
        /// </summary>
        public string HisStudentNumber { get; set; }

        /// <summary>
        /// 備註(資料次學期班級)
        /// 現在班級
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 備註(資料次學期座號)
        /// 現在座號
        /// </summary>
        public string SeatNo { get; set; }

        /// <summary>
        /// 備註(資料次學期學號)
        /// 現在學號
        /// </summary>
        public string StudentNumber { get; set; }

        // 學年度
        public string SchoolYear { get; set; }

        // 學期
        public string Semester { get; set; }

    }
}
