using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.UDT;


namespace SmartSchool.Evaluation.Content
{
    [TableName("semester_entry_score_archive")]

    class SemesterEntryScoreArchive : ActiveRecord
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        [Field(Field = "ref_student_id")]
        public int StudentID { get; set; }

        /// <summary>
        /// 學年度
        /// </summary>
        [Field(Field = "school_year")]
        public int SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        [Field(Field = "semester")]
        public int Semester { get; set; }

        /// <summary>
        /// 成績年級
        /// </summary>
        [Field(Field = "grade_year")]
        public int GradeYear { get; set; }

        /// <summary>
        /// 成績資料
        /// </summary>
        [Field(Field = "score_info")]
        public string ScoreInfo { get; set; }

        ///// <summary>
        ///// 封存日期
        ///// </summary>
        //[Field(Field = "archive_date")]
        //public DateTime ArchiveDate { get; set; }
    }
}
