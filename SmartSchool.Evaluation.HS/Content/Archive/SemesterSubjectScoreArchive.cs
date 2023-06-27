using FISCA.UDT;


namespace SmartSchool.Evaluation.Content
{
    [TableName("semester_subject_score_archive")]

    class SemesterSubjectScoreArchive : ActiveRecord
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

        /// <summary>
        /// 關聯的學期分項成績
        /// </summary>
        [Field(Field = "ref_sems_entry_uid")]
        public int refEntryUid { get; set; }

        ///// <summary>
        ///// 封存日期
        ///// </summary>
        //[Field(Field = "archive_date")]
        //public DateTime ArchiveDate { get; set; }
    }
}
