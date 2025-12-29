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

        /// <summary>
        /// 課程成績來源
        /// </summary>
        public string ScoreSource { get; set; }

        /// <summary>
        /// 是否不列入成績計算（不評分課程）
        /// true = 不列入計算
        /// false = 列入計算
        /// 規則：只有 not_included_in_calc == "0" 才為 false，其餘皆為 true
        /// </summary>
        public bool NotIncludedInCalc { get; set; }

        /// <summary>
        /// DB 原始 not_included_in_calc 值（除錯用）
        /// </summary>
        public string NotIncludedInCalcRaw { get; set; }

    }
}
