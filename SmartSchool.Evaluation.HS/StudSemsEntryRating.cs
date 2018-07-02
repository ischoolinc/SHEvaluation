using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace SmartSchool.Evaluation
{
    public class StudSemsEntryRating
    {
        public string StudentID { get; set; }
        public string SchoolYear { get; set; }
        public string Semester { get; set; }

        public string GradeYear { get; set; }

        public string EntryName { get; set; }

        /// <summary>
        /// 班成績人數
        /// </summary>
        public int? ClassCount { get; set; }

        /// <summary>
        /// 班排名
        /// </summary>
        public int? ClassRank { get; set; }

        /// <summary>
        /// 科成績人數
        /// </summary>
        public int? DeptCount { get; set; }

        /// <summary>
        /// 科排名
        /// </summary>
        public int? DeptRank { get; set; }

        /// <summary>
        /// 校成績人數
        /// </summary>
        public int? YearCount { get; set; }

        /// <summary>
        /// 校排名
        /// </summary>
        public int? YearRank { get; set; }


        /// <summary>
        /// 類1名稱
        /// </summary>
        public string Group1 { get; set; }

        /// <summary>
        /// 類1成績人數
        /// </summary>
        public int? Group1Count { get; set; }

        /// <summary>
        /// 類1排名
        /// </summary>
        public int? Group1Rank { get; set; }

        // 2018/7/2 穎驊，檢驗成績系統，支援匯入匯出，在此新增類2邏輯

        /// <summary>
        /// 類1名稱
        /// </summary>
        public string Group2 { get; set; }

        /// <summary>
        /// 類1成績人數
        /// </summary>
        public int? Group2Count { get; set; }

        /// <summary>
        /// 類1排名
        /// </summary>
        public int? Group2Rank { get; set; }

        /// <summary>
        /// 分數
        /// </summary>
        public string Score { get; set; }



    }
}
