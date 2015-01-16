using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace SHStaticRank2.Data
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
        /// 類1成績人數
        /// </summary>
        public int? Group1Count { get; set; }

        /// <summary>
        /// 類1排名
        /// </summary>
        public int? Group1Rank { get; set; }
    }
}
