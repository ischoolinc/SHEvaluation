using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHStaticRank.Data
{
    public class Configure
    {
        public string SchoolYear { get; set; }
        public string Semester { get; set; }
        public bool CalcGradeYear1 { get; set; }
        public bool CalcGradeYear2 { get; set; }
        public bool CalcGradeYear3 { get; set; }
        public string NotRankTag { get; set; }
        public bool use原始成績 { get; set; }
        public bool use補考成績 { get; set; }
        public bool use重修成績 { get; set; }
        public bool use手動調整成績 { get; set; }
        public bool use學年調整成績 { get; set; }
        public string Rank1Tag { get; set; }
        public string Rank2Tag { get; set; }
        public bool DoNotSaveIt { get; set; }
        public bool 計算學業成績排名 { get; set; }
        public bool 清除不排名學生排名資料 { get; set; }
    }
}
