using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 班級定期評量成績單_固定排名.Model
{


    //【總計成績】一個*學生下 學年度學期的* 某次考試、  有  *多種(班排、類別排、科排、校排)母群  3.比較標的(總分(算數)、加權總分、平均(算術平均)、加權平均)
    //【科目成績】一個*學生下 學年度學期的* 某次考試、  有  *多種(班排、類別排、科排、校排)母群  3.比較標的(科目)
    class StudentFixedRankInfo
    {
        public string StudnetID { get; set; }

        public string StuName { get; set; }

        public string SchoolYear { get; set; }

        public string Semester { get; set; }

        /// <summary>
        /// 【總計成績】
        /// key:rank_name(班排名、科排名、校排名)_item_name(加權平均、算術平均....) 、 rankInfo {item_name,rank,matrix_count})
        /// </summary>
        //public Dictionary<string, Dictionary<string, RankInfo>> DicSubjectTotalFixRank { get; set; }
        public Dictionary<string, RankInfo> DicSubjectTotalFixRank { get; set; }

        /// <summary>
        /// 【科目成績】
        /// key:rank_name(班排名、科排名、校排名)_item_name(國文、數學.....) 、 rankInfo {item_name,rank,matix_count})
        /// </summary>
        //  public Dictionary<string, Dictionary<string, RankInfo>> DicSubjectFixRank { get; set; }
        public Dictionary<string, RankInfo> DicSubjectFixRank { get; set; }

        public StudentFixedRankInfo(string stuID)
        {
            this.StudnetID = stuID;
            this.DicSubjectTotalFixRank = new Dictionary<string, RankInfo>();
            this.DicSubjectFixRank = new Dictionary<string, RankInfo>();
        }

    }
}
