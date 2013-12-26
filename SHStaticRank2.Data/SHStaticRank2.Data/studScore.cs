using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SHStaticRank2.Data
{
    /// <summary>
    /// 報表用學生成績
    /// </summary>
    public class studScore
    {
        /// <summary>
        /// 學生系統編號
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 總分
        /// </summary>
        public decimal sumScore=0;
        /// <summary>
        /// 加權總分
        /// </summary>
        public decimal sumScoreA = 0;

        /// <summary>
        /// 學分加總
        /// </summary>
        public int sumCredit = 0;

        /// <summary>
        /// 平均
        /// </summary>
        public decimal avgScore = 0;

        /// <summary>
        /// 加權平均
        /// </summary>
        public decimal avgScoreA = 0;

        /// <summary>
        /// 科目個數
        /// </summary>
        public int subjCount=0;


        /// <summary>
        /// 總分(類別1)
        /// </summary>
        public decimal sumScoreC1 = 0;
        /// <summary>
        /// 加權總分(類別1)
        /// </summary>
        public decimal sumScoreAC1 = 0;

        /// <summary>
        /// 學分加總(類別1)
        /// </summary>
        public int sumCreditC1 = 0;

        /// <summary>
        /// 平均(類別1)
        /// </summary>
        public decimal avgScoreC1 = 0;

        /// <summary>
        /// 加權平均(類別1)
        /// </summary>
        public decimal avgScoreAC1 = 0;

        /// <summary>
        /// 科目個數(類別1)
        /// </summary>
        public int subjCountC1 = 0;

        /// <summary>
        /// 總分(類別2)
        /// </summary>
        public decimal sumScoreC2 = 0;
        /// <summary>
        /// 加權總分(類別2)
        /// </summary>
        public decimal sumScoreAC2 = 0;

        /// <summary>
        /// 學分加總(類別2)
        /// </summary>
        public int sumCreditC2 = 0;

        /// <summary>
        /// 平均(類別2)
        /// </summary>
        public decimal avgScoreC2 = 0;

        /// <summary>
        /// 加權平均(類別2)
        /// </summary>
        public decimal avgScoreAC2 = 0;

        /// <summary>
        /// 科目個數(類別2)
        /// </summary>
        public int subjCountC2 = 0;

        /// <summary>
        /// 一上成績
        /// </summary>
        public decimal? gsScore11 { get; set; }

        /// <summary>
        /// 一下成績
        /// </summary>
        public decimal? gsScore12 { get; set; }

        /// <summary>
        /// 二上成績
        /// </summary>
        public decimal? gsScore21 { get; set; }

        /// <summary>
        /// 二下成績
        /// </summary>
        public decimal? gsScore22 { get; set; }

        /// <summary>
        /// 三上成績
        /// </summary>
        public decimal? gsScore31 { get; set; }

        /// <summary>
        /// 三下成績
        /// </summary>
        public decimal? gsScore32 { get; set; }

        /// <summary>
        /// 四上成績
        /// </summary>
        public decimal? gsScore41 { get; set; }

        /// <summary>
        /// 四下成績
        /// </summary>
        public decimal? gsScore42 { get; set; }

        /// <summary>
        /// 一上學年度
        /// </summary>
        public int? gsSchoolYear11 { get; set; }

        /// <summary>
        /// 一下學年度
        /// </summary>
        public int? gsSchoolYear12 { get; set; }

        /// <summary>
        /// 二上學年度
        /// </summary>
        public int? gsSchoolYear21 { get; set; }

        /// <summary>
        /// 二下學年度
        /// </summary>
        public int? gsSchoolYear22 { get; set; }

        /// <summary>
        /// 三上學年度
        /// </summary>
        public int? gsSchoolYear31 { get; set; }

        /// <summary>
        /// 三下學年度
        /// </summary>
        public int? gsSchoolYear32 { get; set; }

        /// <summary>
        /// 四上學年度
        /// </summary>
        public int? gsSchoolYear41 { get; set; }

        /// <summary>
        /// 四下學年度
        /// </summary>
        public int? gsSchoolYear42 { get; set; }

        /// <summary>
        /// 一上學分數
        /// </summary>
        public int? gsCredit11 { get; set; }

        /// <summary>
        /// 一下學分數
        /// </summary>
        public int? gsCredit12 { get; set; }

        /// <summary>
        /// 二上學分數
        /// </summary>
        public int? gsCredit21 { get; set; }

        /// <summary>
        /// 二下學分數
        /// </summary>
        public int? gsCredit22 { get; set; }

        /// <summary>
        /// 三上學分數
        /// </summary>
        public int? gsCredit31 { get; set; }

        /// <summary>
        /// 三下學分數
        /// </summary>
        public int? gsCredit32 { get; set; }

        /// <summary>
        /// 四上學分數
        /// </summary>
        public int? gsCredit41 { get; set; }

        /// <summary>
        /// 四下學分數
        /// </summary>
        public int? gsCredit42 { get; set; }

    }
}
