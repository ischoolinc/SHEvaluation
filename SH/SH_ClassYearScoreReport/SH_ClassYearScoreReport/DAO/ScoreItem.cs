using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SH_ClassYearScoreReport.DAO
{
    public class ScoreItem
    {
        /// <summary>
        /// 名稱
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 成績
        /// </summary>
        public decimal Score { get; set; }

        /// <summary>
        /// 排名
        /// </summary>
        public int Rank { get; set; }

        /// <summary>
        /// 成績人數
        /// </summary>
        public int RankT { get; set; }
    }
}

