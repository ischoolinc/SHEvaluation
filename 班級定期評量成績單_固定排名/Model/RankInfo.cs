using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 班級定期評量成績單.Model
{
    /// <summary>
    /// 排名相關資訊(排名標的、排名母數、第幾名)
    /// </summary>
    class RankInfo
    {   
        /// <summary>
        /// 排名標的
        /// </summary>
        public string ItemName { get; set; }

        /// <summary>
        /// 第幾名
        /// </summary>
        public string Rank { get; set; }

        /// <summary>
        /// 母群數
        /// </summary>
        public string MatrixCount { get; set; }


        /// <summary>
        /// 建構子
        /// </summary>
        public RankInfo(string itemName, string rank, string matrix)
        {
            this.ItemName = itemName;

            this.Rank = rank;

            this.MatrixCount = matrix;
        }
    }
}
