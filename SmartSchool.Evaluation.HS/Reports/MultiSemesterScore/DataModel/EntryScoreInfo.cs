using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartSchool.Evaluation.Reports.MultiSemesterScore.DataModel
{
    public class EntryScoreInfo
    {
        public string EntryName { get; set; }

        public decimal Score { get; set; }

        // 原始成績
        public decimal OriginScore { get; set; }
    }
}
