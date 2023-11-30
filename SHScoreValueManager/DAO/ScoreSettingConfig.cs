using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SHScoreValueManager.DAO
{
    public class ScoreSettingConfig
    {
        //    <Setting>
        //	<UseText>輸入內容</UseText>
        //	<UseValue>-1</UseValue>
        //	<ReportValue>缺考原因</ReportValue>
        //	<ScoreType>0分</ScoreType>
        //</Setting>

        // 輸入內容
        public string UseText { get; set; }

        // 成績儲存值
        public string UseValue { get; set; }

        // 缺考原因
        public string ReportValue { get; set; }

        // 分數認定
        public string ScoreType { get; set; }

    }
}
