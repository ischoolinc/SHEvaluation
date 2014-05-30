using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Data;

namespace SemesterScoreReportNewEpost
{
    class Global
    {  /// <summary>
        /// 檢查是否有勾選產生 Excel
        /// </summary>
        public static bool _CheckExportEpost = false;
        // 存放收集到產生至 Excel 資料
        public static DataTable dt = new DataTable();
    }
}
