using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ClassSemesterScoreReportFixed_SH.DAO
{
    /// <summary>
    /// 學生學分數
    /// </summary>
    public class StudCredit
    {
        /// <summary>
        /// Student ID
        /// </summary>
        public string StudentID { get; set; }

        /// <summary>
        /// 應得學分
        /// </summary>
        public int shouldGetCredit = 0;
        /// <summary>
        /// 實得學分
        /// </summary>
        public int gotCredit = 0;
        /// <summary>
        /// 應得學分累計
        /// </summary>
        public int shouldGetTotalCredit = 0;
        /// <summary>
        /// 實得學分累計
        /// </summary>
        public int gotTotalCredit = 0;
    }
}
