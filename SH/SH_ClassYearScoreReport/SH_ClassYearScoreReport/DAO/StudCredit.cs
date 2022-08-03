using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SH_ClassYearScoreReport.DAO
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
        /// 學年應得學分
        /// </summary>
        public decimal shouldGetCredit = 0;

        /// <summary>
        /// 學年實得學分
        /// </summary>
        public decimal gotCredit = 0;

        /// <summary>
        /// 應得學分累計
        /// </summary>
        public decimal shouldGetTotalCredit = 0;

        /// <summary>
        /// 實得學分累計
        /// </summary>
        public decimal gotTotalCredit = 0;
    }
}
