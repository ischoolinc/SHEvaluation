using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RegularAssessmentTranscriptFixedRank.DAO
{
    // 存放評量繩成績有缺考資料
    public class StudSceTakeInfo
    {
        public string StudentID { get; set; }
        public string CourseID { get; set; }
        public string ExamName { get; set; }
        // 缺考輸入內容
        public string UseText { get; set; }
        // 缺考原因
        public string ReportValue { get; set; }
    }
}
