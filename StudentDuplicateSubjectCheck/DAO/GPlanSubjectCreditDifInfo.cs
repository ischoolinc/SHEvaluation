using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentDuplicateSubjectCheck.DAO
{
    public class GPlanSubjectCreditDifInfo
    {
        // 課程規劃編號
        public string GPlanID { get; set; }

        // 課程規劃名稱
        public string GPlanName { get; set; }

        // 年級
        public string GradeYear { get; set; }

        // 學期
        public string Semester { get; set; }

        // 科目名稱
        public string SubjectName { get; set; }

        // 學期學分數
        public string Credit { get; set; }

        // 指定學年科目名稱
        public string SpecifySubjectName { get; set; }
    }
}
