using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartSchool.Evaluation.Content.ChangeSchoolYear
{
    public class SemesterScoreData
    {
        // 成績 ID
        public string ID { get; set; }
        // 學年度
        public string SchoolYear { get; set; }
        // 學期
        public string Semester { get; set; }
        // 年級
        public string GradeYear { get; set; }
        // 學生ID
        public string StudentID { get; set; }

        // 成績資料 XML
        public string ScoreInfo { get; set; }        

    }
}
