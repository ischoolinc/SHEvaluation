using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartSchool.Evaluation.Content.ChangeSchoolYear
{
    public class SubjectScoreInfo
    {

        // 學年度
        public string SchoolYear { get; set; }
        // 學期
        public string Semester { get; set; }
        // 年級
        public string GradeYear { get; set; }

        // 科目名稱
        public string SubjectName { get; set; }
        // 科目級別
        public string SubjectLevel { get; set; }
        // 科目學分
        public string Credit { get; set; }

        // 必選修
        public string Required { get; set; }

        // 校部定
        public string RequiredBy { get; set; }
        

    }
}
