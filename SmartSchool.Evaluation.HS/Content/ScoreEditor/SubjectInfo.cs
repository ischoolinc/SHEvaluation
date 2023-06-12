using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartSchool.Evaluation.Content.ScoreEditor
{
    public class SubjectInfo
    {
        // 科目名稱
        public string SubjectName { get; set; }

        // 科目級別
        public string SubjectLevel { get; set; }

        // 領域
        public string Domain { get; set; }

        // 分項
        public string Entry { get; set; }

        // 校部定
        public string RequiredBy { get; set; }

        // 必選修
        public string Required { get; set; }

        // 學分數
        public string Credit { get; set; }

        // 不計學分
        public string NotIncludedInCredit { get; set; }

        // 不需評分
        public string NotIncludedInCalc { get; set; }

        // 課程代碼
        public string CourseCode { get; set; }

        // 指定學年科目名稱
        public string SchoolYearSubjectName { get; set; }

        // 報部科目名稱
        public string DeptSubjectName { get; set; }

        // 學期成績已經有科目名稱級別
        public bool HasSubjectNameLevel { get; set; }
    }
}
