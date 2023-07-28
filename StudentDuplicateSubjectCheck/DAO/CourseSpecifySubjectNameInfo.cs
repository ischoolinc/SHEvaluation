using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentDuplicateSubjectCheck.DAO
{
    public class CourseSpecifySubjectNameInfo
    {
        // 課程系統編號
        public string CourseID { get; set; }

        // 學年度
        public string SchoolYear { get; set; }

        // 學期
        public string Semester { get; set; }

        // 課程名稱
        public string CourseName { get; set; }

        // 科目名稱
        public string SubjectName { get; set; }

        // 科目級別
        public string SubjectLevel { get; set; }

        // 指定學年科目名稱
        public string SpecifySubjectName { get; set; }

        // 指定學年科目名稱(舊)
        public string SpecifySubjectNameOld { get; set; }

    }
}
