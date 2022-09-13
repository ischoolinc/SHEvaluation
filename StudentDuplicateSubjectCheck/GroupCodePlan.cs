using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StudentDuplicateSubjectCheck
{
    /// <summary>
    /// $moe.subjectcode
    /// </summary>
    public class GroupCodePlan
    {
        /// <summary>
        /// 課程代碼
        /// </summary>
        public string course_code { get; set; }

        /// <summary>
        /// 群科班代碼
        /// </summary>
        public string group_code { get; set; }

        /// <summary>
        /// 必選修
        /// </summary>
        public string is_required { get; set; }

        /// <summary>
        /// 校部定
        /// </summary>
        public string require_by { get; set; }

        /// <summary>
        /// 科目名稱
        /// </summary>
        public string subject_name { get; set; }

        /// <summary>
        /// course_attr
        /// </summary>
        public string course_attr { get; set; }

        /// <summary>
        /// 分項類別
        /// </summary>
        public string Entry { get; set; }

        /// <summary>
        /// 六學期學分
        /// </summary>
        public string credit_period { get; set; }
    }
}
