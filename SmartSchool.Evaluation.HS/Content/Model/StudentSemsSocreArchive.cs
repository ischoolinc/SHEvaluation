using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SmartSchool.Evaluation.Content.Model
{
    //學期成績封存
    class StudentSemsSocreArchive
    {
        /// <summary>
        /// 學期科目成績封存uid
        /// </summary>
        public string Uid { get; set; }


        /// <summary>
        /// 學期科目成績對應的分項成績封存uid
        /// </summary>
        public string RefEntryUid { get; set; }

    }
}
