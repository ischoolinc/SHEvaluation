using System;
using System.Collections.Generic;
using System.Data;
using System.Text;
using System.Xml;
//using SmartSchool.ClassRelated;
using SmartSchool.Customization.Data;

namespace SmartSchool.Evaluation.GraduationPlan
{
    public class GraduationPlanSimple
    {
        private string _SubjectName; // 科目名稱
        private string _Code;        // 科目代碼
        private string _Domain;      // 領域
        private string _Entry;       // 分項
        private string _RequiredBy;  // 校部定
        private string _Required;    // 必選修
        private string _LevelList;   // 級別
        public GraduationPlanSimple(DataRow dr)
        {
            _SubjectName = dr["SubjectName"].ToString();
            _Code = dr["Code"].ToString();
            _Domain = dr["Domain"].ToString();
            _Entry = dr["Entry"].ToString();
            _RequiredBy = dr["RequiredBy"].ToString();
            _Required = dr["Required"].ToString();
            _LevelList = dr["LevelList"].ToString();
        }
        public string SubjectName { get { return _SubjectName; } }
        public string Code { get { return _Code; } }
        public string Domain { get { return _Domain; } }
        public string Entry { get { return _Entry; } }
        public string RequiredBy { get { return _RequiredBy; } }
        public string Required { get { return _Required; } }
        public string LevelList { get { return _LevelList; } }
    }    
}
