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
        private string _Attriubte;   // 科目屬性
        private string _Entry;       // 分項
        private string _RequiredBy;  // 校部定
        private string _Required;    // 必選修
        private string _LevelList;   // 級別

        Dictionary<string, string> cAttrib = new Dictionary<string, string>()
        {{"0", "不分屬性" },{"1", "一般科目" },{"2", "專業科目" },{"3", "實習科目" },{"4", "專精科目" },{"5", "專精科目(核心科目)" },{"6", "特殊需求領域" },{"A", "自主學習" },{"B", "選手培訓" },{"C", "充實補強(不授予學分)" },{"D", "充實補強(授予學分)" },{"E", "學校特色活動" },{"F", "專精(專業)科目" },{"G", "專精(實習)科目" },{"H", "專精(專業)科目(核心)" },{"I", "專精(實習)科目(核心)" }};
        public GraduationPlanSimple(DataRow dr)
        {
            _SubjectName = dr["SubjectName"].ToString();
            _Code = dr["Code"].ToString();
            _Domain = dr["Domain"].ToString();
            if (cAttrib.ContainsKey(dr["CourseAttrib"].ToString()))
            {
                _Attriubte = cAttrib[dr["CourseAttrib"].ToString()];
            }
            else
            {
                _Attriubte = "";
            }
            _Entry = dr["Entry"].ToString();
            _RequiredBy = dr["RequiredBy"].ToString();
            _Required = dr["Required"].ToString();
            _LevelList = dr["LevelList"].ToString();
        }
        public string SubjectName { get { return _SubjectName; } }
        public string Code { get { return _Code; } }
        public string Domain { get { return _Domain; } }
        public string Attribute { get { return _Attriubte; } }
        public string Entry { get { return _Entry; } }
        public string RequiredBy { get { return _RequiredBy; } }
        public string Required { get { return _Required; } }
        public string LevelList { get { return _LevelList; } }
    }    
}
