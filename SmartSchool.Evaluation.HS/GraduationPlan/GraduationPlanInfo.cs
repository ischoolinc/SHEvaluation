using System.Collections.Generic;
using System.Xml;
//using SmartSchool.ClassRelated;

namespace SmartSchool.Evaluation.GraduationPlan
{
    public class GraduationPlanInfo
    {
        private readonly string _ID;
        private readonly string _Name;
        private readonly XmlElement _GraduationPlanElement;

        private readonly string _SchoolYear;
        private readonly string _TrimName;

        internal GraduationPlanInfo(XmlElement gPlanElement)
        {
            _ID = gPlanElement.GetAttribute("ID");
            _Name = gPlanElement.SelectSingleNode("Name").InnerText;
            _GraduationPlanElement = (XmlElement)gPlanElement.SelectSingleNode("Content/GraduationPlan");

            _SchoolYear = _GraduationPlanElement.HasAttribute("SchoolYear") ? _GraduationPlanElement.GetAttribute("SchoolYear") : string.Empty;
            _TrimName = _Name;
            if (!string.IsNullOrEmpty(_SchoolYear))
            {
                //_TrimName = _Name.TrimStart(_SchoolYear.ToCharArray());
                _TrimName = _Name.Substring(_SchoolYear.Length);
            }
        }

        public string ID { get { return _ID; } }
        public string Name { get { return _Name; } }
        public string SchoolYear { get { return _SchoolYear; } }
        public string TrimName { get { return _TrimName; } }
        public XmlElement GraduationPlanElement
        {
            get
            {
                return (XmlElement)(new XmlDocument().ImportNode(_GraduationPlanElement, true));
            }
        }
        public List<GraduationPlanSubject> Subjects
        {
            get
            {
                List<GraduationPlanSubject> list = new List<GraduationPlanSubject>();
                foreach (XmlNode var in _GraduationPlanElement.SelectNodes("Subject"))
                {
                    list.Add(new GraduationPlanSubject((XmlElement)var));
                }
                return list;
            }
        }

        public List<GraduationPlanSubject> SemesterSubjects(int gradeYear, int semester)
        {
            List<GraduationPlanSubject> list = new List<GraduationPlanSubject>();
            foreach (XmlNode var in _GraduationPlanElement.SelectNodes("Subject[@GradeYear='" + gradeYear + "' and @Semester='" + semester + "']"))
            {
                list.Add(new GraduationPlanSubject((XmlElement)var));
            }
            return list;
        }

        public GraduationPlanSubject GetSubjectInfo(string subjectName, string subjectLevel)
        {
            foreach (GraduationPlanSubject var in Subjects)
            {
                if (var.SubjectName == subjectName && var.Level == subjectLevel)
                {
                    return var;
                }
            }
            //todo 讀近來
            XmlDocument doc = new XmlDocument();
            //doc.LoadXml("<Subject Category=\"\" Credit=\"0\" Domain=\"\" Entry=\"學業\" FullName=\"預設\" Level=\"\" NotIncludedInCalc=\"False\" NotIncludedInCredit=\"False\" Required=\"選修\" RequiredBy=\"校訂\" SubjectName=\"預設\"/>");
            doc.LoadXml("<Subject Category=\"\" Credit=\"0\" Domain=\"\" Entry=\"學業\" FullName=\"預設\" Level=\"\" NotIncludedInCalc=\"False\" NotIncludedInCredit=\"False\" Required=\"選修\" RequiredBy=\"校訂\" SubjectName=\"預設\" 課程類別=\"\" 開課方式=\"\" 科目屬性=\"\" 學分=\"\" 領域名稱=\"\" 課程名稱=\"\"  授課學期學分=\"\" 課程代碼=\"\"/>");

            GraduationPlanSubject defaultResponse = new GraduationPlanSubject(doc.DocumentElement);
            foreach (XmlNode var in GraduationPlan.Instance.CommonPlan.SelectNodes("Subject"))
            {
                GraduationPlanSubject subjectInfo = new GraduationPlanSubject((XmlElement)var);
                if (subjectInfo.SubjectName == "預設")
                    defaultResponse = subjectInfo;
                if (subjectInfo.SubjectName == subjectName && subjectInfo.Level == subjectLevel)
                {
                    return subjectInfo;
                }
            }
            return defaultResponse;
        }
    }
}
