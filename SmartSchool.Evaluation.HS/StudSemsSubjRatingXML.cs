using System.Xml.Linq;

namespace SmartSchool.Evaluation
{
    public class StudSemsSubjRatingXML
    {
        public string StudentID { get; set; }
        public string SchoolYear { get; set; }
        public string Semester { get; set; }

        public string GradeYear { get; set; }

        public XElement ClassRankXML { get; set; }
        public XElement DeptRankXML { get; set; }
        public XElement YearRankXML { get; set; }

        public XElement GroupRankXML { get; set; }

    }
}
