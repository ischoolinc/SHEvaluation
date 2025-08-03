using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;

namespace SmartSchool.Evaluation.Process.Wizards.LearningHistory
{
    public class SubjectScoreXML
    {
        public string StudentID { get; set; }
        public string SchoolYear { get; set; }

        public string Semester { get; set; }

        public string GradeYear { get; set; }

        public XElement ScoreXML { get; set; }
    }
}
