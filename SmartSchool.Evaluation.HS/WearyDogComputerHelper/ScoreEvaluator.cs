using SmartSchool.Customization.Data;
using System.Collections.Generic;

namespace SmartSchool.Evaluation.WearyDogComputerHelper
{
    internal class ScoreEvaluator
    {
        private Dictionary<StudentRecord, List<string>> UnreadyStudents { get; set; }
        private AccessHelper DataSource { get; set; }
        private List<StudentRecord> SourceStudents { get; set; }
        private int SchoolYear { get; set; }

        public ScoreEvaluator(int schoolyear, AccessHelper accesshelper, List<StudentRecord> students)
        {
            DataSource = accesshelper;
            SourceStudents = students;
            UnreadyStudents = new Dictionary<StudentRecord, List<string>>();
        }
    }
}