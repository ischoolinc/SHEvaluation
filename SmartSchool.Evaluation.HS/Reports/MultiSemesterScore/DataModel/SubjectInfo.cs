using System.Collections.Generic;

namespace SmartSchool.Evaluation.Reports.MultiSemesterScore.DataModel
{
    public class SubjectInfo : ScoreInfo
    {
        public SubjectInfo(string name)
            : base(name)
        {
        }
    }

    public class SubjectCollection : Dictionary<string, SubjectInfo>
    {
    }
}
