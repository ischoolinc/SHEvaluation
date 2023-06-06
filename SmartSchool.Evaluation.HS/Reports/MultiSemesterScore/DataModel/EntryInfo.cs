using System.Collections.Generic;

namespace SmartSchool.Evaluation.Reports.MultiSemesterScore.DataModel
{
    public class EntryInfo : ScoreInfo
    {
        public EntryInfo(string name)
            : base(name)
        {
        }
    }

    public class EntryCollection : Dictionary<string, EntryInfo>
    {
    }
}
