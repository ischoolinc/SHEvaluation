using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SmartSchool.Evaluation.Process.Wizards.LearningHistory
{
    public class JSubjectInfo
    { // refStudentID
        [JsonProperty("refStudentID", Order = 1)]
        public string refStudentID { get; set; }

        [JsonProperty("serialNo", Order = 2)]
        public string serialNo { get; set; }

        [JsonProperty("name", Order = 3)]
        public string name { get; set; }

        [JsonProperty("schoolYear", Order = 4)]
        public int schoolYear { get; set; }

        [JsonProperty("semester", Order = 5)]
        public int semester { get; set; }

        [JsonProperty("subject", Order = 6)]
        public string subject { get; set; }

        [JsonProperty("subjectLevel", Order = 7)]
        public int? subjectLevel { get; set; }

        [JsonProperty("detail", Order = 8)]
        public List<JSubjectDetail> detail { get; set; }
    }
}
