using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Newtonsoft.Json;

namespace SmartSchool.Evaluation.Process.Wizards.LearningHistory
{
    public class JSubjectDetail
    {
        [JsonProperty("serialNo", Order = 1)]
        public string serialNo { get; set; }

        [JsonProperty("name", Order = 2)]
        public string name { get; set; }

        [JsonProperty("value", Order = 3)]
        public string value { get; set; }
    }
}
