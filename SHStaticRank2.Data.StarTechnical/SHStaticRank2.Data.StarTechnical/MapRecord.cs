using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace SHStaticRank2.Data
{
    [FISCA.UDT.TableName("ischool.StudentSHStaticRank2Data.StarTechnical.Map")]
    public class MapRecord : FISCA.UDT.ActiveRecord
    {
        public MapRecord()
        {
        }
        [FISCA.UDT.Field]
        public string student_tag { get; set; }
        [FISCA.UDT.Field]
        public string code1 { get; set; }        
        [FISCA.UDT.Field]
        public string code2 { get; set; }
        [FISCA.UDT.Field]
        public string note { get; set; }
    }
}
