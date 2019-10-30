using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace RegularAssessmentTranscriptFixedRank
{
    class Permissions
    {
        public static string 定期評量成績單固定排名 { get { return "定期評量成績單固定排名.D04E7F02-89C1-4412-81FA-8D87B96BF847"; } }

        public static bool 定期評量成績單固定排名權限
        {
            get { return FISCA.Permission.UserAcl.Current[定期評量成績單固定排名].Executable; }
        }

    }
}
