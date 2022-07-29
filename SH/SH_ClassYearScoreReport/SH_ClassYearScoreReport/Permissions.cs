using FISCA.Permission;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SH_ClassYearScoreReport
{
    class Permissions
    {
        public static string SH_ClassYearScoreReport { get { return "SH_ClassYearScoreReport"; } }

        public static bool SH_ClassYearScoreReportPermission
        {
            get { return UserAcl.Current[SH_ClassYearScoreReport].Executable; }
        }
    }
}
