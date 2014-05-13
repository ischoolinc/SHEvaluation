using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SH_SemesterScoreReport
{
    class Permissions
    {
        public static string 期末成績通知單 { get { return "SH_SemesterScoreReport.6D7169EF-5BAD-41C5-B999-0946F37A6F92"; } }

        public static bool 期末成績通知單權限
        {
            get { return FISCA.Permission.UserAcl.Current[期末成績通知單].Executable; }
        }
    }
}
