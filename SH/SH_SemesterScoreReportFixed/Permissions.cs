using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace SH_SemesterScoreReportFixed
{
    class Permissions
    {
        public static string 期末成績通知單_固定排名 { get { return "SH_SemesterScoreReport.506FECC7-3DC3-4656-8B6F-FF11CE69BD29"; } }

        public static bool 期末成績通知單_固定排名權限
        {
            get { return FISCA.Permission.UserAcl.Current[期末成績通知單_固定排名].Executable; }
        }
    }
}

