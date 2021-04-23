using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Permission;

namespace ExportSemsArchive
{
    class Permissions
    {
        public static string 學期成績封存報表 { get { return "ExportSemsArchiveData"; } }
        public static bool 學期成績封存報表權限
        {
            get
            {
                return UserAcl.Current[學期成績封存報表].Executable;
            }
        }

    }
}
