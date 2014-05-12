using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 定期評量成績單
{
    class Permissions
    {
        public static string 定期評量成績單 { get { return "定期評量成績單.D04E7F02-89C1-4412-81FA-8D87B96BF847"; } }

        public static bool 定期評量成績單權限
        {
            get { return FISCA.Permission.UserAcl.Current[定期評量成績單].Executable; }
        }
    }
}
