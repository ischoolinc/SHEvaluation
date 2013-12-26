using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.Customization.Data;
using System.IO;
using FISCA.DSAUtil;

namespace SHStaticRank2.Data
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            CalcMutilSemeSubjectRank.Setup();
        }
    }
}
