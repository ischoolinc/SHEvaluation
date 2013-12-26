using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.Customization.Data;
using System.IO;
using FISCA.Presentation;
using SmartSchool.Customization.PlugIn.ImportExport;

namespace SHStaticRank.Data
{

    public static class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            CalcSemeSubjectRank.Setup();

            SmartSchool.Customization.PlugIn.ImportExport.ImportStudent.AddProcess(new ImportSemesterEntryScore());
        }
    }
}
