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

            // 2018/6/22 穎驊 註解， 不再支援這個奇怪位置的學期分項成績 排名 匯入， 統一在一般的成績相關資料的 學期分項成績匯入(專案 SmartSchool.Evaluation.HS)
            //SmartSchool.Customization.PlugIn.ImportExport.ImportStudent.AddProcess(new ImportSemesterEntryScore());
        }
    }
}
