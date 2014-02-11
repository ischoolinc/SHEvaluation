using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.Customization.Data;
using System.IO;
using FISCA.DSAUtil;
using SHStaticRank2.Data;

namespace SHStaticRank2.Data.StarUniversity
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            FISCA.Permission.Catalog cat = FISCA.Permission.RoleAclSource.Instance["教務作業"]["功能按鈕"];
            cat.Add(new FISCA.Permission.RibbonFeature("SHSchool.SHStaticRank2.Data", "計算固定排名(測試版)"));

            var button = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["計算固定排名(測試版)"]["計算多學期成績固定排名(103學年度大學繁星)"];
            button.Enable = FISCA.Permission.UserAcl.Current["SHSchool.SHStaticRank2.Data"].Executable;
            button.Click += delegate
            {
                var conf = new StarUniversity();
                conf.ShowDialog();
                if (conf.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    CalcMutilSemeSubjectRank.Setup(conf.Configure);
                }
            };
        }
    }
}
