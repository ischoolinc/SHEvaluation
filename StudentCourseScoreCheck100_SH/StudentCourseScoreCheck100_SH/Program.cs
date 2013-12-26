using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Presentation;
using FISCA.Permission;

namespace StudentCourseScoreCheck100_SH
{
    /// <summary>
    /// 課程成績非0~100分檢查
    /// </summary>
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            string regCode = "SHStudentCourseScoreCheck100_SH";
            RibbonBarItem rptItem = MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"];
            rptItem["成績作業"]["課程及評量成績百分檢查"].Enable = UserAcl.Current[regCode].Executable;
            rptItem["成績作業"]["課程及評量成績百分檢查"].Click += delegate
            {
                StudentCourseScoreForm scsf = new StudentCourseScoreForm();
                scsf.ShowDialog();
            };

            // 課程成績非0~100分檢查
            Catalog catalog1a = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            catalog1a.Add(new RibbonFeature(regCode, "課程及評量成績百分檢查"));

        }
    }
}
