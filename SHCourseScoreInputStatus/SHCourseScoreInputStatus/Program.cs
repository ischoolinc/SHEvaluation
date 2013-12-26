using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Permission;
using FISCA.Presentation;

namespace SHCourseScoreInputStatus
{
    /// <summary>
    /// 課程成績輸入狀況
    /// </summary>
    public class Program
    {
        [FISCA.MainMethod()]
        public static void Main()
        {
            string regCode1 = "SHSchool.SHCourseScoreInputStatus";

            // 課程成績輸入狀況
            RibbonBarItem rptItem = MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"];
            rptItem["成績作業"]["課程成績輸入狀況"].Enable = UserAcl.Current[regCode1].Executable;
            rptItem["成績作業"]["課程成績輸入狀況"].Click += delegate
            {
                CourseScoreInputForm csif = new CourseScoreInputForm();
                csif.ShowDialog();
            };

            // 課程成績輸入狀況
            Catalog catalog1 = RoleAclSource.Instance["教務作業"]["課程成績輸入狀況"];
            catalog1.Add(new RibbonFeature(regCode1, "課程成績輸入狀況"));

        }
    }
}
