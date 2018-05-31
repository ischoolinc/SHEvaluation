using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Presentation;
using FISCA.Permission;

namespace StudentDuplicateSubjectCheck
{
    //2018/5/24 穎驊執行[H成績][06] 課程重讀重修設定 項目

    /// <summary>
    /// 本學期修課紀錄 若其科目級別 已與舊有的學期科目成績重覆檢查
    /// </summary>
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            string regCode = "StudentDuplicateSubjectCheck";
            RibbonBarItem rptItem = MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"];
            rptItem["成績作業"]["重覆修課採計方式"].Enable = UserAcl.Current[regCode].Executable;
            rptItem["成績作業"]["重覆修課採計方式"].Click += delegate
            {
                SelectGradeYear sgyf = new SelectGradeYear();
                sgyf.ShowDialog();

            };
            
            Catalog catalog1a = RoleAclSource.Instance["教務作業"]["功能按鈕"];
            catalog1a.Add(new RibbonFeature(regCode, "重覆修課採計方式"));
        }
    }
}
