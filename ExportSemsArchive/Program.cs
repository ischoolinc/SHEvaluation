using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using FISCA.Permission;
using FISCA.Presentation;
using FISCA.Presentation.Controls;
using K12.Data;
using K12.Presentation;

namespace ExportSemsArchive
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {
            //設定權限
            Catalog catalog = RoleAclSource.Instance["學生"]["報表"];
            catalog.Add(new RibbonFeature(Permissions.學期成績封存報表, "學期成績(封存)報表"));

            //匯出學期成績(封存)
            RibbonBarItem item = MotherForm.RibbonBarItems["學生", "資料統計"];
            item["報表"]["其它相關報表"]["學期成績(封存)報表"].Enable = UserAcl.Current[Permissions.學期成績封存報表].Executable;
            item["報表"]["其它相關報表"]["學期成績(封存)報表"].Click += delegate
            {
                if (NLDPanels.Student.SelectedSource.Count > 0) //有選擇學生
                {
                    ExportSemsArchive.Report.ExportSemsArchiveData data = new Report.ExportSemsArchiveData(NLDPanels.Student.SelectedSource);
                    data.Run();
                }
                else
                {
                    MsgBox.Show("請選擇學生。");
                }
            };
        }
    }
}
