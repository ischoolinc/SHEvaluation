using FISCA.Permission;
using FISCA.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SH_ClassYearScoreReport
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            //var btn = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["班級學年成績單"];
            //btn.Enable = false;
            //K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            //{
            //    btn.Enable = Permissions.期末成績通知單_固定排名權限 && (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0);
            //};
            //btn.Click += new EventHandler(Program_Click);

            RibbonBarItem rbItem1 = MotherForm.RibbonBarItems["班級", "資料統計"];
            rbItem1["報表"]["成績相關報表"]["班級學年成績單"].Enable = UserAcl.Current["SH_ClassYearScoreReport"].Executable;
            rbItem1["報表"]["成績相關報表"]["班級學年成績單"].Click += delegate
            {
                if (K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0)
                {

                    MainForm mainForm = new MainForm();
                    mainForm.SetClassIDList(K12.Presentation.NLDPanels.Class.SelectedSource);
                    mainForm.ShowDialog();
                }
                else
                {
                    FISCA.Presentation.Controls.MsgBox.Show("請選擇選班級。");
                    return;
                }

            };



            //權限設定
            Catalog permission = RoleAclSource.Instance["班級"]["功能按鈕"];
            permission.Add(new RibbonFeature("SH_ClassYearScoreReport", "班級學年成績單"));
        }

    }
}
