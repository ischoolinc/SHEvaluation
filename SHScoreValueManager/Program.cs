using FISCA.Permission;
using FISCA.Presentation;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using SHScoreValueManager.UIForm;

namespace SHScoreValueManager
{
    public class Program
    {
        [FISCA.MainMethod()]
        public static void main()
        {
            Catalog ribbon11 = RoleAclSource.Instance["教務作業"]["基本設定"];
            ribbon11.Add(new RibbonFeature("DB304162-AAF8-4636-B5BF-C29740842EDE", "缺考設定"));

            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["缺考設定"].Enable = UserAcl.Current["DB304162-AAF8-4636-B5BF-C29740842EDE"].Executable;
            MotherForm.RibbonBarItems["教務作業", "基本設定"]["設定"]["缺考設定"].Click += delegate
            {
                frmSetScoreValue frm = new frmSetScoreValue();
                frm.ShowDialog();
            };
        }
    }
}
