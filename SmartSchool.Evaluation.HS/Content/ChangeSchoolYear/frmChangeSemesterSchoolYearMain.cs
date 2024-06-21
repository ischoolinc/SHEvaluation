using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;

namespace SmartSchool.Evaluation.Content.ChangeSchoolYear
{
    public partial class frmChangeSemesterSchoolYearMain : BaseForm
    {
        // 來源訊息
        private string SourceMessage = "";

        public frmChangeSemesterSchoolYearMain()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            btnChange.Enabled = false;
            
            // 開啟確認視窗
            

            btnChange.Enabled = true;
        }

        private void frmChangeSemesterSchoolYearMain_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            labelSourceMsg.Text = SourceMessage;

            try
            {
                // 學年度選項
                int defSchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
                for (int i = defSchoolYear + 3; i >= defSchoolYear - 3; i--)
                    cboSchoolYear.Items.Add(i + "");

            }
            catch (Exception ex)
            { Console.WriteLine(ex.Message); }


            // 填入預設學年度            
            cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
        }

        // 設定來源學年度、學期、年級
        public void SetSourceInfo(string SchoolYear, string Semester, string GradeYear)
        {
            SourceMessage = string.Format("{0}學年度第{1}學期成績年級{2}年級", SchoolYear, Semester, GradeYear);
        }
    }
}
