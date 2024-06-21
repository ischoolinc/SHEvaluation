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
    public partial class frmChangeSemesterSchoolYearMsg : BaseForm
    {
        private string SourceMessage = "";
        private string ChangeMessage = "";

        private string SourceSchoolYear = "", ChangeSchoolYear = "", Semester = "", GradeYear = "";

        public frmChangeSemesterSchoolYearMsg()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            btnChange.Enabled = false;

            btnChange.Enabled = true;
        }

        private void frmChangeSemesterSchoolYearMsg_Load(object sender, EventArgs e)
        {
            this.MinimumSize = this.MaximumSize = this.Size;
            labelSourceMsg.Text = SourceMessage;
            labelChangeMsg.Text = ChangeMessage;
        }

        public void SetSourceMessage(string SchoolYear, string Semester, string GradeYear)
        {
            SourceSchoolYear = SchoolYear;
            this.Semester = Semester;
            this.GradeYear = GradeYear;

            SourceMessage = "將學年度：" + SourceSchoolYear + " 學期：" + Semester + " 年級：" + GradeYear;         
        }
    

        public void SetChangeMessage(string SchoolYear, string Semester, string GradeYear)
        {
            ChangeSchoolYear = SchoolYear;
            this.Semester = Semester;
            this.GradeYear = GradeYear;

            ChangeMessage = "調整為學年度：" + ChangeSchoolYear + " 學期：" + Semester + " 年級：" + GradeYear;
        }        
    }
}
