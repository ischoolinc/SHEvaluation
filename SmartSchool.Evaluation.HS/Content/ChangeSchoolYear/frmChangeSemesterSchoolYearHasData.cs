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
    public partial class frmChangeSemesterSchoolYearHasData : BaseForm
    {
        public frmChangeSemesterSchoolYearHasData()
        {
            InitializeComponent();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void lnkHasDataList_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkHasDataList.Enabled = false;

            lnkHasDataList.Enabled = true;
        }

        private void frmChangeSemesterSchoolYearHasData_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MaximumSize = this.Size;
        }
    }
}
