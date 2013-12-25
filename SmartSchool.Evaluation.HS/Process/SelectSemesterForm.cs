using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SmartSchool.Common;

namespace SmartSchool.Evaluation.Process
{
    public partial class SelectSemesterForm : BaseForm
    {
        public SelectSemesterForm()
        {
            InitializeComponent();

            Text = "選擇學年度學期";

            try
            {
                for (int i = -2; i <= 2; i++)
                {
                    cboSchoolYear.Items.Add(CurrentUser.Instance.SchoolYear + i);
                }
                cboSchoolYear.Text = CurrentUser.Instance.SchoolYear.ToString();
                cboSemester.Text = CurrentUser.Instance.Semester.ToString();
            }
            catch (Exception ex)
            {
                CurrentUser.ReportError(ex);
            }
        }

        private EnhancedErrorProvider ErrorProvider
        {
            get { return errorProvider1; }
        }

        public int SchoolYear
        {
            get
            {
                int a;
                if (int.TryParse(cboSchoolYear.Text, out a))
                    return a;
                return CurrentUser.Instance.SchoolYear;
            }
        }

        public int Semester
        {
            get
            {
                int a;
                if (int.TryParse(cboSemester.Text, out a))
                    return a;
                return CurrentUser.Instance.Semester;
            }
        }

        private bool Validate()
        {
            ErrorProvider.Clear();
            int i;
            if (!int.TryParse(cboSchoolYear.Text, out i))
                ErrorProvider.SetError(cboSchoolYear, "學年度必須為數字");
            if (int.TryParse(cboSchoolYear.Text, out i) && i <= 0)
                ErrorProvider.SetError(cboSchoolYear, "學年度必須為正整數");
            if (!int.TryParse(cboSemester.Text, out i))
                ErrorProvider.SetError(cboSemester, "學期必須為 1 或 2");

            return !ErrorProvider.HasError;
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (!Validate())
                return;

            DialogResult = DialogResult.OK;
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}