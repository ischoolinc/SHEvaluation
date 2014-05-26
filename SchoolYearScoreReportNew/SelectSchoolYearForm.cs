using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Reflection;

namespace SchoolYearScoreReport
{
    public partial class SelectSchoolYearForm : FISCA.Presentation.Controls.BaseForm
    {
        private SchoolYearScoreReport.Config _config;

        public SelectSchoolYearForm(SchoolYearScoreReport.Config config)
        {
            InitializeComponent();

            //string DALMessage = "『";

            //foreach (Assembly Assembly in AppDomain.CurrentDomain.GetAssemblies().Where(x => x.GetName().Name.Equals("SchoolYearScoreReportNew")))
            //    DALMessage += "版本號：" + Assembly.GetName().Version + " ";

            //DALMessage += "』";

            //this.Text += DALMessage;

            this._config = config;
            this.numUD.Value = this.Config.SchoolYear;

            this.buttonX1.AccessibleRole = AccessibleRole.PushButton;
            this.buttonX2.AccessibleRole = AccessibleRole.PushButton;
        }
        
        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.Config.SchoolYear = this.numUD.Value;
            base.DialogResult = DialogResult.OK;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new ConfigForm(this._config).ShowDialog();
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            new SelectTypeForm("學年成績單", this._config).ShowDialog();
        }

        private SchoolYearScoreReport.Config Config
        {
            get
            {
                return this._config;
            }
        }

    }
}
