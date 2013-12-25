using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using SmartSchool.Common;
using SmartSchool.Evaluation.Process.Rating;

namespace SmartSchool.Evaluation.Process.Wizards
{
    public partial class CalsSemesterScoreWizard : BaseForm
    {
        public CalsSemesterScoreWizard()
        {
            InitializeComponent();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            new CalcSemesterSubjectScoreWizard(SelectType.GradeYearStudent).ShowDialog();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            new CalcSemesterEntryScoreWizard(SelectType.GradeYearStudent).ShowDialog();
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            new SemesterRatingForm().ShowDialog();
        }
    }
}