using System;
using System.Windows.Forms;
using SmartSchool.Common;
using SmartSchool.Evaluation.Process.Rating;

namespace SmartSchool.Evaluation.Process.Wizards
{
    public partial class CalcSchoolYearScoreWizard : BaseForm
    {
        public CalcSchoolYearScoreWizard()
        {
            InitializeComponent();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            // 因為目前學年科目成績計算會有BUG，所以留下此後門供客服人員使用，當使用Control+Shift+滑鼠點擊就能正常開啟功能 --2022/12/22 俊緯留
            if (Control.ModifierKeys == (Keys.Control | Keys.Shift))
            {
                new CalcSchoolYearSubjectScoreWizard(SelectType.GradeYearStudent).ShowDialog();
            }
            else
            {
                if (MessageBox.Show("欲計算學年成績，請與客服人員聯絡。\r\n選取「Yes」與客服聯絡。", "提示訊息", MessageBoxButtons.YesNo, MessageBoxIcon.Stop) == DialogResult.Yes)
                {
                    System.Diagnostics.Process.Start(@"https://line.me/R/ti/p/@vvo4068m");
                }
            }
        }

        private void buttonX5_Click(object sender, EventArgs e)
        {
            new CalcSchoolYearApplyScoreWizard(SelectType.GradeYearStudent).ShowDialog();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            new CalcSchoolYearEntryScoreWizard(SelectType.GradeYearStudent).ShowDialog();
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            new SchoolYearRatingForm().ShowDialog();
        }

        private void buttonX4_Click(object sender, EventArgs e)
        {
            new CalcSchoolYearEntryScoreWizard(SelectType.GradeYearStudent, SchoolYearScoreCalcType.SchoolYearSubject).ShowDialog();
        }
    }
}