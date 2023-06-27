using FISCA.Presentation;
using SmartSchool.AccessControl;
using SmartSchool.Evaluation.Process.Rating;
//using SmartSchool.StudentRelated;
using SmartSchool.Evaluation.Process.Wizards;
using System;

namespace SmartSchool.Evaluation.Process
{
    [FeatureCode("Button0705")]
    public partial class CalculationBatch : SmartSchool.Evaluation.Process.RibbonBarBase
    {
        public CalculationBatch()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(CalculationBatch));
            var buttonItem103 = MotherForm.RibbonBarItems["�аȧ@�~", "�妸�@�~/�˵�"]["���Z�@�~"];
            buttonItem103.Enable = CurrentUser.Acl["Button0670"].Executable;
            //buttonItem103.Image = ( (System.Drawing.Image)( resources.GetObject("buttonItem103.Image") ) );
            buttonItem103["�Ǵ����Z�B�z"].BeginGroup = true;
            buttonItem103["�Ǵ����Z�B�z"].Enable = CurrentUser.Acl["Button0670"].Executable;
            buttonItem103["�Ǧ~���Z�B�z"].Enable = CurrentUser.Acl["Button0670"].Executable;

            buttonItem103["�Ǵ����Z�B�z"].Click += new System.EventHandler(this.buttonItem6_Click_1);
            buttonItem103["�Ǧ~���Z�B�z"].Click += new System.EventHandler(this.buttonItem8_Click_1);

            var buttonItem9 = MotherForm.RibbonBarItems["�ǰȧ@�~", "���Z�@�~"]["�w�榨�Z(�¨�)"];
            buttonItem9.Image = ((System.Drawing.Image)(resources.GetObject("buttonItem9.Image")));
            buttonItem9.Enable = CurrentUser.Acl["Button0705"].Executable;
            buttonItem9["�p��w��Ǵ����Z(�¨�)"].Click += new System.EventHandler(this.buttonItem6_Click);
            buttonItem9["�p��w��Ǧ~���Z(�¨�)"].Click += new System.EventHandler(this.buttonItem8_Click);
        }
        public override string ProcessTabName
        {
            get
            {
                return "���Z�B�z";
            }
        }
        /// <summary>
        /// �p��Ǵ��������Z
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonItem5_Click(object sender, EventArgs e)
        {
            new CalcSemesterEntryScoreWizard(SelectType.GradeYearStudent).ShowDialog();

        }
        /// <summary>
        /// �p��Ǧ~��ئ��Z
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void buttonItem4_Click(object sender, EventArgs e)
        {
            new CalcSemesterSubjectScoreWizard(SelectType.GradeYearStudent).ShowDialog();
        }

        private void buttonItem6_Click(object sender, EventArgs e)
        {
            new CalcSemesterMoralScoreWizard(SelectType.GradeYearStudent).ShowDialog();
        }

        private void buttonItem3_Click(object sender, EventArgs e)
        {
            new CalcSchoolYearEntryScoreWizard(SelectType.GradeYearStudent).ShowDialog();
        }

        private void buttonItem8_Click(object sender, EventArgs e)
        {
            new CalcSchoolYearMoralScoreWizard(SelectType.GradeYearStudent).ShowDialog();
        }

        private void buttonItem2_Click(object sender, EventArgs e)
        {
            new CalcSchoolYearSubjectScoreWizard(SelectType.GradeYearStudent).ShowDialog();
        }

        private void buttonItem2_1_Click(object sender, EventArgs e)
        {
            new CalcSchoolYearApplyScoreWizard(SelectType.GradeYearStudent).ShowDialog();
        }

        private void btnSemesterRank_Click(object sender, EventArgs e)
        {
            new SemesterRatingForm().ShowDialog();
        }

        private void btnSchoolYearRank_Click(object sender, EventArgs e)
        {
            new SchoolYearRatingForm().ShowDialog();
        }

        private void buttonItem6_Click_1(object sender, EventArgs e)
        {
            new CalsSemesterScoreWizard().ShowDialog();
        }

        private void buttonItem8_Click_1(object sender, EventArgs e)
        {
            new CalcSchoolYearScoreWizard().ShowDialog();
        }
    }
}

