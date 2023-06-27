using FISCA.Presentation;
using SmartSchool.Evaluation.Reports;
using System;

namespace SmartSchool.Evaluation.Process
{
    public partial class Retake : SmartSchool.Evaluation.Process.RibbonBarBase
    {
        public Retake()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Retake));
            var buttonItem52 = MotherForm.RibbonBarItems["�аȧ@�~", "�妸�@�~/�˵�"]["���Z�@�~"];
            buttonItem52.Image = Properties.Resources.calc_save_64;
            buttonItem52.Size = RibbonBarButton.MenuButtonSize.Large;
            buttonItem52["��ĳ���צW��-�̬��"].BeginGroup = true;
            buttonItem52["��ĳ���צW��-�̬��"].Click += new System.EventHandler(this.buttonItem1_Click);
            buttonItem52["��ĳ���צW��-�̾ǥ�"].Click += new System.EventHandler(this.buttonItem2_Click);
            buttonItem52["���צ��Z�פJ��"].Click += new System.EventHandler(this.buttonItem3_Click);

            buttonItem52.Enable = CurrentUser.Acl["Button0680"].Executable;
        }
        public override string ProcessTabName
        {
            get
            {
                return "���Z�B�z";
            }
        }
        #region ��ĳ���צW��-�̬�� (��ؤ��ή�W��)
        private void buttonItem1_Click(object sender, EventArgs e)
        {
            new RetakeListBySubject();
        }
        #endregion

        #region ��ĳ���צW��-�̾ǥ� (�ǥͭ��׬�ئW��)
        private void buttonItem2_Click(object sender, EventArgs e)
        {
            new RetakeListByStudent();
        }
        #endregion

        #region ���צ��Z�פJ��
        private void buttonItem3_Click(object sender, EventArgs e)
        {
            new RetakeScoreImport();
        }
        #endregion
    }
}

