using SmartSchool.Common;
using System;
using System.Windows.Forms;
using System.Xml;

namespace SmartSchool.Evaluation.Reports
{
    public partial class SemesterMoralScoreTotalConfig : BaseForm
    {
        public SemesterMoralScoreTotalConfig(bool over100, int size)
        {
            InitializeComponent();

            comboBoxEx1.SelectedIndex = size;
            checkBoxX1.Checked = over100;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
            this.Close();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            #region �x�s Preference

            XmlElement config = CurrentUser.Instance.Preference["�w�榨�Z�`��"];

            if (config == null)
            {
                config = new XmlDocument().CreateElement("�w�榨�Z�`��");
            }

            XmlElement print = config.OwnerDocument.CreateElement("Print");
            print.SetAttribute("AllowMoralScoreOver100", checkBoxX1.Checked.ToString());
            print.SetAttribute("PaperSize", comboBoxEx1.SelectedIndex.ToString());

            if (config.SelectSingleNode("Print") == null)
                config.AppendChild(print);
            else
                config.ReplaceChild(print, config.SelectSingleNode("Print"));

            CurrentUser.Instance.Preference["�w�榨�Z�`��"] = config;

            #endregion

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}