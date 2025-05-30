using SmartSchool.Common;
using System;
using System.Windows.Forms;
using System.Xml;

namespace SmartSchool.Evaluation.Reports
{
    public partial class SemesterMoralScoreCalcConfig : BaseForm
    {
        public SemesterMoralScoreCalcConfig(bool over100, int size)
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
            #region 儲存 Preference

            XmlElement config = CurrentUser.Instance.Preference["德行成績試算表"];

            if (config == null)
            {
                config = new XmlDocument().CreateElement("德行成績試算表");
            }

            XmlElement print = config.OwnerDocument.CreateElement("Print");
            print.SetAttribute("AllowMoralScoreOver100", checkBoxX1.Checked.ToString());
            print.SetAttribute("PaperSize", comboBoxEx1.SelectedIndex.ToString());

            if (config.SelectSingleNode("Print") == null)
                config.AppendChild(print);
            else
                config.ReplaceChild(print, config.SelectSingleNode("Print"));

            CurrentUser.Instance.Preference["德行成績試算表"] = config;

            #endregion

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}