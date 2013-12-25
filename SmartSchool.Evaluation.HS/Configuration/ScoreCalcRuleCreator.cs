using System;
using System.ComponentModel;
using System.Xml;
using DevComponents.Editors;
using SmartSchool.Evaluation.ScoreCalcRule;
using SmartSchool.Feature.ScoreCalcRule;

namespace SmartSchool.Evaluation.Configuration
{
    public partial class ScoreCalcRuleCreator : FISCA.Presentation.Controls.BaseForm
    {
        private BackgroundWorker _BGWScoreCalcRuleLoader;

        private XmlElement _copyContent;

        public ScoreCalcRuleCreator()
        {
            InitializeComponent();

            //_copyElement = new XmlDocument().CreateElement("ScoreCalcRule");
            //_copyElement.AppendChild(new XmlDocument().CreateElement("Name"));
            //_copyElement.AppendChild(new XmlDocument().CreateElement("Content"));
            comboItem1.Tag = _copyContent;
            comboBoxEx1.SelectedIndex = 0;

            _BGWScoreCalcRuleLoader = new BackgroundWorker();
            _BGWScoreCalcRuleLoader.DoWork += new DoWorkEventHandler(_BGWScoreCalcRuleLoader_DoWork);
            _BGWScoreCalcRuleLoader.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_BGWScoreCalcRuleLoader_RunWorkerCompleted);
            _BGWScoreCalcRuleLoader.RunWorkerAsync();
        }

        void _BGWScoreCalcRuleLoader_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            comboBoxEx1.Items.Remove(comboItem2);

            ScoreCalcRuleInfoCollection resp = (ScoreCalcRuleInfoCollection)e.Result;
            foreach (ScoreCalcRuleInfo scr in resp)
            {
                ComboItem item = new ComboItem();
                item.Text = scr.Name;
                item.Tag = scr.ScoreCalcRuleElement;
                comboBoxEx1.Items.Add(item);
            }
        }

        void _BGWScoreCalcRuleLoader_DoWork(object sender, DoWorkEventArgs e)
        {
            e.Result = ScoreCalcRule.ScoreCalcRule.Instance.Items;
        }

        private void button_cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button_save_Click(object sender, EventArgs e)
        {
            if (textBoxX1.Text != "")
            {
                if (comboBoxEx1.SelectedIndex == 0)
                    AddScoreCalcRule.Insert(textBoxX1.Text);
                else
                {
                    AddScoreCalcRule.Insert(textBoxX1.Text, _copyContent);
                }
                ScoreCalcRule.ScoreCalcRule.Instance.Invok_ScoreCalcRuleInserted();
                this.Close();
            }
            else
                this.Close();
        }

        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (comboBoxEx1.SelectedItem == comboItem2)
                comboBoxEx1.SelectedIndex = 0;
            else
            {
                _copyContent = ((comboBoxEx1.SelectedItem as ComboItem).Tag as XmlElement);
            }
        }
    }
}