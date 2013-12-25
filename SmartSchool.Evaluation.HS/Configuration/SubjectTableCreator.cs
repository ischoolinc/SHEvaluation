using System;
using System.Xml;
using SmartSchool.Common;
using SmartSchool.ExceptionHandler;

namespace SmartSchool.Evaluation.Configuration
{
    public partial class SubjectTableCreator : FISCA.Presentation.Controls.BaseForm
    {
        private string _Catalog;

        private XmlElement _CopyElement;

        public SubjectTableCreator(string catalog)
        {
            this._Catalog=catalog;
            InitializeComponent();

            this.Text = "新增" + catalog;

            comboBoxEx1.SelectedItem = comboItem1;
            foreach ( SubjectTableItem var in SubjectTable.Items[_Catalog])
            {
                comboBoxEx1.Items.Add(var) ;   
            }

            textBoxX1.Focus();
        }

        private void textBoxX1_TextChanged(object sender, EventArgs e)
        {
            foreach ( SubjectTableItem var in SubjectTable.Items[_Catalog] )
            {
                if ( var.Name == textBoxX1.Text )
                {
                    errorProvider1.SetError(textBoxX1, "名稱不可重複。");
                    //MsgBox.Show("名稱不可重複。");
                    return;
                }
                else
                    errorProvider1.Clear();
            }
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            if ( textBoxX1.Text != "" )
            {
                foreach ( SubjectTableItem var in SubjectTable.Items[_Catalog] )
                {
                    if ( var.Name == textBoxX1.Text )
                    {
                        MsgBox.Show("名稱不可重複。");
                        return;
                    }
                }
                try
                {
                    SmartSchool.Feature.SubjectTable.AddSubejctTable.Insert(textBoxX1.Text, _Catalog, _CopyElement);
                }
                catch(Exception ex)
                {
                    CurrentUser user = CurrentUser.Instance;
                    BugReporter.ReportException("SmartSchool", user.SystemVersion, ex, false);
                    MsgBox.Show("新增" + _Catalog + "時發生未預期之錯誤。\n系統已回報此錯誤內容。");
                }
                SubjectTable.Items[_Catalog].Reflash();
                this.Close();
            }
            else
            {
                MsgBox.Show("必需輸入" + _Catalog + "名稱。");
            }
        }

        private void comboBoxEx1_SelectedIndexChanged(object sender, EventArgs e)
        {
            if ( comboBoxEx1.SelectedItem == comboItem1 )
                _CopyElement = new XmlDocument().CreateElement("SubjectTableContent");
            else
            {
                _CopyElement = (XmlElement)( (SubjectTableItem)comboBoxEx1.SelectedItem ).Content.SelectSingleNode("SubjectTableContent");
            }
        }
    }
}