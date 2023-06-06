using Aspose.Words;
using DevComponents.DotNetBar.Rendering;
using SmartSchool.Common;
using System;
using System.Collections.Generic;
using System.IO;
using System.Windows.Forms;

namespace SmartSchool.Evaluation.Reports.MultiSemesterScore.Forms
{
    public partial class TemplateSetupForm : BaseForm
    {
        private ReportOptions _option;
        private ReportOptions Option { get { return _option; } }

        private byte[] temp_buffer = null;

        public TemplateSetupForm(ReportOptions option)
        {
            InitializeComponent();

            #region �p�G�t�Ϊ�Renderer�OOffice2007Renderer�A�P��_ClassTeacherView,_CategoryView���C��
            if (GlobalManager.Renderer is Office2007Renderer)
            {
                ((Office2007Renderer)GlobalManager.Renderer).ColorTableChanged += new EventHandler(ScoreCalcRuleEditor_ColorTableChanged);
                SetForeColor(this);
            }
            #endregion

            _option = option;
            Load();
        }
        void ScoreCalcRuleEditor_ColorTableChanged(object sender, EventArgs e)
        {
            SetForeColor(this);
        }

        private void SetForeColor(Control parent)
        {
            foreach (Control var in parent.Controls)
            {
                if (var is CheckBox)
                    var.ForeColor = ((Office2007Renderer)GlobalManager.Renderer).ColorTable.CheckBoxItem.Default.Text;
                SetForeColor(var);
            }
        }

        private void Load()
        {
            if (Option.IsDefaultTemplate)
                radioButton1.Checked = true;
            else
                radioButton2.Checked = true;

            foreach (CheckBox cbox in new CheckBox[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 })
            {
                cbox.Checked = Option.PrintEntries.Contains(cbox.Text);
            }
            checkBox7.Checked = !Option.FixMoralScore;

            integerInput1.Value = Option.PrintSemester;
        }

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton1.Checked)
                radioButton2.Checked = false;
        }

        private void radioButton2_CheckedChanged(object sender, EventArgs e)
        {
            if (radioButton2.Checked)
                radioButton1.Checked = false;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "�t�s�s��";
            sfd.FileName = "�h�Ǵ����Z��.doc";
            sfd.Filter = "Word�ɮ� (*.doc)|*.doc|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    fs.Write(Option.DefaultTemplate, 0, Option.DefaultTemplate.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            SaveFileDialog sfd = new SaveFileDialog();
            sfd.Title = "�t�s�s��";
            sfd.FileName = "�ۭq�h�Ǵ����Z��.doc";
            sfd.Filter = "Word�ɮ� (*.doc)|*.doc|�Ҧ��ɮ� (*.*)|*.*";
            if (sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(sfd.FileName, FileMode.Create);
                    if (Aspose.Words.Document.DetectFileFormat(new MemoryStream(Option.CustomizeTemplate)) == Aspose.Words.LoadFormat.Doc)
                        fs.Write(Option.CustomizeTemplate, 0, Option.CustomizeTemplate.Length);
                    fs.Close();
                    System.Diagnostics.Process.Start(sfd.FileName);
                }
                catch
                {
                    MsgBox.Show("���w���|�L�k�s���C", "�t�s�ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "��ܦۭq���h�Ǵ����Z��d��";
            ofd.Filter = "Word�ɮ� (*.doc)|*.doc";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (Document.DetectFileFormat(ofd.FileName) == LoadFormat.Doc)
                    {
                        FileStream fs = new FileStream(ofd.FileName, FileMode.Open);

                        temp_buffer = new byte[fs.Length];
                        fs.Read(temp_buffer, 0, temp_buffer.Length);
                        fs.Close();
                        MsgBox.Show("�W�Ǧ��\�C");
                    }
                    else
                        MsgBox.Show("�W���ɮ׮榡����");
                }
                catch
                {
                    MsgBox.Show("���w���|�L�k�s���C", "�}���ɮץ���", MessageBoxButtons.OK, MessageBoxIcon.Error);
                    return;
                }
            }
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            Option.CustomizeTemplate = temp_buffer;
            Option.IsDefaultTemplate = radioButton1.Checked;
            Option.PrintSemester = integerInput1.Value;

            List<string> list = new List<string>();
            foreach (CheckBox cbox in new CheckBox[] { checkBox1, checkBox2, checkBox3, checkBox4, checkBox5, checkBox6 })
            {
                if (cbox.Checked)
                {
                    list.Add(cbox.Text);
                }
            }
            Option.PrintEntries = list;
            Option.FixMoralScore = !checkBox7.Checked;
            this.DialogResult = DialogResult.OK;
            this.Close();
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}