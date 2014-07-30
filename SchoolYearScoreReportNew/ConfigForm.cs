using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SchoolYearScoreReport;
using Aspose.Words;
using System.IO;
using System.Diagnostics;

namespace SchoolYearScoreReport
{
    public partial class ConfigForm : FISCA.Presentation.Controls.BaseForm
    {
        private Config _config;
        
        public ConfigForm(Config config)
        {
            InitializeComponent();

            this.buttonX2.AccessibleRole = AccessibleRole.PushButton;
            this.buttonX1.AccessibleRole = AccessibleRole.PushButton;
            this._config = config;
            if (this.Config.CustomTemplate != null)
            {
                this.linkLabel2.Enabled = true;
                this.radioButton2.Enabled = true;
            }
            if (this.Config.UseDefault)
            {
                this.radioButton1.Checked = true;
                this.radioButton2.Checked = false;
            }
            else
            {
                this.radioButton1.Checked = false;
                this.radioButton2.Checked = true;
            }
            this.comboBoxEx1.SelectedIndex = this.Config.ReceiveNameIndex;
            this.comboBoxEx2.SelectedIndex = this.Config.ReceiveAddressIndex;
            this.textBoxX1.Text = this.Config.ResitSign;
            this.textBoxX2.Text = this.Config.RepeatSign;
            this.txtSYSa.Text = this.Config.YearResitSign;
            this.txtSYSb.Text = this.Config.YearRepeatSign;
            this.checkBoxX1.Checked = this.Config.AllowOver;
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            this.Config.UseDefault = this.radioButton1.Checked;
            this.Config.SetReceiveInfo(this.comboBoxEx1.SelectedIndex, this.comboBoxEx2.SelectedIndex);
            this.Config.SetSign(this.textBoxX1.Text, this.textBoxX2.Text,txtSYSa.Text,txtSYSb.Text);
            this.Config.AllowOver = this.checkBoxX1.Checked;
            this.Config.Save();
            base.DialogResult = DialogResult.OK;
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            base.Close();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(this.sfd.FileName, FileMode.Create);
                    fs.Write(this.Config.DefaultTemplate, 0, this.Config.DefaultTemplate.Length);
                    fs.Close();
                    Process.Start(this.sfd.FileName);
                }
                catch
                {
                    MessageBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.sfd.FileName = "自訂學年成績單範本.doc";
            if (this.sfd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    FileStream fs = new FileStream(this.sfd.FileName, FileMode.Create);
                    if ((this.Config.CustomTemplate != null) && (Document.DetectFileFormat(new MemoryStream(this.Config.CustomTemplate)) == LoadFormat.Doc))
                    {
                        fs.Write(this.Config.CustomTemplate, 0, this.Config.CustomTemplate.Length);
                    }
                    else
                    {
                        fs.Write(this.Config.DefaultTemplate, 0, this.Config.DefaultTemplate.Length);
                    }
                    fs.Close();
                    Process.Start(this.sfd.FileName);
                }
                catch
                {
                    MessageBox.Show("指定路徑無法存取。", "另存檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.ofd.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    if (Document.DetectFileFormat(this.ofd.FileName) == LoadFormat.Doc)
                    {
                        FileStream fs = new FileStream(this.ofd.FileName, FileMode.Open);
                        byte[] custom = new byte[fs.Length];
                        fs.Read(custom, 0, custom.Length);
                        this.Config.UploadCustomTemplate(custom);
                        this.linkLabel2.Enabled = true;
                        this.radioButton2.Enabled = true;
                        this.labelX1.Visible = true;
                        fs.Close();
                    }
                    else
                    {
                        MessageBox.Show("上傳檔案格式不符");
                    }
                }
                catch
                {
                    MessageBox.Show("指定路徑無法存取。", "開啟檔案失敗", MessageBoxButtons.OK, MessageBoxIcon.Hand);
                }
            }
        }
        // Properties
        private Config Config
        {
            get
            {
                return this._config;
            }
        }

        private void ConfigForm_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
        }

    }
}
