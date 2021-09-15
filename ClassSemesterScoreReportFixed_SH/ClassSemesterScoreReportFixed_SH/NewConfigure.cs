using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace ClassSemesterScoreReportFixed_SH
{
    public partial class NewConfigure : FISCA.Presentation.Controls.BaseForm
    {
        public Aspose.Words.Document Template { get; private set; }
        public int SubjectLimit { get; private set; }
        public int StudentLimit { get; private set; }
        public string ConfigName { get; private set; }

        public NewConfigure()
        {
            InitializeComponent();
            checkBoxX1.CheckedChanged += new EventHandler(SetupTemplate);
            checkBoxX2.CheckedChanged += new EventHandler(UploadTemplate);
        }
        private void SetupTemplate(object sender, EventArgs e)
        {
            if (checkBoxX1.Checked)
            {
                this.SubjectLimit = 25;
                this.StudentLimit = 55;
                try
                {
                    Template = new Aspose.Words.Document(new MemoryStream(Properties.Resources.高中班級學期成績單樣版));                    
                    List<string> fields = new List<string>(Template.MailMerge.GetFieldNames());
                    this.SubjectLimit = 0;
                    while (fields.Contains("科目名稱" + (this.SubjectLimit + 1)))
                    {
                        this.SubjectLimit++;
                    }
                    this.StudentLimit = 0;
                    while (fields.Contains("姓名" + (this.StudentLimit + 1)))
                    {
                        this.StudentLimit++;
                    }
                }
                catch
                {
                    MessageBox.Show("樣板讀取失敗");
                 
                }                
            }
        }
        private void UploadTemplate(object sender, EventArgs e)
        {
            if (checkBoxX2.Checked)
            {
                OpenFileDialog dialog = new OpenFileDialog();
                dialog.Title = "上傳樣板";
                dialog.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        Template = new Aspose.Words.Document(dialog.FileName);
                        List<string> fields = new List<string>(Template.MailMerge.GetFieldNames());
                        this.SubjectLimit = 0;
                        while (fields.Contains("科目名稱" + (this.SubjectLimit + 1)))
                        {
                            this.SubjectLimit++;
                        }
                        this.StudentLimit = 0;
                        while (fields.Contains("姓名" + (this.StudentLimit + 1)))
                        {
                            this.StudentLimit++;
                        }
                    }
                    catch
                    {
                        MessageBox.Show("樣板開啟失敗");
                        checkBoxX2.Checked = false;
                    }
                }
                else
                    checkBoxX2.Checked = false;
            }
        }

        private void checkReady(object sender, EventArgs e)
        {
            bool ready = true;
            if (txtName.Text == "")
                ready = false;
            else
                ConfigName = txtName.Text;
            if (!checkBoxX1.Checked && !checkBoxX2.Checked)
            {
                ready = false;
            }
            btnSubmit.Enabled = ready;
        }





        private void btnSubmit_Click(object sender, EventArgs e)
        {
            if (string.IsNullOrWhiteSpace(txtName.Text))
                return;

            // 檢查是否是預設設定檔名稱，如果是無法新增
            if (Global.DefaultConfigNameList().Contains(txtName.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("樣板名稱與預設設定檔名稱相同，無法新增!");
                return;
            }
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            Close();
        }

        private void lnkViewMapping_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkViewMapping.Enabled = false;
            Program.CreateFieldTemplate();
            lnkViewMapping.Enabled = true;
        }

        private void lnkViewTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkViewTemplate.Enabled = false;
            #region 儲存檔案
            string inputReportName = "班級學期成績單(固定排名)樣板.doc";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName);

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                //document.Save(path, Aspose.Words.SaveFormat.Doc);
                System.IO.FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                stream.Write(Properties.Resources.高中班級學期成績單樣版, 0, Properties.Resources.高中班級學期成績單樣版.Length);
                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName;
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        //document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.高中班級學期成績單樣版, 0, Properties.Resources.高中班級學期成績單樣版.Length);
                        stream.Flush();
                        stream.Close();

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            lnkViewTemplate.Enabled = true;
            #endregion
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
