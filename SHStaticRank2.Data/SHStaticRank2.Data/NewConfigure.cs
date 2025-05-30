﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace SHStaticRank2.Data
{
    public partial class NewConfigure : FISCA.Presentation.Controls.BaseForm
    {
        public Aspose.Words.Document Template { get; private set; }
        public int SubjectLimit { get; private set; }
        public string ConfigName { get; private set; }
        public bool WithSchoolYearScore { get; private set; }
        public bool WithPrevSemesterScore { get; private set; }
        private List<string> _NameList = new List<string>();

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
                Template = new Aspose.Words.Document(new MemoryStream(Properties.Resources.高中多學期學生成績證明書));
                this.SubjectLimit = 26;
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


                        //row["學期科目原始成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("原始成績");
                        //row["學期科目補考成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("補考成績");
                        //row["學期科目重修成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("重修成績");
                        //row["學期科目手動調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("擇優採計成績");
                        //row["學期科目學年調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("學年調整成績");
                        //row["學期科目成績" + subjectIndex] = semesterSubjectScore.Score;

                        //row["上學期科目原始成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("原始成績");
                        //row["上學期科目補考成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("補考成績");
                        //row["上學期科目重修成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("重修成績");
                        //row["上學期科目手動調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("擇優採計成績");
                        //row["上學期科目學年調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("學年調整成績");
                        //row["上學期科目成績" + subjectIndex] = semesterSubjectScore.Score;

                        //row["學年科目成績" + subjectIndex] = schoolYearSubjectScore.Score;
             
                        Template = new Aspose.Words.Document(dialog.FileName);
                        List<string> fields = new List<string>(Template.MailMerge.GetFieldNames());
                        this.SubjectLimit = 0;
                        while (fields.Contains("科目名稱" + (this.SubjectLimit + 1)))
                        {
                            if (fields.Contains("上學期科目原始成績" + (this.SubjectLimit + 1))) WithPrevSemesterScore = true;
                            if (fields.Contains("上學期科目補考成績" + (this.SubjectLimit + 1))) WithPrevSemesterScore = true;
                            if (fields.Contains("上學期科目重修成績" + (this.SubjectLimit + 1))) WithPrevSemesterScore = true;
                            if (fields.Contains("上學期科目手動調整成績" + (this.SubjectLimit + 1))) WithPrevSemesterScore = true;
                            if (fields.Contains("上學期科目學年調整成績" + (this.SubjectLimit + 1))) WithPrevSemesterScore = true;
                            if (fields.Contains("上學期科目成績" + (this.SubjectLimit + 1))) WithPrevSemesterScore = true;
                            if (fields.Contains("學年科目成績" + (this.SubjectLimit + 1))) WithSchoolYearScore = true;
                            this.SubjectLimit++;
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

        public void SetNameList(List<string> nameList)
        {
            _NameList = nameList;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 儲存檔案
            string inputReportName = "多學期成績單樣板(高中).doc";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

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
                stream.Write(Properties.Resources.高中多學期學生成績證明書, 0, Properties.Resources.高中多學期學生成績證明書.Length);
                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Excel檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        //document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.高中多學期學生成績證明書, 0, Properties.Resources.高中多學期學生成績證明書.Length);
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
            #endregion
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            #region 儲存檔案
            string inputReportName = "多學期成績單合併欄位總表.doc";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

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
                stream.Write(Properties.Resources.合併欄位總表, 0, Properties.Resources.合併欄位總表.Length);
                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        //document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.合併欄位總表, 0, Properties.Resources.合併欄位總表.Length);
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
            #endregion
        }

        private void btnSubmit_Click(object sender, EventArgs e)
        {
            if (_NameList.Contains(txtName.Text))
            {
                FISCA.Presentation.Controls.MsgBox.Show("已有相同設定檔名稱無法新增!");
                return;
            }

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            Close();
        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.Close();
        }

        private void NewConfigure_Load(object sender, EventArgs e)
        {

        }
    }
}
