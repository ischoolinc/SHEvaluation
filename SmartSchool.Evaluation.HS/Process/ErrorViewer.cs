using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using DevComponents.DotNetBar;
using SmartSchool.Customization.Data;

namespace SmartSchool.Evaluation.Process
{
    public partial class ErrorViewer : SmartSchool.Common.BaseForm
    {
        public ErrorViewer()
        {
            InitializeComponent();
        }
        public void SetMessage(StudentRecord stu, List<string> messages)
        {
            bool first = true;
            foreach (string msg in messages)
            {
                string s="";
                if (first)
                {
                    if (stu.RefClass != null)
                    {
                        s += stu.RefClass.ClassName;
                        if (stu.SeatNo != "")
                            s += "(" + stu.SeatNo + "號)";
                        s += " ";
                    }
                    if (stu.StudentNumber != "")
                        s += stu.StudentNumber + " ";
                    if (s == "")
                        s += "學生：";
                    s += stu.StudentName;
                    dataGridViewX1.Rows.Add(s, msg);
                    first = false;
                }
                else
                    dataGridViewX1.Rows.Add("", msg);
            }
            toolStripStatusLabel1.Text ="總計"+dataGridViewX1.Rows.Count+"個錯誤。";
        }

        public void Clear()
        {
            this.Hide();
            this.dataGridViewX1.Rows.Clear();
            toolStripStatusLabel1.Text = "";
        }

        private void ErrorViewer_FormClosing(object sender, FormClosingEventArgs e)
        {
            this.Hide();
            e.Cancel = true;
        }
    }
}