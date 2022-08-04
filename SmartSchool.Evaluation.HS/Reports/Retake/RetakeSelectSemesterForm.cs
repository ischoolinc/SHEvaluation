﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SmartSchool.Evaluation.Reports.Retake
{
    public partial class RetakeSelectSemesterForm : SelectSemesterForm
    {
        public bool IsPrintAllSemester
        {
            get { return chkAllSemester.Checked; }
        }

        //public bool IsPrintAllGradeYear
        //{
        //    get { return chkAllYear.Checked; }
        //}

        //public int GradeYear
        //{
        //    get { return (int)gradeYearInput.Value; }
        //}

        public RetakeSelectSemesterForm(string titleName)
        {
            InitializeComponent();
            Text = titleName;
        }

        private void chkAllSemester_CheckedChanged(object sender, EventArgs e)
        {
            numericUpDown1.Enabled = !chkAllSemester.Checked;
            numericUpDown2.Enabled = !chkAllSemester.Checked;
        }

        //private void checkBoxX1_CheckedChanged(object sender, EventArgs e)
        //{
        //    gradeYearInput.Enabled = !chkAllYear.Checked;
        //}
    }
}