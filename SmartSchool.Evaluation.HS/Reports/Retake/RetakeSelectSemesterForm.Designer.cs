namespace SmartSchool.Evaluation.Reports.Retake
{
    partial class RetakeSelectSemesterForm
    {
        /// <summary>
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該公開 Managed 資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form 設計工具產生的程式碼

        /// <summary>
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.chkAllSemester = new DevComponents.DotNetBar.Controls.CheckBoxX();
            this.chkAllYear = new DevComponents.DotNetBar.Controls.CheckBoxX();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gradeYearInput)).BeginInit();
            this.SuspendLayout();
            // 
            // buttonX1
            // 
            this.buttonX1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonX1.Location = new System.Drawing.Point(310, 81);
            // 
            // chkAllSemester
            // 
            this.chkAllSemester.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkAllSemester.AutoSize = true;
            this.chkAllSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.chkAllSemester.BackgroundStyle.Class = "";
            this.chkAllSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkAllSemester.Location = new System.Drawing.Point(6, 81);
            this.chkAllSemester.Name = "chkAllSemester";
            this.chkAllSemester.Size = new System.Drawing.Size(107, 21);
            this.chkAllSemester.TabIndex = 4;
            this.chkAllSemester.Text = "列印全部學期";
            this.chkAllSemester.CheckedChanged += new System.EventHandler(this.chkAllSemester_CheckedChanged);
            // 
            // chkAllYear
            // 
            this.chkAllYear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.chkAllYear.AutoSize = true;
            this.chkAllYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.chkAllYear.BackgroundStyle.Class = "";
            this.chkAllYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.chkAllYear.Location = new System.Drawing.Point(129, 81);
            this.chkAllYear.Name = "chkAllYear";
            this.chkAllYear.Size = new System.Drawing.Size(107, 21);
            this.chkAllYear.TabIndex = 5;
            this.chkAllYear.Text = "列印全部年級";
            this.chkAllYear.CheckedChanged += new System.EventHandler(this.checkBoxX1_CheckedChanged);
            // 
            // RetakeSelectSemesterForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(397, 110);
            this.Controls.Add(this.chkAllYear);
            this.Controls.Add(this.chkAllSemester);
            this.Name = "RetakeSelectSemesterForm";
            this.Text = "RetakeSelectSemesterForm";
            this.Controls.SetChildIndex(this.gradeYearInput, 0);
            this.Controls.SetChildIndex(this.chkAllSemester, 0);
            this.Controls.SetChildIndex(this.chkAllYear, 0);
            this.Controls.SetChildIndex(this.buttonX1, 0);
            this.Controls.SetChildIndex(this.numericUpDown1, 0);
            this.Controls.SetChildIndex(this.numericUpDown2, 0);
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.numericUpDown2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gradeYearInput)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.CheckBoxX chkAllSemester;
        private DevComponents.DotNetBar.Controls.CheckBoxX chkAllYear;
    }
}