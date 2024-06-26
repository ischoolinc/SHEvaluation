namespace SmartSchool.Evaluation.Content.ChangeSchoolYear
{
    partial class frmChangeSemesterSchoolYearHasData
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lblMsg = new DevComponents.DotNetBar.LabelX();
            this.lnkHasDataList = new System.Windows.Forms.LinkLabel();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.SuspendLayout();
            // 
            // lblMsg
            // 
            this.lblMsg.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblMsg.BackgroundStyle.Class = "";
            this.lblMsg.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblMsg.Location = new System.Drawing.Point(12, 22);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(332, 45);
            this.lblMsg.TabIndex = 0;
            this.lblMsg.Text = "將107學年度第1學期成績年級1年級 已有重複科目名稱+級別，請確認後再進行作業。";
            this.lblMsg.TextLineAlignment = System.Drawing.StringAlignment.Near;
            this.lblMsg.WordWrap = true;
            // 
            // lnkHasDataList
            // 
            this.lnkHasDataList.AutoSize = true;
            this.lnkHasDataList.BackColor = System.Drawing.Color.Transparent;
            this.lnkHasDataList.Location = new System.Drawing.Point(12, 85);
            this.lnkHasDataList.Name = "lnkHasDataList";
            this.lnkHasDataList.Size = new System.Drawing.Size(86, 17);
            this.lnkHasDataList.TabIndex = 1;
            this.lnkHasDataList.TabStop = true;
            this.lnkHasDataList.Text = "查看重複清單";
            this.lnkHasDataList.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkHasDataList_LinkClicked);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(244, 85);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 2;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // frmChangeSemesterSchoolYearHasData
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(355, 128);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.lnkHasDataList);
            this.Controls.Add(this.lblMsg);
            this.DoubleBuffered = true;
            this.Name = "frmChangeSemesterSchoolYearHasData";
            this.Text = "變更學年度學期";
            this.Load += new System.EventHandler(this.frmChangeSemesterSchoolYearHasData_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX lblMsg;
        private System.Windows.Forms.LinkLabel lnkHasDataList;
        private DevComponents.DotNetBar.ButtonX btnExit;
    }
}