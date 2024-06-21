namespace SmartSchool.Evaluation.Content.ChangeSchoolYear
{
    partial class frmChangeSemesterSchoolYearMsg
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
            this.labelSourceMsg = new DevComponents.DotNetBar.LabelX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.btnChange = new DevComponents.DotNetBar.ButtonX();
            this.labelChangeMsg = new DevComponents.DotNetBar.LabelX();
            this.SuspendLayout();
            // 
            // labelSourceMsg
            // 
            this.labelSourceMsg.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelSourceMsg.BackgroundStyle.Class = "";
            this.labelSourceMsg.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelSourceMsg.Location = new System.Drawing.Point(12, 12);
            this.labelSourceMsg.Name = "labelSourceMsg";
            this.labelSourceMsg.Size = new System.Drawing.Size(232, 23);
            this.labelSourceMsg.TabIndex = 4;
            this.labelSourceMsg.Text = "將112學年度第1學期成績年級1年級";
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(193, 90);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 6;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnChange
            // 
            this.btnChange.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnChange.BackColor = System.Drawing.Color.Transparent;
            this.btnChange.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnChange.Location = new System.Drawing.Point(98, 90);
            this.btnChange.Name = "btnChange";
            this.btnChange.Size = new System.Drawing.Size(75, 23);
            this.btnChange.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnChange.TabIndex = 5;
            this.btnChange.Text = "確認";
            this.btnChange.Click += new System.EventHandler(this.btnChange_Click);
            // 
            // labelChangeMsg
            // 
            this.labelChangeMsg.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelChangeMsg.BackgroundStyle.Class = "";
            this.labelChangeMsg.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelChangeMsg.Location = new System.Drawing.Point(12, 43);
            this.labelChangeMsg.Name = "labelChangeMsg";
            this.labelChangeMsg.Size = new System.Drawing.Size(270, 23);
            this.labelChangeMsg.TabIndex = 7;
            this.labelChangeMsg.Text = "調整為112學年度第1學期成績年級1年級";
            // 
            // frmChangeSemesterSchoolYearMsg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(288, 123);
            this.Controls.Add(this.labelChangeMsg);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnChange);
            this.Controls.Add(this.labelSourceMsg);
            this.DoubleBuffered = true;
            this.Name = "frmChangeSemesterSchoolYearMsg";
            this.Text = "變更學年度學期";
            this.Load += new System.EventHandler(this.frmChangeSemesterSchoolYearMsg_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelSourceMsg;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.ButtonX btnChange;
        private DevComponents.DotNetBar.LabelX labelChangeMsg;
    }
}