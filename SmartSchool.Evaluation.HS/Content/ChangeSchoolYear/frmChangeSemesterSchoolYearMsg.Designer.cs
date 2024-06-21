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
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
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
            this.labelSourceMsg.Location = new System.Drawing.Point(149, 19);
            this.labelSourceMsg.Name = "labelSourceMsg";
            this.labelSourceMsg.Size = new System.Drawing.Size(215, 23);
            this.labelSourceMsg.TabIndex = 4;
            this.labelSourceMsg.Text = "112學年度第1學期成績年級1年級";
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(9, 19);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(139, 23);
            this.labelX1.TabIndex = 3;
            this.labelX1.Text = "欲調整之學期成績為：";
            // 
            // frmChangeSemesterSchoolYearMsg
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(373, 151);
            this.Controls.Add(this.labelSourceMsg);
            this.Controls.Add(this.labelX1);
            this.DoubleBuffered = true;
            this.Name = "frmChangeSemesterSchoolYearMsg";
            this.Text = "變更學年度學期";
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelSourceMsg;
        private DevComponents.DotNetBar.LabelX labelX1;
    }
}