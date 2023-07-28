namespace StudentDuplicateSubjectCheck
{
    partial class hasSubjectCreditDiffForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgData = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.lblMsg = new DevComponents.DotNetBar.LabelX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.btnExportList = new DevComponents.DotNetBar.ButtonX();
            this.colGplanName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colGradeYear = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSubjectName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCreditA = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCreditB = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSpecifySubjectName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).BeginInit();
            this.SuspendLayout();
            // 
            // dgData
            // 
            this.dgData.AllowUserToAddRows = false;
            this.dgData.AllowUserToDeleteRows = false;
            this.dgData.AllowUserToResizeRows = false;
            this.dgData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgData.BackgroundColor = System.Drawing.Color.White;
            this.dgData.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgData.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colGplanName,
            this.colGradeYear,
            this.colSubjectName,
            this.colCreditA,
            this.colCreditB,
            this.colSpecifySubjectName});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgData.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgData.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgData.Location = new System.Drawing.Point(12, 40);
            this.dgData.Name = "dgData";
            this.dgData.RowTemplate.Height = 24;
            this.dgData.Size = new System.Drawing.Size(837, 300);
            this.dgData.TabIndex = 6;
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(12, 12);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(794, 23);
            this.labelX1.TabIndex = 12;
            this.labelX1.Text = "課程規劃表科目同年級學分數不同，需要到課程規劃表設定\"指定學年科目名稱\"。";
            // 
            // lblMsg
            // 
            this.lblMsg.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.lblMsg.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.lblMsg.BackgroundStyle.Class = "";
            this.lblMsg.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lblMsg.Location = new System.Drawing.Point(124, 356);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(427, 23);
            this.lblMsg.TabIndex = 10;
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(774, 355);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 11;
            this.btnExit.Text = " 離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // btnExportList
            // 
            this.btnExportList.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExportList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnExportList.AutoSize = true;
            this.btnExportList.BackColor = System.Drawing.Color.Transparent;
            this.btnExportList.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExportList.Location = new System.Drawing.Point(12, 355);
            this.btnExportList.Name = "btnExportList";
            this.btnExportList.Size = new System.Drawing.Size(91, 25);
            this.btnExportList.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExportList.TabIndex = 8;
            this.btnExportList.Text = "匯出清單";
            this.btnExportList.Click += new System.EventHandler(this.btnExportList_Click);
            // 
            // colGplanName
            // 
            this.colGplanName.HeaderText = "課程規劃名稱";
            this.colGplanName.Name = "colGplanName";
            this.colGplanName.ReadOnly = true;
            this.colGplanName.Width = 250;
            // 
            // colGradeYear
            // 
            this.colGradeYear.HeaderText = "年級";
            this.colGradeYear.Name = "colGradeYear";
            this.colGradeYear.ReadOnly = true;
            this.colGradeYear.Width = 60;
            // 
            // colSubjectName
            // 
            this.colSubjectName.HeaderText = "科目名稱";
            this.colSubjectName.Name = "colSubjectName";
            this.colSubjectName.ReadOnly = true;
            this.colSubjectName.Width = 150;
            // 
            // colCreditA
            // 
            this.colCreditA.HeaderText = "上學期學分數";
            this.colCreditA.Name = "colCreditA";
            this.colCreditA.ReadOnly = true;
            this.colCreditA.Width = 80;
            // 
            // colCreditB
            // 
            this.colCreditB.HeaderText = "下學期學分數";
            this.colCreditB.Name = "colCreditB";
            this.colCreditB.ReadOnly = true;
            this.colCreditB.Width = 80;
            // 
            // colSpecifySubjectName
            // 
            this.colSpecifySubjectName.HeaderText = "指定學年科目名稱";
            this.colSpecifySubjectName.Name = "colSpecifySubjectName";
            this.colSpecifySubjectName.ReadOnly = true;
            this.colSpecifySubjectName.Width = 150;
            // 
            // hasSubjectCreditDiffForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(871, 393);
            this.Controls.Add(this.dgData);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnExportList);
            this.DoubleBuffered = true;
            this.Name = "hasSubjectCreditDiffForm";
            this.Text = "課程規劃表科目同年級學分數不同";
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dgData;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX lblMsg;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.ButtonX btnExportList;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGplanName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colGradeYear;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSubjectName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCreditA;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCreditB;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSpecifySubjectName;
    }
}