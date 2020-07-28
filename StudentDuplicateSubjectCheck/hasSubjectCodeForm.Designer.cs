namespace StudentDuplicateSubjectCheck
{
    partial class hasSubjectCodeForm
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.dgData = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colSchoolYear = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSemester = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colCourseName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colStudentName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colStudentNumber = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSubjCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btnAddTemp = new DevComponents.DotNetBar.ButtonX();
            this.btnExportList = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.lblMsg = new DevComponents.DotNetBar.LabelX();
            this.btnWrite = new DevComponents.DotNetBar.ButtonX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
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
            this.colSchoolYear,
            this.colSemester,
            this.colCourseName,
            this.colStudentName,
            this.colStudentNumber,
            this.colSubjCode});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgData.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgData.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgData.Location = new System.Drawing.Point(22, 41);
            this.dgData.Name = "dgData";
            this.dgData.RowTemplate.Height = 24;
            this.dgData.Size = new System.Drawing.Size(869, 300);
            this.dgData.TabIndex = 0;
            // 
            // colSchoolYear
            // 
            this.colSchoolYear.HeaderText = "學年度";
            this.colSchoolYear.Name = "colSchoolYear";
            this.colSchoolYear.ReadOnly = true;
            this.colSchoolYear.Width = 70;
            // 
            // colSemester
            // 
            this.colSemester.HeaderText = "學期";
            this.colSemester.Name = "colSemester";
            this.colSemester.ReadOnly = true;
            this.colSemester.Width = 60;
            // 
            // colCourseName
            // 
            this.colCourseName.HeaderText = "課程名稱";
            this.colCourseName.Name = "colCourseName";
            this.colCourseName.ReadOnly = true;
            this.colCourseName.Width = 150;
            // 
            // colStudentName
            // 
            this.colStudentName.HeaderText = "姓名";
            this.colStudentName.Name = "colStudentName";
            this.colStudentName.ReadOnly = true;
            // 
            // colStudentNumber
            // 
            this.colStudentNumber.HeaderText = "學號";
            this.colStudentNumber.Name = "colStudentNumber";
            this.colStudentNumber.ReadOnly = true;
            // 
            // colSubjCode
            // 
            this.colSubjCode.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.AllCells;
            this.colSubjCode.HeaderText = "課程代碼";
            this.colSubjCode.Name = "colSubjCode";
            this.colSubjCode.ReadOnly = true;
            this.colSubjCode.Width = 85;
            // 
            // btnAddTemp
            // 
            this.btnAddTemp.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAddTemp.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAddTemp.AutoSize = true;
            this.btnAddTemp.BackColor = System.Drawing.Color.Transparent;
            this.btnAddTemp.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAddTemp.Location = new System.Drawing.Point(22, 356);
            this.btnAddTemp.Name = "btnAddTemp";
            this.btnAddTemp.Size = new System.Drawing.Size(105, 25);
            this.btnAddTemp.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnAddTemp.TabIndex = 1;
            this.btnAddTemp.Text = "學生加入待處理";
            this.btnAddTemp.Click += new System.EventHandler(this.btnAddTemp_Click);
            // 
            // btnExportList
            // 
            this.btnExportList.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExportList.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnExportList.AutoSize = true;
            this.btnExportList.BackColor = System.Drawing.Color.Transparent;
            this.btnExportList.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExportList.Location = new System.Drawing.Point(141, 356);
            this.btnExportList.Name = "btnExportList";
            this.btnExportList.Size = new System.Drawing.Size(91, 25);
            this.btnExportList.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExportList.TabIndex = 2;
            this.btnExportList.Text = "匯出學生清單";
            this.btnExportList.Click += new System.EventHandler(this.btnExportList_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(816, 356);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 4;
            this.btnExit.Text = " 離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
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
            this.lblMsg.Location = new System.Drawing.Point(262, 357);
            this.lblMsg.Name = "lblMsg";
            this.lblMsg.Size = new System.Drawing.Size(459, 23);
            this.lblMsg.TabIndex = 4;
            this.lblMsg.Click += new System.EventHandler(this.lblMsg_Click);
            // 
            // btnWrite
            // 
            this.btnWrite.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnWrite.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnWrite.AutoSize = true;
            this.btnWrite.BackColor = System.Drawing.Color.Transparent;
            this.btnWrite.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnWrite.Location = new System.Drawing.Point(735, 356);
            this.btnWrite.Name = "btnWrite";
            this.btnWrite.Size = new System.Drawing.Size(75, 25);
            this.btnWrite.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnWrite.TabIndex = 3;
            this.btnWrite.Text = "覆蓋";
            this.btnWrite.Click += new System.EventHandler(this.btnWrite_Click);
            // 
            // labelX1
            // 
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(22, 13);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(869, 23);
            this.labelX1.TabIndex = 5;
            this.labelX1.Text = "覆蓋：依課程規劃表完整取代，離開：僅將空白的欄位依照課程規劃表填入。";
            // 
            // hasSubjectCodeForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(903, 393);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.btnWrite);
            this.Controls.Add(this.lblMsg);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnExportList);
            this.Controls.Add(this.btnAddTemp);
            this.Controls.Add(this.dgData);
            this.DoubleBuffered = true;
            this.Name = "hasSubjectCodeForm";
            this.Text = "已有課程代碼學生清單";
            this.Load += new System.EventHandler(this.HasPassScoreForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dgData)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dgData;
        private DevComponents.DotNetBar.ButtonX btnAddTemp;
        private DevComponents.DotNetBar.ButtonX btnExportList;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.DotNetBar.LabelX lblMsg;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSchoolYear;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSemester;
        private System.Windows.Forms.DataGridViewTextBoxColumn colCourseName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colStudentName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colStudentNumber;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSubjCode;
        private DevComponents.DotNetBar.ButtonX btnWrite;
        private DevComponents.DotNetBar.LabelX labelX1;
    }
}