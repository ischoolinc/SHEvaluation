namespace SHImportExportScoreRank.UIForm
{
    partial class frmStudentRankDetail
    {
        private System.ComponentModel.IContainer components = null;

        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        private void InitializeComponent()
        {
            this.grpHeader = new System.Windows.Forms.GroupBox();
            this.lblBatchNameValue = new System.Windows.Forms.Label();
            this.lblCreateTimeValue = new System.Windows.Forms.Label();
            this.lblCreateTypeValue = new System.Windows.Forms.Label();
            this.lblScoreItemValue = new System.Windows.Forms.Label();
            this.lblScoreTypeValue = new System.Windows.Forms.Label();
            this.lblGradeYearValue = new System.Windows.Forms.Label();
            this.lblSchoolYearValue = new System.Windows.Forms.Label();
            this.lblBatchName = new System.Windows.Forms.Label();
            this.lblCreateTime = new System.Windows.Forms.Label();
            this.lblCreateType = new System.Windows.Forms.Label();
            this.lblScoreItem = new System.Windows.Forms.Label();
            this.lblScoreType = new System.Windows.Forms.Label();
            this.lblGradeYear = new System.Windows.Forms.Label();
            this.lblSchoolYear = new System.Windows.Forms.Label();
            this.dgvRankDetail = new System.Windows.Forms.DataGridView();
            this.btnSave = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.grpHeader.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRankDetail)).BeginInit();
            this.SuspendLayout();
            // 
            // grpHeader
            // 
            this.grpHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.grpHeader.BackColor = System.Drawing.Color.Transparent;
            this.grpHeader.Controls.Add(this.lblBatchNameValue);
            this.grpHeader.Controls.Add(this.lblCreateTimeValue);
            this.grpHeader.Controls.Add(this.lblCreateTypeValue);
            this.grpHeader.Controls.Add(this.lblScoreItemValue);
            this.grpHeader.Controls.Add(this.lblScoreTypeValue);
            this.grpHeader.Controls.Add(this.lblGradeYearValue);
            this.grpHeader.Controls.Add(this.lblSchoolYearValue);
            this.grpHeader.Controls.Add(this.lblBatchName);
            this.grpHeader.Controls.Add(this.lblCreateTime);
            this.grpHeader.Controls.Add(this.lblCreateType);
            this.grpHeader.Controls.Add(this.lblScoreItem);
            this.grpHeader.Controls.Add(this.lblScoreType);
            this.grpHeader.Controls.Add(this.lblGradeYear);
            this.grpHeader.Controls.Add(this.lblSchoolYear);
            this.grpHeader.Location = new System.Drawing.Point(12, 12);
            this.grpHeader.Name = "grpHeader";
            this.grpHeader.Size = new System.Drawing.Size(860, 120);
            this.grpHeader.TabIndex = 0;
            this.grpHeader.TabStop = false;
            // 
            // lblBatchNameValue
            // 
            this.lblBatchNameValue.AutoSize = true;
            this.lblBatchNameValue.Location = new System.Drawing.Point(548, 66);
            this.lblBatchNameValue.Name = "lblBatchNameValue";
            this.lblBatchNameValue.Size = new System.Drawing.Size(0, 17);
            this.lblBatchNameValue.TabIndex = 11;
            // 
            // lblCreateTimeValue
            // 
            this.lblCreateTimeValue.AutoSize = true;
            this.lblCreateTimeValue.Location = new System.Drawing.Point(548, 42);
            this.lblCreateTimeValue.Name = "lblCreateTimeValue";
            this.lblCreateTimeValue.Size = new System.Drawing.Size(0, 17);
            this.lblCreateTimeValue.TabIndex = 10;
            // 
            // lblCreateTypeValue
            // 
            this.lblCreateTypeValue.AutoSize = true;
            this.lblCreateTypeValue.Location = new System.Drawing.Point(548, 18);
            this.lblCreateTypeValue.Name = "lblCreateTypeValue";
            this.lblCreateTypeValue.Size = new System.Drawing.Size(0, 17);
            this.lblCreateTypeValue.TabIndex = 9;
            // 
            // lblScoreItemValue
            // 
            this.lblScoreItemValue.AutoSize = true;
            this.lblScoreItemValue.Location = new System.Drawing.Point(308, 90);
            this.lblScoreItemValue.Name = "lblScoreItemValue";
            this.lblScoreItemValue.Size = new System.Drawing.Size(0, 17);
            this.lblScoreItemValue.TabIndex = 8;
            // 
            // lblScoreTypeValue
            // 
            this.lblScoreTypeValue.AutoSize = true;
            this.lblScoreTypeValue.Location = new System.Drawing.Point(308, 66);
            this.lblScoreTypeValue.Name = "lblScoreTypeValue";
            this.lblScoreTypeValue.Size = new System.Drawing.Size(0, 17);
            this.lblScoreTypeValue.TabIndex = 7;
            // 
            // lblGradeYearValue
            // 
            this.lblGradeYearValue.AutoSize = true;
            this.lblGradeYearValue.Location = new System.Drawing.Point(308, 42);
            this.lblGradeYearValue.Name = "lblGradeYearValue";
            this.lblGradeYearValue.Size = new System.Drawing.Size(0, 17);
            this.lblGradeYearValue.TabIndex = 12;
            // 
            // lblSchoolYearValue
            // 
            this.lblSchoolYearValue.AutoSize = true;
            this.lblSchoolYearValue.Location = new System.Drawing.Point(308, 18);
            this.lblSchoolYearValue.Name = "lblSchoolYearValue";
            this.lblSchoolYearValue.Size = new System.Drawing.Size(0, 17);
            this.lblSchoolYearValue.TabIndex = 6;
            // 
            // lblBatchName
            // 
            this.lblBatchName.AutoSize = true;
            this.lblBatchName.Location = new System.Drawing.Point(472, 66);
            this.lblBatchName.Name = "lblBatchName";
            this.lblBatchName.Size = new System.Drawing.Size(60, 17);
            this.lblBatchName.TabIndex = 5;
            this.lblBatchName.Text = "排名批次";
            // 
            // lblCreateTime
            // 
            this.lblCreateTime.AutoSize = true;
            this.lblCreateTime.Location = new System.Drawing.Point(472, 42);
            this.lblCreateTime.Name = "lblCreateTime";
            this.lblCreateTime.Size = new System.Drawing.Size(60, 17);
            this.lblCreateTime.TabIndex = 4;
            this.lblCreateTime.Text = "建立時間";
            // 
            // lblCreateType
            // 
            this.lblCreateType.AutoSize = true;
            this.lblCreateType.Location = new System.Drawing.Point(472, 18);
            this.lblCreateType.Name = "lblCreateType";
            this.lblCreateType.Size = new System.Drawing.Size(60, 17);
            this.lblCreateType.TabIndex = 3;
            this.lblCreateType.Text = "建立方式";
            // 
            // lblScoreItem
            // 
            this.lblScoreItem.AutoSize = true;
            this.lblScoreItem.Location = new System.Drawing.Point(232, 90);
            this.lblScoreItem.Name = "lblScoreItem";
            this.lblScoreItem.Size = new System.Drawing.Size(60, 17);
            this.lblScoreItem.TabIndex = 2;
            this.lblScoreItem.Text = "成績項目";
            // 
            // lblScoreType
            // 
            this.lblScoreType.AutoSize = true;
            this.lblScoreType.Location = new System.Drawing.Point(232, 66);
            this.lblScoreType.Name = "lblScoreType";
            this.lblScoreType.Size = new System.Drawing.Size(60, 17);
            this.lblScoreType.TabIndex = 1;
            this.lblScoreType.Text = "成績類型";
            // 
            // lblGradeYear
            // 
            this.lblGradeYear.AutoSize = true;
            this.lblGradeYear.Location = new System.Drawing.Point(232, 42);
            this.lblGradeYear.Name = "lblGradeYear";
            this.lblGradeYear.Size = new System.Drawing.Size(60, 17);
            this.lblGradeYear.TabIndex = 13;
            this.lblGradeYear.Text = "成績年級";
            // 
            // lblSchoolYear
            // 
            this.lblSchoolYear.AutoSize = true;
            this.lblSchoolYear.Location = new System.Drawing.Point(232, 18);
            this.lblSchoolYear.Name = "lblSchoolYear";
            this.lblSchoolYear.Size = new System.Drawing.Size(47, 17);
            this.lblSchoolYear.TabIndex = 0;
            this.lblSchoolYear.Text = "學年度";
            // 
            // dgvRankDetail
            // 
            this.dgvRankDetail.AllowUserToAddRows = false;
            this.dgvRankDetail.AllowUserToDeleteRows = false;
            this.dgvRankDetail.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvRankDetail.AutoSizeColumnsMode = System.Windows.Forms.DataGridViewAutoSizeColumnsMode.Fill;
            this.dgvRankDetail.BackgroundColor = System.Drawing.Color.White;
            this.dgvRankDetail.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvRankDetail.Location = new System.Drawing.Point(12, 138);
            this.dgvRankDetail.MultiSelect = false;
            this.dgvRankDetail.Name = "dgvRankDetail";
            this.dgvRankDetail.ReadOnly = true;
            this.dgvRankDetail.RowHeadersVisible = false;
            this.dgvRankDetail.RowTemplate.Height = 24;
            this.dgvRankDetail.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgvRankDetail.Size = new System.Drawing.Size(860, 278);
            this.dgvRankDetail.TabIndex = 1;
            // 
            // btnSave
            // 
            this.btnSave.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSave.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSave.BackColor = System.Drawing.Color.Transparent;
            this.btnSave.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSave.Enabled = false;
            this.btnSave.Location = new System.Drawing.Point(716, 428);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(75, 23);
            this.btnSave.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSave.TabIndex = 2;
            this.btnSave.Text = "儲存";
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(797, 428);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 23);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 3;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // frmStudentRankDetail
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(884, 463);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.dgvRankDetail);
            this.Controls.Add(this.grpHeader);
            this.DoubleBuffered = true;
            this.MinimumSize = new System.Drawing.Size(900, 500);
            this.Name = "frmStudentRankDetail";
            this.Text = "學生排名資料";
            this.Load += new System.EventHandler(this.frmStudentRankDetail_Load);
            this.grpHeader.ResumeLayout(false);
            this.grpHeader.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvRankDetail)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox grpHeader;
        private System.Windows.Forms.Label lblSchoolYear;
        private System.Windows.Forms.Label lblGradeYear;
        private System.Windows.Forms.Label lblScoreType;
        private System.Windows.Forms.Label lblScoreItem;
        private System.Windows.Forms.Label lblCreateType;
        private System.Windows.Forms.Label lblCreateTime;
        private System.Windows.Forms.Label lblBatchName;
        private System.Windows.Forms.Label lblSchoolYearValue;
        private System.Windows.Forms.Label lblGradeYearValue;
        private System.Windows.Forms.Label lblScoreTypeValue;
        private System.Windows.Forms.Label lblScoreItemValue;
        private System.Windows.Forms.Label lblCreateTypeValue;
        private System.Windows.Forms.Label lblCreateTimeValue;
        private System.Windows.Forms.Label lblBatchNameValue;
        private System.Windows.Forms.DataGridView dgvRankDetail;
        private DevComponents.DotNetBar.ButtonX btnSave;
        private DevComponents.DotNetBar.ButtonX btnExit;
    }
}
