namespace SmartSchool.Evaluation.Content.ScoreEditor
{
    partial class frmPickSubject
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
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle2 = new System.Windows.Forms.DataGridViewCellStyle();
            this.gpMainData = new DevComponents.DotNetBar.Controls.GroupPanel();
            this.dgvMain = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colMainSubjectName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMainDomainName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMainEntry = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMainRequiredBy = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMainRequired = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMain1_1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMain1_2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMain2_1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMain2_2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMain3_1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMain3_2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.不計學分 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.不需評分 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMainOfficialSubjectName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colMainSchoolYearGroupName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.課程代碼 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.labelGPName = new DevComponents.DotNetBar.LabelX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.btnAddSubject = new DevComponents.DotNetBar.ButtonX();
            this.gpMainData.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgvMain)).BeginInit();
            this.SuspendLayout();
            // 
            // gpMainData
            // 
            this.gpMainData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.gpMainData.BackColor = System.Drawing.Color.Transparent;
            this.gpMainData.CanvasColor = System.Drawing.SystemColors.Control;
            this.gpMainData.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.gpMainData.Controls.Add(this.dgvMain);
            this.gpMainData.Location = new System.Drawing.Point(-2, 42);
            this.gpMainData.Name = "gpMainData";
            this.gpMainData.Size = new System.Drawing.Size(1209, 600);
            // 
            // 
            // 
            this.gpMainData.Style.BackColor2SchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.gpMainData.Style.BackColorGradientAngle = 90;
            this.gpMainData.Style.BackColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.gpMainData.Style.BorderBottom = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpMainData.Style.BorderBottomWidth = 1;
            this.gpMainData.Style.BorderColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.gpMainData.Style.BorderLeft = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpMainData.Style.BorderLeftWidth = 1;
            this.gpMainData.Style.BorderRight = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpMainData.Style.BorderRightWidth = 1;
            this.gpMainData.Style.BorderTop = DevComponents.DotNetBar.eStyleBorderType.Solid;
            this.gpMainData.Style.BorderTopWidth = 1;
            this.gpMainData.Style.Class = "";
            this.gpMainData.Style.CornerDiameter = 4;
            this.gpMainData.Style.CornerType = DevComponents.DotNetBar.eCornerType.Rounded;
            this.gpMainData.Style.TextColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.gpMainData.Style.TextLineAlignment = DevComponents.DotNetBar.eStyleTextAlignment.Near;
            // 
            // 
            // 
            this.gpMainData.StyleMouseDown.Class = "";
            this.gpMainData.StyleMouseDown.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            // 
            // 
            // 
            this.gpMainData.StyleMouseOver.Class = "";
            this.gpMainData.StyleMouseOver.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.gpMainData.TabIndex = 8;
            this.gpMainData.Text = "課程規畫表內容";
            // 
            // dgvMain
            // 
            this.dgvMain.AllowUserToAddRows = false;
            this.dgvMain.AllowUserToDeleteRows = false;
            this.dgvMain.AllowUserToOrderColumns = true;
            this.dgvMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dgvMain.BackgroundColor = System.Drawing.Color.White;
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleCenter;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Control;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.WindowText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.HighlightText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.True;
            this.dgvMain.ColumnHeadersDefaultCellStyle = dataGridViewCellStyle1;
            this.dgvMain.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgvMain.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colMainSubjectName,
            this.colMainDomainName,
            this.colMainEntry,
            this.colMainRequiredBy,
            this.colMainRequired,
            this.colMain1_1,
            this.colMain1_2,
            this.colMain2_1,
            this.colMain2_2,
            this.colMain3_1,
            this.colMain3_2,
            this.不計學分,
            this.不需評分,
            this.colMainOfficialSubjectName,
            this.colMainSchoolYearGroupName,
            this.課程代碼});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgvMain.DefaultCellStyle = dataGridViewCellStyle2;
            this.dgvMain.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgvMain.Location = new System.Drawing.Point(0, 0);
            this.dgvMain.MultiSelect = false;
            this.dgvMain.Name = "dgvMain";
            this.dgvMain.ReadOnly = true;
            this.dgvMain.RowTemplate.Height = 24;
            this.dgvMain.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
            this.dgvMain.Size = new System.Drawing.Size(1203, 573);
            this.dgvMain.TabIndex = 6;
            this.dgvMain.CellClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dgvMain_CellClick);
            // 
            // colMainSubjectName
            // 
            this.colMainSubjectName.HeaderText = "科目名稱";
            this.colMainSubjectName.MinimumWidth = 60;
            this.colMainSubjectName.Name = "colMainSubjectName";
            this.colMainSubjectName.ReadOnly = true;
            this.colMainSubjectName.Width = 85;
            // 
            // colMainDomainName
            // 
            this.colMainDomainName.HeaderText = "領域名稱";
            this.colMainDomainName.MinimumWidth = 25;
            this.colMainDomainName.Name = "colMainDomainName";
            this.colMainDomainName.ReadOnly = true;
            this.colMainDomainName.Width = 67;
            // 
            // colMainEntry
            // 
            this.colMainEntry.HeaderText = "分項類別";
            this.colMainEntry.MinimumWidth = 25;
            this.colMainEntry.Name = "colMainEntry";
            this.colMainEntry.ReadOnly = true;
            this.colMainEntry.Width = 60;
            // 
            // colMainRequiredBy
            // 
            this.colMainRequiredBy.HeaderText = "校部定";
            this.colMainRequiredBy.MinimumWidth = 25;
            this.colMainRequiredBy.Name = "colMainRequiredBy";
            this.colMainRequiredBy.ReadOnly = true;
            this.colMainRequiredBy.Width = 75;
            // 
            // colMainRequired
            // 
            this.colMainRequired.HeaderText = "必選修";
            this.colMainRequired.MinimumWidth = 25;
            this.colMainRequired.Name = "colMainRequired";
            this.colMainRequired.ReadOnly = true;
            this.colMainRequired.Width = 75;
            // 
            // colMain1_1
            // 
            this.colMain1_1.HeaderText = "一上";
            this.colMain1_1.MinimumWidth = 50;
            this.colMain1_1.Name = "colMain1_1";
            this.colMain1_1.ReadOnly = true;
            this.colMain1_1.Width = 60;
            // 
            // colMain1_2
            // 
            this.colMain1_2.HeaderText = "一下";
            this.colMain1_2.MinimumWidth = 50;
            this.colMain1_2.Name = "colMain1_2";
            this.colMain1_2.ReadOnly = true;
            this.colMain1_2.Width = 60;
            // 
            // colMain2_1
            // 
            this.colMain2_1.HeaderText = "二上";
            this.colMain2_1.MinimumWidth = 50;
            this.colMain2_1.Name = "colMain2_1";
            this.colMain2_1.ReadOnly = true;
            this.colMain2_1.Width = 60;
            // 
            // colMain2_2
            // 
            this.colMain2_2.HeaderText = "二下";
            this.colMain2_2.MinimumWidth = 50;
            this.colMain2_2.Name = "colMain2_2";
            this.colMain2_2.ReadOnly = true;
            this.colMain2_2.Width = 60;
            // 
            // colMain3_1
            // 
            this.colMain3_1.HeaderText = "三上";
            this.colMain3_1.MinimumWidth = 50;
            this.colMain3_1.Name = "colMain3_1";
            this.colMain3_1.ReadOnly = true;
            this.colMain3_1.Width = 60;
            // 
            // colMain3_2
            // 
            this.colMain3_2.HeaderText = "三下";
            this.colMain3_2.MinimumWidth = 50;
            this.colMain3_2.Name = "colMain3_2";
            this.colMain3_2.ReadOnly = true;
            this.colMain3_2.Width = 60;
            // 
            // 不計學分
            // 
            this.不計學分.HeaderText = "不計學分";
            this.不計學分.Name = "不計學分";
            this.不計學分.ReadOnly = true;
            this.不計學分.Width = 60;
            // 
            // 不需評分
            // 
            this.不需評分.HeaderText = "不需評分";
            this.不需評分.MinimumWidth = 50;
            this.不需評分.Name = "不需評分";
            this.不需評分.ReadOnly = true;
            this.不需評分.Width = 60;
            // 
            // colMainOfficialSubjectName
            // 
            this.colMainOfficialSubjectName.HeaderText = "報部科目名稱";
            this.colMainOfficialSubjectName.MinimumWidth = 25;
            this.colMainOfficialSubjectName.Name = "colMainOfficialSubjectName";
            this.colMainOfficialSubjectName.ReadOnly = true;
            this.colMainOfficialSubjectName.Width = 80;
            // 
            // colMainSchoolYearGroupName
            // 
            this.colMainSchoolYearGroupName.HeaderText = "指定學年科目名稱";
            this.colMainSchoolYearGroupName.MinimumWidth = 100;
            this.colMainSchoolYearGroupName.Name = "colMainSchoolYearGroupName";
            this.colMainSchoolYearGroupName.ReadOnly = true;
            // 
            // 課程代碼
            // 
            this.課程代碼.HeaderText = "課程代碼";
            this.課程代碼.Name = "課程代碼";
            this.課程代碼.ReadOnly = true;
            this.課程代碼.Width = 120;
            // 
            // labelGPName
            // 
            this.labelGPName.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelGPName.BackgroundStyle.Class = "";
            this.labelGPName.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelGPName.Font = new System.Drawing.Font("Microsoft JhengHei", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.labelGPName.Location = new System.Drawing.Point(12, 12);
            this.labelGPName.Name = "labelGPName";
            this.labelGPName.Size = new System.Drawing.Size(504, 23);
            this.labelGPName.TabIndex = 9;
            this.labelGPName.Text = "課規名稱";
            // 
            // labelX1
            // 
            this.labelX1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.ForeColor = System.Drawing.Color.Red;
            this.labelX1.Location = new System.Drawing.Point(1016, 22);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(166, 23);
            this.labelX1.TabIndex = 10;
            this.labelX1.Text = "請點選學分數進行選取。";
            // 
            // btnAddSubject
            // 
            this.btnAddSubject.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAddSubject.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.btnAddSubject.BackColor = System.Drawing.Color.Transparent;
            this.btnAddSubject.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAddSubject.Location = new System.Drawing.Point(921, 22);
            this.btnAddSubject.Name = "btnAddSubject";
            this.btnAddSubject.Size = new System.Drawing.Size(75, 23);
            this.btnAddSubject.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnAddSubject.TabIndex = 11;
            this.btnAddSubject.Text = "加入科目";
            this.btnAddSubject.Click += new System.EventHandler(this.btnAddSubject_Click);
            // 
            // frmPickSubject
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1207, 639);
            this.Controls.Add(this.btnAddSubject);
            this.Controls.Add(this.labelX1);
            this.Controls.Add(this.labelGPName);
            this.Controls.Add(this.gpMainData);
            this.DoubleBuffered = true;
            this.Name = "frmPickSubject";
            this.Text = "選取科目";
            this.Load += new System.EventHandler(this.frmPickSubject_Load);
            this.gpMainData.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dgvMain)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.Controls.GroupPanel gpMainData;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgvMain;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMainSubjectName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMainDomainName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMainEntry;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMainRequiredBy;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMainRequired;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMain1_1;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMain1_2;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMain2_1;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMain2_2;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMain3_1;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMain3_2;
        private System.Windows.Forms.DataGridViewTextBoxColumn 不計學分;
        private System.Windows.Forms.DataGridViewTextBoxColumn 不需評分;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMainOfficialSubjectName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colMainSchoolYearGroupName;
        private System.Windows.Forms.DataGridViewTextBoxColumn 課程代碼;
        private DevComponents.DotNetBar.LabelX labelGPName;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.ButtonX btnAddSubject;
    }
}