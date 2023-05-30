namespace SmartSchool.Evaluation.Configuration
{
    partial class GraduationPlanSimplePicker
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
            this.components = new System.ComponentModel.Container();
            System.Windows.Forms.DataGridViewCellStyle dataGridViewCellStyle1 = new System.Windows.Forms.DataGridViewCellStyle();
            this.panelEx1 = new DevComponents.DotNetBar.PanelEx();
            this.cboEntry = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.emptyEntry = new DevComponents.Editors.ComboItem();
            this.dgGraduationPlan = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.colDomain = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colEntry = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colSubjName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colRequiredBy = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colRequired = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.colLevelList = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.cboRequired = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.emptyRequired = new DevComponents.Editors.ComboItem();
            this.labelX7 = new DevComponents.DotNetBar.LabelX();
            this.cboRequiredBy = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.emptyRequiredBy = new DevComponents.Editors.ComboItem();
            this.labelX6 = new DevComponents.DotNetBar.LabelX();
            this.cboDomain = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.emptyDomain = new DevComponents.Editors.ComboItem();
            this.iiSchoolYear = new DevComponents.Editors.IntegerInput();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.cboGPlan = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.buttonX2 = new DevComponents.DotNetBar.ButtonX();
            this.btnSelect = new DevComponents.DotNetBar.ButtonX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX4 = new DevComponents.DotNetBar.LabelX();
            this.errorProvider1 = new System.Windows.Forms.ErrorProvider(this.components);
            this.panelEx1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgGraduationPlan)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iiSchoolYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).BeginInit();
            this.SuspendLayout();
            // 
            // panelEx1
            // 
            this.panelEx1.CanvasColor = System.Drawing.SystemColors.Control;
            this.panelEx1.ColorSchemeStyle = DevComponents.DotNetBar.eDotNetBarStyle.Office2007;
            this.panelEx1.Controls.Add(this.cboEntry);
            this.panelEx1.Controls.Add(this.dgGraduationPlan);
            this.panelEx1.Controls.Add(this.cboRequired);
            this.panelEx1.Controls.Add(this.labelX7);
            this.panelEx1.Controls.Add(this.cboRequiredBy);
            this.panelEx1.Controls.Add(this.labelX6);
            this.panelEx1.Controls.Add(this.cboDomain);
            this.panelEx1.Controls.Add(this.iiSchoolYear);
            this.panelEx1.Controls.Add(this.labelX3);
            this.panelEx1.Controls.Add(this.cboGPlan);
            this.panelEx1.Controls.Add(this.buttonX2);
            this.panelEx1.Controls.Add(this.btnSelect);
            this.panelEx1.Controls.Add(this.labelX2);
            this.panelEx1.Controls.Add(this.labelX1);
            this.panelEx1.Controls.Add(this.labelX4);
            this.panelEx1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panelEx1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            this.panelEx1.Location = new System.Drawing.Point(0, 0);
            this.panelEx1.Name = "panelEx1";
            this.panelEx1.Size = new System.Drawing.Size(629, 452);
            this.panelEx1.Style.Alignment = System.Drawing.StringAlignment.Center;
            this.panelEx1.Style.BackColor1.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground;
            this.panelEx1.Style.BackColor2.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBackground2;
            this.panelEx1.Style.Border = DevComponents.DotNetBar.eBorderType.SingleLine;
            this.panelEx1.Style.BorderColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelBorder;
            this.panelEx1.Style.ForeColor.ColorSchemePart = DevComponents.DotNetBar.eColorSchemePart.PanelText;
            this.panelEx1.Style.GradientAngle = 90;
            this.panelEx1.TabIndex = 0;
            // 
            // cboEntry
            // 
            this.cboEntry.DisplayMember = "Text";
            this.cboEntry.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboEntry.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboEntry.FormattingEnabled = true;
            this.cboEntry.ItemHeight = 19;
            this.cboEntry.Items.AddRange(new object[] {
            this.emptyEntry});
            this.cboEntry.Location = new System.Drawing.Point(281, 39);
            this.cboEntry.Name = "cboEntry";
            this.cboEntry.Size = new System.Drawing.Size(103, 25);
            this.cboEntry.TabIndex = 17;
            this.cboEntry.SelectedIndexChanged += new System.EventHandler(this.cbo_SelectedIndexChanged);
            // 
            // dgGraduationPlan
            // 
            this.dgGraduationPlan.AllowUserToAddRows = false;
            this.dgGraduationPlan.AllowUserToDeleteRows = false;
            this.dgGraduationPlan.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dgGraduationPlan.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.colDomain,
            this.colEntry,
            this.colSubjName,
            this.colRequiredBy,
            this.colRequired,
            this.colLevelList});
            dataGridViewCellStyle1.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle1.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle1.Font = new System.Drawing.Font("Microsoft JhengHei", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle1.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle1.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle1.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dgGraduationPlan.DefaultCellStyle = dataGridViewCellStyle1;
            this.dgGraduationPlan.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dgGraduationPlan.Location = new System.Drawing.Point(3, 70);
            this.dgGraduationPlan.Name = "dgGraduationPlan";
            this.dgGraduationPlan.RowTemplate.Height = 24;
            this.dgGraduationPlan.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
            this.dgGraduationPlan.Size = new System.Drawing.Size(618, 342);
            this.dgGraduationPlan.TabIndex = 16;
            // 
            // colDomain
            // 
            this.colDomain.HeaderText = "領域";
            this.colDomain.Name = "colDomain";
            this.colDomain.ReadOnly = true;
            this.colDomain.Width = 80;
            // 
            // colEntry
            // 
            this.colEntry.HeaderText = "分項";
            this.colEntry.Name = "colEntry";
            this.colEntry.ReadOnly = true;
            this.colEntry.Width = 80;
            // 
            // colSubjName
            // 
            this.colSubjName.HeaderText = "科目名稱";
            this.colSubjName.Name = "colSubjName";
            this.colSubjName.ReadOnly = true;
            this.colSubjName.Width = 180;
            // 
            // colRequiredBy
            // 
            this.colRequiredBy.HeaderText = "校部定";
            this.colRequiredBy.Name = "colRequiredBy";
            this.colRequiredBy.ReadOnly = true;
            this.colRequiredBy.Width = 70;
            // 
            // colRequired
            // 
            this.colRequired.HeaderText = "必選修";
            this.colRequired.Name = "colRequired";
            this.colRequired.ReadOnly = true;
            this.colRequired.Width = 70;
            // 
            // colLevelList
            // 
            this.colLevelList.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.colLevelList.HeaderText = "級別";
            this.colLevelList.Name = "colLevelList";
            this.colLevelList.ReadOnly = true;
            // 
            // cboRequired
            // 
            this.cboRequired.DisplayMember = "Text";
            this.cboRequired.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboRequired.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboRequired.FormattingEnabled = true;
            this.cboRequired.ItemHeight = 19;
            this.cboRequired.Items.AddRange(new object[] {
            this.emptyRequired});
            this.cboRequired.Location = new System.Drawing.Point(557, 39);
            this.cboRequired.Name = "cboRequired";
            this.cboRequired.Size = new System.Drawing.Size(61, 25);
            this.cboRequired.TabIndex = 15;
            this.cboRequired.SelectedIndexChanged += new System.EventHandler(this.cbo_SelectedIndexChanged);
            // 
            // labelX7
            // 
            this.labelX7.AutoSize = true;
            // 
            // 
            // 
            this.labelX7.BackgroundStyle.Class = "";
            this.labelX7.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX7.Enabled = false;
            this.labelX7.Location = new System.Drawing.Point(512, 41);
            this.labelX7.Name = "labelX7";
            this.labelX7.Size = new System.Drawing.Size(50, 21);
            this.labelX7.TabIndex = 14;
            this.labelX7.Text = "必選修:";
            // 
            // cboRequiredBy
            // 
            this.cboRequiredBy.DisplayMember = "Text";
            this.cboRequiredBy.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboRequiredBy.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboRequiredBy.FormattingEnabled = true;
            this.cboRequiredBy.ItemHeight = 19;
            this.cboRequiredBy.Items.AddRange(new object[] {
            this.emptyRequiredBy});
            this.cboRequiredBy.Location = new System.Drawing.Point(436, 39);
            this.cboRequiredBy.Name = "cboRequiredBy";
            this.cboRequiredBy.Size = new System.Drawing.Size(67, 25);
            this.cboRequiredBy.TabIndex = 13;
            this.cboRequiredBy.SelectedIndexChanged += new System.EventHandler(this.cbo_SelectedIndexChanged);
            // 
            // labelX6
            // 
            this.labelX6.AutoSize = true;
            // 
            // 
            // 
            this.labelX6.BackgroundStyle.Class = "";
            this.labelX6.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX6.Enabled = false;
            this.labelX6.Location = new System.Drawing.Point(390, 41);
            this.labelX6.Name = "labelX6";
            this.labelX6.Size = new System.Drawing.Size(50, 21);
            this.labelX6.TabIndex = 12;
            this.labelX6.Text = "校部定:";
            // 
            // cboDomain
            // 
            this.cboDomain.DisplayMember = "Text";
            this.cboDomain.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboDomain.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboDomain.FormattingEnabled = true;
            this.cboDomain.ItemHeight = 19;
            this.cboDomain.Items.AddRange(new object[] {
            this.emptyDomain});
            this.cboDomain.Location = new System.Drawing.Point(52, 39);
            this.cboDomain.Name = "cboDomain";
            this.cboDomain.Size = new System.Drawing.Size(184, 25);
            this.cboDomain.TabIndex = 7;
            this.cboDomain.SelectedIndexChanged += new System.EventHandler(this.cbo_SelectedIndexChanged);
            // 
            // iiSchoolYear
            // 
            // 
            // 
            // 
            this.iiSchoolYear.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iiSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iiSchoolYear.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iiSchoolYear.Location = new System.Drawing.Point(54, 7);
            this.iiSchoolYear.MinValue = 0;
            this.iiSchoolYear.Name = "iiSchoolYear";
            this.iiSchoolYear.ShowUpDown = true;
            this.iiSchoolYear.Size = new System.Drawing.Size(55, 25);
            this.iiSchoolYear.TabIndex = 4;
            this.iiSchoolYear.ValueChanged += new System.EventHandler(this.iiSchoolYear_ValueChanged);
            // 
            // labelX3
            // 
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(3, 8);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(74, 23);
            this.labelX3.TabIndex = 3;
            this.labelX3.Text = "學年度：";
            // 
            // cboGPlan
            // 
            this.cboGPlan.DisplayMember = "Text";
            this.cboGPlan.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboGPlan.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboGPlan.FormattingEnabled = true;
            this.cboGPlan.ItemHeight = 19;
            this.cboGPlan.Location = new System.Drawing.Point(191, 7);
            this.cboGPlan.Name = "cboGPlan";
            this.cboGPlan.Size = new System.Drawing.Size(427, 25);
            this.cboGPlan.TabIndex = 1;
            this.cboGPlan.SelectedIndexChanged += new System.EventHandler(this.cboGPlan_SelectedIndexChanged);
            // 
            // buttonX2
            // 
            this.buttonX2.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.buttonX2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.buttonX2.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.buttonX2.Location = new System.Drawing.Point(558, 419);
            this.buttonX2.Name = "buttonX2";
            this.buttonX2.Size = new System.Drawing.Size(59, 23);
            this.buttonX2.TabIndex = 2;
            this.buttonX2.Text = "取消";
            this.buttonX2.Click += new System.EventHandler(this.buttonX2_Click);
            // 
            // btnSelect
            // 
            this.btnSelect.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSelect.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSelect.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSelect.Location = new System.Drawing.Point(493, 419);
            this.btnSelect.Name = "btnSelect";
            this.btnSelect.Size = new System.Drawing.Size(59, 23);
            this.btnSelect.TabIndex = 0;
            this.btnSelect.Text = "選擇";
            this.btnSelect.Click += new System.EventHandler(this.buttonX1_Click);
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Enabled = false;
            this.labelX2.Location = new System.Drawing.Point(115, 9);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(87, 21);
            this.labelX2.TabIndex = 0;
            this.labelX2.Text = "課程規劃表：";
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Enabled = false;
            this.labelX1.Location = new System.Drawing.Point(12, 41);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(47, 21);
            this.labelX1.TabIndex = 6;
            this.labelX1.Text = "領域：";
            // 
            // labelX4
            // 
            this.labelX4.AutoSize = true;
            // 
            // 
            // 
            this.labelX4.BackgroundStyle.Class = "";
            this.labelX4.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX4.Enabled = false;
            this.labelX4.Location = new System.Drawing.Point(242, 41);
            this.labelX4.Name = "labelX4";
            this.labelX4.Size = new System.Drawing.Size(47, 21);
            this.labelX4.TabIndex = 8;
            this.labelX4.Text = "分項：";
            // 
            // errorProvider1
            // 
            this.errorProvider1.ContainerControl = this;
            // 
            // GraduationPlanSimplePicker
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.ClientSize = new System.Drawing.Size(629, 452);
            this.Controls.Add(this.panelEx1);
            this.DoubleBuffered = true;
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Name = "GraduationPlanSimplePicker";
            this.Text = "從課程規劃表選擇科目";
            this.panelEx1.ResumeLayout(false);
            this.panelEx1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dgGraduationPlan)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iiSchoolYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.errorProvider1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.PanelEx panelEx1;
        private DevComponents.DotNetBar.ButtonX buttonX2;
        private DevComponents.DotNetBar.ButtonX btnSelect;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboGPlan;
        private DevComponents.Editors.IntegerInput iiSchoolYear;
        private DevComponents.DotNetBar.LabelX labelX3;
        private System.Windows.Forms.ErrorProvider errorProvider1;
        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboDomain;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboRequired;
        private DevComponents.DotNetBar.LabelX labelX7;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboRequiredBy;
        private DevComponents.DotNetBar.Controls.DataGridViewX dgGraduationPlan;
        private DevComponents.DotNetBar.LabelX labelX6;
        private DevComponents.DotNetBar.LabelX labelX4;
        private DevComponents.Editors.ComboItem emptyDomain;
        private DevComponents.Editors.ComboItem emptyRequired;
        private DevComponents.Editors.ComboItem emptyRequiredBy;
        private System.Windows.Forms.DataGridViewTextBoxColumn colDomain;
        private System.Windows.Forms.DataGridViewTextBoxColumn colEntry;
        private System.Windows.Forms.DataGridViewTextBoxColumn colSubjName;
        private System.Windows.Forms.DataGridViewTextBoxColumn colRequiredBy;
        private System.Windows.Forms.DataGridViewTextBoxColumn colRequired;
        private System.Windows.Forms.DataGridViewTextBoxColumn colLevelList;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboEntry;
        private DevComponents.Editors.ComboItem emptyEntry;
    }
}