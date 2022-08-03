
namespace SH_ClassYearScoreReport
{
    partial class MainForm
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
            this.cboConfigure = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.linkLabel4 = new System.Windows.Forms.LinkLabel();
            this.labelX11 = new DevComponents.DotNetBar.LabelX();
            this.labelX7 = new DevComponents.DotNetBar.LabelX();
            this.cboSchoolYear = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.btnSaveConfig = new DevComponents.DotNetBar.ButtonX();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnCancel = new DevComponents.DotNetBar.ButtonX();
            this.lnkViewMapping = new System.Windows.Forms.LinkLabel();
            this.lnkUploadTenplate = new System.Windows.Forms.LinkLabel();
            this.lnkDownloadTemplate = new System.Windows.Forms.LinkLabel();
            this.picLoding = new System.Windows.Forms.PictureBox();
            ((System.ComponentModel.ISupportInitialize)(this.picLoding)).BeginInit();
            this.SuspendLayout();
            // 
            // cboConfigure
            // 
            this.cboConfigure.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.cboConfigure.DisplayMember = "Name";
            this.cboConfigure.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboConfigure.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboConfigure.FormattingEnabled = true;
            this.cboConfigure.ItemHeight = 19;
            this.cboConfigure.Location = new System.Drawing.Point(92, 10);
            this.cboConfigure.Name = "cboConfigure";
            this.cboConfigure.Size = new System.Drawing.Size(338, 25);
            this.cboConfigure.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboConfigure.TabIndex = 6;
            this.cboConfigure.SelectedIndexChanged += new System.EventHandler(this.cboConfigure_SelectedIndexChanged);
            // 
            // linkLabel4
            // 
            this.linkLabel4.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.linkLabel4.AutoSize = true;
            this.linkLabel4.BackColor = System.Drawing.Color.Transparent;
            this.linkLabel4.Location = new System.Drawing.Point(436, 16);
            this.linkLabel4.Name = "linkLabel4";
            this.linkLabel4.Size = new System.Drawing.Size(73, 17);
            this.linkLabel4.TabIndex = 9;
            this.linkLabel4.TabStop = true;
            this.linkLabel4.Text = "刪除設定檔";
            this.linkLabel4.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel4_LinkClicked);
            // 
            // labelX11
            // 
            this.labelX11.AutoSize = true;
            this.labelX11.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX11.BackgroundStyle.Class = "";
            this.labelX11.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX11.Location = new System.Drawing.Point(12, 12);
            this.labelX11.Name = "labelX11";
            this.labelX11.Size = new System.Drawing.Size(74, 21);
            this.labelX11.TabIndex = 12;
            this.labelX11.Text = "樣板設定檔";
            // 
            // labelX7
            // 
            this.labelX7.AutoSize = true;
            this.labelX7.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX7.BackgroundStyle.Class = "";
            this.labelX7.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX7.Location = new System.Drawing.Point(39, 48);
            this.labelX7.Name = "labelX7";
            this.labelX7.Size = new System.Drawing.Size(47, 21);
            this.labelX7.TabIndex = 13;
            this.labelX7.Text = "學年度";
            // 
            // cboSchoolYear
            // 
            this.cboSchoolYear.DisplayMember = "Text";
            this.cboSchoolYear.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboSchoolYear.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboSchoolYear.FormattingEnabled = true;
            this.cboSchoolYear.ItemHeight = 19;
            this.cboSchoolYear.Location = new System.Drawing.Point(92, 46);
            this.cboSchoolYear.Name = "cboSchoolYear";
            this.cboSchoolYear.Size = new System.Drawing.Size(61, 25);
            this.cboSchoolYear.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboSchoolYear.TabIndex = 8;
            // 
            // btnSaveConfig
            // 
            this.btnSaveConfig.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnSaveConfig.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnSaveConfig.BackColor = System.Drawing.Color.Transparent;
            this.btnSaveConfig.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnSaveConfig.Enabled = false;
            this.btnSaveConfig.Location = new System.Drawing.Point(272, 94);
            this.btnSaveConfig.Name = "btnSaveConfig";
            this.btnSaveConfig.Size = new System.Drawing.Size(75, 23);
            this.btnSaveConfig.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnSaveConfig.TabIndex = 14;
            this.btnSaveConfig.Text = "儲存設定";
            this.btnSaveConfig.Tooltip = "儲存當前的樣板設定。";
            this.btnSaveConfig.Visible = false;
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Enabled = false;
            this.btnPrint.Location = new System.Drawing.Point(360, 94);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 23);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 15;
            this.btnPrint.Text = "確定";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnCancel
            // 
            this.btnCancel.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnCancel.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnCancel.BackColor = System.Drawing.Color.Transparent;
            this.btnCancel.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnCancel.Location = new System.Drawing.Point(441, 94);
            this.btnCancel.Name = "btnCancel";
            this.btnCancel.Size = new System.Drawing.Size(75, 23);
            this.btnCancel.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnCancel.TabIndex = 16;
            this.btnCancel.Text = "取消";
            this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
            // 
            // lnkViewMapping
            // 
            this.lnkViewMapping.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lnkViewMapping.AutoSize = true;
            this.lnkViewMapping.BackColor = System.Drawing.Color.Transparent;
            this.lnkViewMapping.Location = new System.Drawing.Point(197, 100);
            this.lnkViewMapping.Name = "lnkViewMapping";
            this.lnkViewMapping.Size = new System.Drawing.Size(112, 17);
            this.lnkViewMapping.TabIndex = 19;
            this.lnkViewMapping.TabStop = true;
            this.lnkViewMapping.Text = "下載合併欄位總表";
            this.lnkViewMapping.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkViewMapping_LinkClicked);
            // 
            // lnkUploadTenplate
            // 
            this.lnkUploadTenplate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lnkUploadTenplate.AutoSize = true;
            this.lnkUploadTenplate.BackColor = System.Drawing.Color.Transparent;
            this.lnkUploadTenplate.Location = new System.Drawing.Point(105, 100);
            this.lnkUploadTenplate.Name = "lnkUploadTenplate";
            this.lnkUploadTenplate.Size = new System.Drawing.Size(86, 17);
            this.lnkUploadTenplate.TabIndex = 18;
            this.lnkUploadTenplate.TabStop = true;
            this.lnkUploadTenplate.Text = "變更套印樣板";
            this.lnkUploadTenplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkUploadTenplate_LinkClicked);
            // 
            // lnkDownloadTemplate
            // 
            this.lnkDownloadTemplate.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.lnkDownloadTemplate.AutoSize = true;
            this.lnkDownloadTemplate.BackColor = System.Drawing.Color.Transparent;
            this.lnkDownloadTemplate.Location = new System.Drawing.Point(13, 100);
            this.lnkDownloadTemplate.Name = "lnkDownloadTemplate";
            this.lnkDownloadTemplate.Size = new System.Drawing.Size(86, 17);
            this.lnkDownloadTemplate.TabIndex = 17;
            this.lnkDownloadTemplate.TabStop = true;
            this.lnkDownloadTemplate.Text = "檢視套印樣板";
            this.lnkDownloadTemplate.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkDownloadTemplate_LinkClicked);
            // 
            // picLoding
            // 
            this.picLoding.BackColor = System.Drawing.Color.Transparent;
            this.picLoding.Image = global::SH_ClassYearScoreReport.Properties.Resources.loading;
            this.picLoding.Location = new System.Drawing.Point(241, 41);
            this.picLoding.Name = "picLoding";
            this.picLoding.Size = new System.Drawing.Size(35, 34);
            this.picLoding.TabIndex = 20;
            this.picLoding.TabStop = false;
            this.picLoding.Visible = false;
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(524, 123);
            this.Controls.Add(this.picLoding);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.btnCancel);
            this.Controls.Add(this.cboConfigure);
            this.Controls.Add(this.cboSchoolYear);
            this.Controls.Add(this.lnkViewMapping);
            this.Controls.Add(this.lnkUploadTenplate);
            this.Controls.Add(this.lnkDownloadTemplate);
            this.Controls.Add(this.linkLabel4);
            this.Controls.Add(this.labelX11);
            this.Controls.Add(this.labelX7);
            this.Controls.Add(this.btnSaveConfig);
            this.DoubleBuffered = true;
            this.Name = "MainForm";
            this.Text = "";
            this.Load += new System.EventHandler(this.MainForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.picLoding)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.Controls.ComboBoxEx cboConfigure;
        private System.Windows.Forms.LinkLabel linkLabel4;
        private DevComponents.DotNetBar.LabelX labelX11;
        private DevComponents.DotNetBar.LabelX labelX7;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboSchoolYear;
        private DevComponents.DotNetBar.ButtonX btnSaveConfig;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnCancel;
        private System.Windows.Forms.LinkLabel lnkViewMapping;
        private System.Windows.Forms.LinkLabel lnkUploadTenplate;
        private System.Windows.Forms.LinkLabel lnkDownloadTemplate;
        private System.Windows.Forms.PictureBox picLoding;
    }
}