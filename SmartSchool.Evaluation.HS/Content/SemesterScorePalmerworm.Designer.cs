﻿namespace SmartSchool.Evaluation.Content
{
    partial class SemesterScorePalmerworm
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
            UpdatePreference();
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改這個方法的內容。
        ///
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SemesterScorePalmerworm));
            this.listView1 = new SmartSchool.Common.ListViewEX();
            this.colSchoolYear = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colSemester = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.cboGradeYear = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.col學業成績 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.col實習科目 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.col專業科目 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.col體育成績 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.col國防通識 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.col健康與護理 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.col德行成績 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.col實得學分 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.lvColumnManager1 = new IntelliSchool.DSA.ClientFramework.ControlCommunication.LVColumnManager();
            this.btnAdd = new DevComponents.DotNetBar.ButtonX();
            this.btnModify = new DevComponents.DotNetBar.ButtonX();
            this.btnDelete = new DevComponents.DotNetBar.ButtonX();
            this.picWaiting = new System.Windows.Forms.PictureBox();
            this.btnView = new DevComponents.DotNetBar.ButtonX();
            this.btnArchive = new DevComponents.DotNetBar.ButtonX();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.btnChangeSchoolYear = new DevComponents.DotNetBar.ButtonX();
            ((System.ComponentModel.ISupportInitialize)(this.picWaiting)).BeginInit();
            this.SuspendLayout();
            // 
            // listView1
            // 
            this.listView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            // 
            // 
            // 
            this.listView1.Border.Class = "ListViewBorder";
            this.listView1.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.listView1.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colSchoolYear,
            this.colSemester,
            this.cboGradeYear,
            this.col學業成績,
            this.col實習科目,
            this.col專業科目,
            this.col體育成績,
            this.col國防通識,
            this.col健康與護理,
            this.col德行成績,
            this.col實得學分});
            this.listView1.FullRowSelect = true;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(17, 16);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.ShowItemToolTips = true;
            this.listView1.Size = new System.Drawing.Size(517, 171);
            this.listView1.SmallImageList = this.imageList1;
            this.listView1.TabIndex = 1;
            this.listView1.UseCompatibleStateImageBehavior = false;
            this.listView1.View = System.Windows.Forms.View.Details;
            this.listView1.SelectedIndexChanged += new System.EventHandler(this.listView1_SelectedIndexChanged);
            this.listView1.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.listView1_MouseDoubleClick);
            // 
            // colSchoolYear
            // 
            this.colSchoolYear.Text = "學年度";
            // 
            // colSemester
            // 
            this.colSemester.Text = "學期";
            this.colSemester.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // cboGradeYear
            // 
            this.cboGradeYear.Text = "年級";
            this.cboGradeYear.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            // 
            // col學業成績
            // 
            this.col學業成績.Text = "學業成績";
            this.col學業成績.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.col學業成績.Width = 85;
            // 
            // col實習科目
            // 
            this.col實習科目.Text = "實習科目";
            this.col實習科目.Width = 85;
            // 
            // col專業科目
            // 
            this.col專業科目.Text = "專業科目";
            this.col專業科目.Width = 85;
            // 
            // col體育成績
            // 
            this.col體育成績.Text = "體育成績";
            this.col體育成績.Width = 85;
            // 
            // col國防通識
            // 
            this.col國防通識.Text = "國防通識";
            this.col國防通識.Width = 85;
            // 
            // col健康與護理
            // 
            this.col健康與護理.Text = "健康與護理";
            this.col健康與護理.Width = 85;
            // 
            // col德行成績
            // 
            this.col德行成績.Text = "德行成績";
            this.col德行成績.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.col德行成績.Width = 85;
            // 
            // col實得學分
            // 
            this.col實得學分.Text = "實得學分";
            this.col實得學分.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.col實得學分.Width = 85;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "error.ico");
            // 
            // lvColumnManager1
            // 
            this.lvColumnManager1.TargetListView = this.listView1;
            // 
            // btnAdd
            // 
            this.btnAdd.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAdd.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAdd.Location = new System.Drawing.Point(17, 195);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 2;
            this.btnAdd.Text = "新增";
            this.btnAdd.Click += new System.EventHandler(this.buttonX1_Click);
            // 
            // btnModify
            // 
            this.btnModify.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnModify.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnModify.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnModify.Enabled = false;
            this.btnModify.Location = new System.Drawing.Point(98, 195);
            this.btnModify.Name = "btnModify";
            this.btnModify.Size = new System.Drawing.Size(75, 23);
            this.btnModify.TabIndex = 2;
            this.btnModify.Text = "修改";
            this.btnModify.Click += new System.EventHandler(this.buttonX2_Click);
            // 
            // btnDelete
            // 
            this.btnDelete.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDelete.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnDelete.Enabled = false;
            this.btnDelete.Location = new System.Drawing.Point(179, 195);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.TabIndex = 2;
            this.btnDelete.Text = "刪除";
            this.btnDelete.Click += new System.EventHandler(this.buttonX3_Click);
            // 
            // picWaiting
            // 
            this.picWaiting.BackColor = System.Drawing.SystemColors.Window;
            this.picWaiting.Image = global::SmartSchool.Evaluation.Properties.Resources.loading5;
            this.picWaiting.Location = new System.Drawing.Point(259, 96);
            this.picWaiting.Margin = new System.Windows.Forms.Padding(4);
            this.picWaiting.Name = "picWaiting";
            this.picWaiting.Size = new System.Drawing.Size(32, 32);
            this.picWaiting.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.picWaiting.TabIndex = 4;
            this.picWaiting.TabStop = false;
            this.picWaiting.Visible = false;
            // 
            // btnView
            // 
            this.btnView.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnView.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnView.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnView.Enabled = false;
            this.btnView.Location = new System.Drawing.Point(17, 195);
            this.btnView.Name = "btnView";
            this.btnView.Size = new System.Drawing.Size(75, 23);
            this.btnView.TabIndex = 5;
            this.btnView.Text = "檢視";
            this.btnView.Click += new System.EventHandler(this.btnView_Click);
            // 
            // btnArchive
            // 
            this.btnArchive.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnArchive.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnArchive.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnArchive.Location = new System.Drawing.Point(260, 195);
            this.btnArchive.Name = "btnArchive";
            this.btnArchive.Size = new System.Drawing.Size(75, 23);
            this.btnArchive.TabIndex = 6;
            this.btnArchive.Text = "封存";
            this.btnArchive.Click += new System.EventHandler(this.btnArchive_Click);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(495, 198);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(29, 12);
            this.linkLabel1.TabIndex = 11;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "說明";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // btnChangeSchoolYear
            // 
            this.btnChangeSchoolYear.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnChangeSchoolYear.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnChangeSchoolYear.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnChangeSchoolYear.Location = new System.Drawing.Point(341, 195);
            this.btnChangeSchoolYear.Name = "btnChangeSchoolYear";
            this.btnChangeSchoolYear.Size = new System.Drawing.Size(91, 23);
            this.btnChangeSchoolYear.TabIndex = 12;
            this.btnChangeSchoolYear.Text = "調整學年度";
            this.btnChangeSchoolYear.Click += new System.EventHandler(this.btnChangeSchoolYear_Click);
            // 
            // SemesterScorePalmerworm
            // 
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None;
            this.Controls.Add(this.btnChangeSchoolYear);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.btnArchive);
            this.Controls.Add(this.picWaiting);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnModify);
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.listView1);
            this.Controls.Add(this.btnView);
            this.Name = "SemesterScorePalmerworm";
            this.Size = new System.Drawing.Size(550, 225);
            ((System.ComponentModel.ISupportInitialize)(this.picWaiting)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private SmartSchool.Common.ListViewEX listView1;
        private System.Windows.Forms.ColumnHeader colSchoolYear;
        private System.Windows.Forms.ColumnHeader colSemester;
        private System.Windows.Forms.ColumnHeader col學業成績;
        private System.Windows.Forms.ColumnHeader col德行成績;
        private System.Windows.Forms.ColumnHeader col實得學分;
        private System.Windows.Forms.ColumnHeader col體育成績;
        private System.Windows.Forms.ColumnHeader col國防通識;
        private System.Windows.Forms.ColumnHeader col健康與護理;
        private System.Windows.Forms.ColumnHeader col實習科目;
        private IntelliSchool.DSA.ClientFramework.ControlCommunication.LVColumnManager lvColumnManager1;
        private DevComponents.DotNetBar.ButtonX btnAdd;
        private DevComponents.DotNetBar.ButtonX btnModify;
        private DevComponents.DotNetBar.ButtonX btnDelete;
        private System.Windows.Forms.ColumnHeader cboGradeYear;
        private System.Windows.Forms.ImageList imageList1;
        protected System.Windows.Forms.PictureBox picWaiting;
        private DevComponents.DotNetBar.ButtonX btnView;
        private System.Windows.Forms.ColumnHeader col專業科目;
        private DevComponents.DotNetBar.ButtonX btnArchive;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private DevComponents.DotNetBar.ButtonX btnChangeSchoolYear;
    }
}
