namespace StudentCourseScoreCheck100_SH
{
    partial class StudentCourseScoreForm
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
            System.Windows.Forms.ListViewItem listViewItem1 = new System.Windows.Forms.ListViewItem("1年級");
            System.Windows.Forms.ListViewItem listViewItem2 = new System.Windows.Forms.ListViewItem("2年級");
            System.Windows.Forms.ListViewItem listViewItem3 = new System.Windows.Forms.ListViewItem("3年級");
            this.labelX1 = new DevComponents.DotNetBar.LabelX();
            this.labelX2 = new DevComponents.DotNetBar.LabelX();
            this.lvData = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.labelX3 = new DevComponents.DotNetBar.LabelX();
            this.btnPrint = new DevComponents.DotNetBar.ButtonX();
            this.btnExit = new DevComponents.DotNetBar.ButtonX();
            this.iptSchoolYear = new DevComponents.Editors.IntegerInput();
            this.iptSemester = new DevComponents.Editors.IntegerInput();
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).BeginInit();
            this.SuspendLayout();
            // 
            // labelX1
            // 
            this.labelX1.AutoSize = true;
            this.labelX1.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX1.BackgroundStyle.Class = "";
            this.labelX1.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX1.Location = new System.Drawing.Point(12, 9);
            this.labelX1.Name = "labelX1";
            this.labelX1.Size = new System.Drawing.Size(47, 21);
            this.labelX1.TabIndex = 1;
            this.labelX1.Text = "學年度";
            // 
            // labelX2
            // 
            this.labelX2.AutoSize = true;
            this.labelX2.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX2.BackgroundStyle.Class = "";
            this.labelX2.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX2.Location = new System.Drawing.Point(150, 9);
            this.labelX2.Name = "labelX2";
            this.labelX2.Size = new System.Drawing.Size(34, 21);
            this.labelX2.TabIndex = 3;
            this.labelX2.Text = "學期";
            // 
            // lvData
            // 
            this.lvData.Alignment = System.Windows.Forms.ListViewAlignment.Left;
            this.lvData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            // 
            // 
            // 
            this.lvData.Border.Class = "ListViewBorder";
            this.lvData.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lvData.CheckBoxes = true;
            listViewItem1.StateImageIndex = 0;
            listViewItem2.StateImageIndex = 0;
            listViewItem3.StateImageIndex = 0;
            this.lvData.Items.AddRange(new System.Windows.Forms.ListViewItem[] {
            listViewItem1,
            listViewItem2,
            listViewItem3});
            this.lvData.Location = new System.Drawing.Point(61, 38);
            this.lvData.Name = "lvData";
            this.lvData.Scrollable = false;
            this.lvData.Size = new System.Drawing.Size(189, 60);
            this.lvData.TabIndex = 4;
            this.lvData.UseCompatibleStateImageBehavior = false;
            this.lvData.View = System.Windows.Forms.View.List;
            // 
            // labelX3
            // 
            this.labelX3.AutoSize = true;
            this.labelX3.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labelX3.BackgroundStyle.Class = "";
            this.labelX3.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labelX3.Location = new System.Drawing.Point(25, 36);
            this.labelX3.Name = "labelX3";
            this.labelX3.Size = new System.Drawing.Size(34, 21);
            this.labelX3.TabIndex = 5;
            this.labelX3.Text = "年級";
            // 
            // btnPrint
            // 
            this.btnPrint.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnPrint.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnPrint.AutoSize = true;
            this.btnPrint.BackColor = System.Drawing.Color.Transparent;
            this.btnPrint.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnPrint.Location = new System.Drawing.Point(88, 106);
            this.btnPrint.Name = "btnPrint";
            this.btnPrint.Size = new System.Drawing.Size(75, 25);
            this.btnPrint.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnPrint.TabIndex = 6;
            this.btnPrint.Text = "列印";
            this.btnPrint.Click += new System.EventHandler(this.btnPrint_Click);
            // 
            // btnExit
            // 
            this.btnExit.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnExit.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.btnExit.AutoSize = true;
            this.btnExit.BackColor = System.Drawing.Color.Transparent;
            this.btnExit.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnExit.Location = new System.Drawing.Point(174, 106);
            this.btnExit.Name = "btnExit";
            this.btnExit.Size = new System.Drawing.Size(75, 25);
            this.btnExit.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btnExit.TabIndex = 7;
            this.btnExit.Text = "離開";
            this.btnExit.Click += new System.EventHandler(this.btnExit_Click);
            // 
            // iptSchoolYear
            // 
            this.iptSchoolYear.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptSchoolYear.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptSchoolYear.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptSchoolYear.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptSchoolYear.Location = new System.Drawing.Point(60, 7);
            this.iptSchoolYear.MaxValue = 1000;
            this.iptSchoolYear.MinValue = 1;
            this.iptSchoolYear.Name = "iptSchoolYear";
            this.iptSchoolYear.ShowUpDown = true;
            this.iptSchoolYear.Size = new System.Drawing.Size(80, 25);
            this.iptSchoolYear.TabIndex = 8;
            this.iptSchoolYear.Value = 1;
            // 
            // iptSemester
            // 
            this.iptSemester.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.iptSemester.BackgroundStyle.Class = "DateTimeInputBackground";
            this.iptSemester.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.iptSemester.ButtonFreeText.Shortcut = DevComponents.DotNetBar.eShortcut.F2;
            this.iptSemester.Location = new System.Drawing.Point(189, 7);
            this.iptSemester.MaxValue = 2;
            this.iptSemester.MinValue = 1;
            this.iptSemester.Name = "iptSemester";
            this.iptSemester.ShowUpDown = true;
            this.iptSemester.Size = new System.Drawing.Size(59, 25);
            this.iptSemester.TabIndex = 9;
            this.iptSemester.Value = 1;
            // 
            // StudentCourseScoreForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(262, 138);
            this.Controls.Add(this.iptSemester);
            this.Controls.Add(this.iptSchoolYear);
            this.Controls.Add(this.btnExit);
            this.Controls.Add(this.btnPrint);
            this.Controls.Add(this.labelX3);
            this.Controls.Add(this.lvData);
            this.Controls.Add(this.labelX2);
            this.Controls.Add(this.labelX1);
            this.DoubleBuffered = true;
            this.Name = "StudentCourseScoreForm";
            this.Text = "課程及評量成績百分檢查";
            this.Load += new System.EventHandler(this.StudentCourseScoreForm_Load);
            ((System.ComponentModel.ISupportInitialize)(this.iptSchoolYear)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.iptSemester)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevComponents.DotNetBar.LabelX labelX1;
        private DevComponents.DotNetBar.LabelX labelX2;
        private DevComponents.DotNetBar.Controls.ListViewEx lvData;
        private DevComponents.DotNetBar.LabelX labelX3;
        private DevComponents.DotNetBar.ButtonX btnPrint;
        private DevComponents.DotNetBar.ButtonX btnExit;
        private DevComponents.Editors.IntegerInput iptSchoolYear;
        private DevComponents.Editors.IntegerInput iptSemester;
    }
}