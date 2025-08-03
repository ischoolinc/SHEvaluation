
namespace SmartSchool.Evaluation.Content
{
    partial class SemesterScoreDataAchive
    {
        /// <summary> 
        /// 設計工具所需的變數。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary> 
        /// 清除任何使用中的資源。
        /// </summary>
        /// <param name="disposing">如果應該處置受控資源則為 true，否則為 false。</param>
        protected override void Dispose(bool disposing)
        {
            //FillListView();
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region 元件設計工具產生的程式碼

        /// <summary> 
        /// 此為設計工具支援所需的方法 - 請勿使用程式碼編輯器修改
        /// 這個方法的內容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SemesterScoreDataAchive));
            this.btnDelete = new DevComponents.DotNetBar.ButtonX();
            this.btnView = new DevComponents.DotNetBar.ButtonX();
            this.picWaiting = new System.Windows.Forms.PictureBox();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.listView1 = new SmartSchool.Common.ListViewEX();
            this.colSchoolYear = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colSemester = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.cboGradeYear = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.col學業成績 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.columnTime = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.btnAdd = new DevComponents.DotNetBar.ButtonX();
            ((System.ComponentModel.ISupportInitialize)(this.picWaiting)).BeginInit();
            this.SuspendLayout();
            // 
            // btnDelete
            // 
            this.btnDelete.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnDelete.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnDelete.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnDelete.Enabled = false;
            this.btnDelete.Location = new System.Drawing.Point(183, 194);
            this.btnDelete.Name = "btnDelete";
            this.btnDelete.Size = new System.Drawing.Size(75, 23);
            this.btnDelete.TabIndex = 6;
            this.btnDelete.Text = "刪除";
            this.btnDelete.Click += new System.EventHandler(this.btnDelete_Click);
            // 
            // btnView
            // 
            this.btnView.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnView.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnView.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnView.Enabled = false;
            this.btnView.Location = new System.Drawing.Point(102, 194);
            this.btnView.Name = "btnView";
            this.btnView.Size = new System.Drawing.Size(75, 23);
            this.btnView.TabIndex = 7;
            this.btnView.Text = "修改";
            this.btnView.Click += new System.EventHandler(this.btnView_Click);
            // 
            // picWaiting
            // 
            this.picWaiting.BackColor = System.Drawing.SystemColors.Window;
            this.picWaiting.Image = global::SmartSchool.Evaluation.Properties.Resources.loading5;
            this.picWaiting.Location = new System.Drawing.Point(259, 91);
            this.picWaiting.Margin = new System.Windows.Forms.Padding(4);
            this.picWaiting.Name = "picWaiting";
            this.picWaiting.Size = new System.Drawing.Size(32, 32);
            this.picWaiting.SizeMode = System.Windows.Forms.PictureBoxSizeMode.AutoSize;
            this.picWaiting.TabIndex = 8;
            this.picWaiting.TabStop = false;
            this.picWaiting.Visible = false;
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "error.ico");
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
            this.columnTime});
            this.listView1.FullRowSelect = true;
            this.listView1.HideSelection = false;
            this.listView1.Location = new System.Drawing.Point(17, 16);
            this.listView1.MultiSelect = false;
            this.listView1.Name = "listView1";
            this.listView1.ShowItemToolTips = true;
            this.listView1.Size = new System.Drawing.Size(517, 171);
            this.listView1.TabIndex = 9;
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
            // columnTime
            // 
            this.columnTime.Text = "封存時間";
            this.columnTime.Width = 200;
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(495, 198);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(34, 17);
            this.linkLabel1.TabIndex = 10;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "說明";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // btnAdd
            // 
            this.btnAdd.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btnAdd.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnAdd.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btnAdd.Location = new System.Drawing.Point(19, 194);
            this.btnAdd.Name = "btnAdd";
            this.btnAdd.Size = new System.Drawing.Size(75, 23);
            this.btnAdd.TabIndex = 11;
            this.btnAdd.Text = "新增";
            this.btnAdd.Click += new System.EventHandler(this.btnAdd_Click);
            // 
            // SemesterScoreDataAchive
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.btnAdd);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.picWaiting);
            this.Controls.Add(this.btnDelete);
            this.Controls.Add(this.btnView);
            this.Controls.Add(this.listView1);
            this.Name = "SemesterScoreDataAchive";
            this.Size = new System.Drawing.Size(550, 225);
            ((System.ComponentModel.ISupportInitialize)(this.picWaiting)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        protected System.Windows.Forms.PictureBox picWaiting;
        private DevComponents.DotNetBar.ButtonX btnDelete;
        private DevComponents.DotNetBar.ButtonX btnView;
        private System.Windows.Forms.ImageList imageList1;
        private Common.ListViewEX listView1;
        private System.Windows.Forms.ColumnHeader colSchoolYear;
        private System.Windows.Forms.ColumnHeader colSemester;
        private System.Windows.Forms.ColumnHeader cboGradeYear;
        private System.Windows.Forms.ColumnHeader col學業成績;
        private System.Windows.Forms.ColumnHeader columnTime;
        private System.Windows.Forms.LinkLabel linkLabel1;
        private DevComponents.DotNetBar.ButtonX btnAdd;
    }
}
