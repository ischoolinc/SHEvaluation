namespace 班級定期評量成績單_固定排名
{
    partial class FixedRankInclued
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
            this.dataGridViewX1 = new DevComponents.DotNetBar.Controls.DataGridViewX();
            this.subject = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.btncheckSubj = new DevComponents.DotNetBar.ButtonX();
            this.cboRankType = new DevComponents.DotNetBar.Controls.ComboBoxEx();
            this.labClass = new DevComponents.DotNetBar.LabelX();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewX1)).BeginInit();
            this.SuspendLayout();
            // 
            // dataGridViewX1
            // 
            this.dataGridViewX1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewX1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewX1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.subject});
            dataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft;
            dataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window;
            dataGridViewCellStyle2.Font = new System.Drawing.Font("微軟正黑體", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(136)));
            dataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight;
            dataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.ControlText;
            dataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.False;
            this.dataGridViewX1.DefaultCellStyle = dataGridViewCellStyle2;
            this.dataGridViewX1.GridColor = System.Drawing.Color.FromArgb(((int)(((byte)(208)))), ((int)(((byte)(215)))), ((int)(((byte)(229)))));
            this.dataGridViewX1.Location = new System.Drawing.Point(12, 45);
            this.dataGridViewX1.Name = "dataGridViewX1";
            this.dataGridViewX1.RowTemplate.Height = 24;
            this.dataGridViewX1.Size = new System.Drawing.Size(445, 376);
            this.dataGridViewX1.TabIndex = 0;
            // 
            // subject
            // 
            this.subject.AutoSizeMode = System.Windows.Forms.DataGridViewAutoSizeColumnMode.Fill;
            this.subject.HeaderText = "科目";
            this.subject.Name = "subject";
            this.subject.ReadOnly = true;
            // 
            // btncheckSubj
            // 
            this.btncheckSubj.AccessibleRole = System.Windows.Forms.AccessibleRole.PushButton;
            this.btncheckSubj.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.btncheckSubj.BackColor = System.Drawing.Color.Transparent;
            this.btncheckSubj.ColorTable = DevComponents.DotNetBar.eButtonColor.OrangeWithBackground;
            this.btncheckSubj.Location = new System.Drawing.Point(235, 432);
            this.btncheckSubj.Name = "btncheckSubj";
            this.btncheckSubj.Size = new System.Drawing.Size(222, 23);
            this.btncheckSubj.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.btncheckSubj.TabIndex = 1;
            this.btncheckSubj.Text = "帶入所選科目";
            this.btncheckSubj.Click += new System.EventHandler(this.btncheckSubj_Click);
            // 
            // cboRankType
            // 
            this.cboRankType.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.cboRankType.DisplayMember = "Text";
            this.cboRankType.DrawMode = System.Windows.Forms.DrawMode.OwnerDrawFixed;
            this.cboRankType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboRankType.FormattingEnabled = true;
            this.cboRankType.ItemHeight = 19;
            this.cboRankType.Location = new System.Drawing.Point(321, 12);
            this.cboRankType.Name = "cboRankType";
            this.cboRankType.Size = new System.Drawing.Size(138, 25);
            this.cboRankType.Style = DevComponents.DotNetBar.eDotNetBarStyle.StyleManagerControlled;
            this.cboRankType.TabIndex = 2;
            this.cboRankType.SelectionChangeCommitted += new System.EventHandler(this.cboRankType_SelectionChangeCommitted);
            // 
            // labClass
            // 
            this.labClass.BackColor = System.Drawing.Color.Transparent;
            // 
            // 
            // 
            this.labClass.BackgroundStyle.Class = "";
            this.labClass.BackgroundStyle.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.labClass.Location = new System.Drawing.Point(187, 13);
            this.labClass.Name = "labClass";
            this.labClass.Size = new System.Drawing.Size(131, 23);
            this.labClass.TabIndex = 3;
            this.labClass.Text = "固定排名計算類別：";
            // 
            // FixedRankInclued
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(471, 462);
            this.Controls.Add(this.labClass);
            this.Controls.Add(this.cboRankType);
            this.Controls.Add(this.btncheckSubj);
            this.Controls.Add(this.dataGridViewX1);
            this.DoubleBuffered = true;
            this.Name = "FixedRankInclued";
            this.Text = "所選班級固定排名計算之科目";
            this.Load += new System.EventHandler(this.FixedRankInclued_Load);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewX1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevComponents.DotNetBar.Controls.DataGridViewX dataGridViewX1;
        private System.Windows.Forms.DataGridViewTextBoxColumn subject;
        private DevComponents.DotNetBar.ButtonX btncheckSubj;
        private DevComponents.DotNetBar.Controls.ComboBoxEx cboRankType;
        private DevComponents.DotNetBar.LabelX labClass;
    }
}