namespace SHImportExportScoreRank.DetailContent
{
    partial class UCStudentRank
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

        #region Component Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.lvData = new DevComponents.DotNetBar.Controls.ListViewEx();
            this.SuspendLayout();
            //
            // lvData
            //
            this.lvData.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom)
            | System.Windows.Forms.AnchorStyles.Left)
            | System.Windows.Forms.AnchorStyles.Right)));
            //
            //
            //
            this.lvData.Border.Class = "ListViewBorder";
            this.lvData.Border.CornerType = DevComponents.DotNetBar.eCornerType.Square;
            this.lvData.FullRowSelect = true;
            this.lvData.HideSelection = false;
            this.lvData.Location = new System.Drawing.Point(14, 14);
            this.lvData.MultiSelect = false;
            this.lvData.Name = "lvData";
            this.lvData.Size = new System.Drawing.Size(522, 227);
            this.lvData.TabIndex = 0;
            this.lvData.UseCompatibleStateImageBehavior = false;
            this.lvData.View = System.Windows.Forms.View.Details;
            //
            // UCStudentSchoolYearEntryRank
            //
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 17F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.Controls.Add(this.lvData);
            this.Name = "UCStudentSchoolYearEntryRank";
            this.Size = new System.Drawing.Size(550, 255);
            this.ResumeLayout(false);
        }

        #endregion

        private DevComponents.DotNetBar.Controls.ListViewEx lvData;
    }
}
