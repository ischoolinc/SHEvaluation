using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using SHSchool.Data;
using SHCourseScoreInputStatus.DAO;

namespace SHCourseScoreInputStatus
{
    public partial class CourseScoreInputForm : FISCA.Presentation.Controls.BaseForm
    {
        BackgroundWorker _bgWork;
        private int _SchoolYear = 0;
        private int _Semester = 0;
        private bool _chkNotInput = false;
        List<CourseScoreBase> _CourseScoreBaseList;
        private ListViewColumnSorter lvwColumnSorter;


        public CourseScoreInputForm()
        {
            InitializeComponent();
            // 放入學年度學期,現在學年度、學期
            int sy = int.Parse(K12.Data.School.DefaultSchoolYear);

            for (int i = (sy - 3); i <= (sy + 3); i++)
                cbxSchoolYear.Items.Add(i);

            cbxSemester.Items.Add("1");
            cbxSemester.Items.Add("2");

            cbxSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
            cbxSemester.Text = K12.Data.School.DefaultSemester;

            _CourseScoreBaseList = new List<CourseScoreBase>();
            _bgWork = new BackgroundWorker();
            _bgWork.DoWork += new DoWorkEventHandler(_bgWork_DoWork);
            _bgWork.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWork_RunWorkerCompleted);
            //LoadData();

            lvwColumnSorter = new ListViewColumnSorter();
            this.lvData.ListViewItemSorter = lvwColumnSorter;
        }

        void _bgWork_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            btnExport.Enabled = true;
            //btnAddTemp.Enabled = true;
            btnReload.Enabled = true;
            BindDataToListView();
        }

        // 將資料放入畫面ListView
        private void BindDataToListView()
        {
            lvData.Items.Clear();
            foreach (CourseScoreBase csb in _CourseScoreBaseList)
            {
                ListViewItem lvi = new ListViewItem();
                //System.Windows.Forms.ListViewItem.ListViewSubItem subitem1 = lvi.SubItems.Add("");
                lvi.Tag = csb;
                lvi.Text = csb.CourseName;
                lvi.SubItems.Add(csb.TeacherName);

                lvi.UseItemStyleForSubItems = false;

                string val = "--/--";
                bool chkNotInput = true;
                if (csb.hasScoreCount.HasValue && csb.CourseStudentCount.HasValue)
                {
                    val = csb.hasScoreCount.Value + "/" + csb.CourseStudentCount.Value;

                    if (csb.hasScoreCount.Value == csb.CourseStudentCount.Value)
                        chkNotInput = false;

                }
                if (csb.hasScoreCount.HasValue && csb.CourseStudentCount.HasValue == false)
                {
                    val = csb.hasScoreCount.Value + "/--";
                }
                if (csb.hasScoreCount.HasValue == false && csb.CourseStudentCount.HasValue)
                {
                    val = "--/" + csb.CourseStudentCount.Value;
                }

                lvi.SubItems.Add(val);
                if (chkNotInput)
                    lvi.SubItems[2].ForeColor = Color.Red;

                lvi.SubItems.Add(csb.ScoreSource);
                // 處理當勾選只列出未輸入完成
                if (chkNotHasScore.Checked)
                {
                    if (chkNotInput)
                        lvData.Items.Add(lvi);
                }
                else
                    lvData.Items.Add(lvi);
            }
            lblMsg.Text = "共 " + lvData.Items.Count + " 筆課程";
            //btnReload.Enabled = false;
        }

        void _bgWork_DoWork(object sender, DoWorkEventArgs e)
        {
            _CourseScoreBaseList = QueryData.GetCourseScoreBaseByCourseSchoolYearSemester(_SchoolYear, _Semester);
        }

        private void CourseScoreInputForm_Load(object sender, EventArgs e)
        {

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnReload_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        // 載入資料
        private void LoadData()
        {

            int sc, ss;
            if (int.TryParse(cbxSchoolYear.Text, out sc))
            {
                _SchoolYear = sc;
            }
            else
                _SchoolYear = 0;

            if (int.TryParse(cbxSemester.Text, out ss))
            {
                _Semester = ss;
            }
            else
                _Semester = 0;

            btnExport.Enabled = false;
            btnAddTemp.Enabled = false;
            btnReload.Enabled = false;
            lvData.Items.Clear();
            _bgWork.RunWorkerAsync();
        }



        private void btnAddTemp_Click(object sender, EventArgs e)
        {
            if (lvData.SelectedItems.Count > 0)
            {
                btnAddTemp.Enabled = false;
                // 將所選的加入待處理
                List<string> addIDList = new List<string>();
                foreach (ListViewItem lvi in lvData.SelectedItems)
                {
                    CourseScoreBase data = lvi.Tag as CourseScoreBase;
                    if (data != null)
                    {
                        if (!K12.Presentation.NLDPanels.Course.TempSource.Contains(data.CourseID))
                            addIDList.Add(data.CourseID);
                    }
                }
                if (addIDList.Count > 0)
                    K12.Presentation.NLDPanels.Course.AddToTemp(addIDList);

                FISCA.Presentation.Controls.MsgBox.Show("加入 " + addIDList.Count + " 筆課程至課程待處理");

                btnAddTemp.Enabled = true;
            }
        }

        private void chkNotHasScore_CheckedChanged(object sender, EventArgs e)
        {
            //if (lvData.Items.Count > 0)
            BindDataToListView();
        }

        private void cbxSemester_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxSemester.Text != _Semester.ToString())
                btnReload.Enabled = true;
            //else
            //    btnReload.Enabled = false;
        }

        private void cbxSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbxSchoolYear.Text != _SchoolYear.ToString())
                btnReload.Enabled = true;
            //else
            //    btnReload.Enabled = false;
        }

        private void btnExport_Click(object sender, EventArgs e)
        {
            btnExport.Enabled = false;
            if (lvData.Items.Count > 0)
            {
                DataTable dt = new DataTable();
                Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
                dt.Columns.Add("課程名稱");
                dt.Columns.Add("授課教師");
                dt.Columns.Add("課程成績");

                foreach (ListViewItem lvi in lvData.Items)
                {
                    DataRow dr = dt.NewRow();
                    dr["課程名稱"] = lvi.Text;
                    if (lvi.SubItems[1] != null)
                        dr["授課教師"] = lvi.SubItems[1].Text;

                    if (lvi.SubItems[2] != null)
                        dr["課程成績"] = lvi.SubItems[2].Text;

                    dt.Rows.Add(dr);
                }
                Utility.CompletedXls("課程成績輸入狀況", dt, wb);
            }
            btnExport.Enabled = true;
        }

        private void lvData_ColumnClick(object sender, ColumnClickEventArgs e)
        {
            // Determine if clicked column is already the column that is being sorted.
            if (e.Column == lvwColumnSorter.SortColumn)
            {
                // Reverse the current sort direction for this column.
                if (lvwColumnSorter.Order == SortOrder.Ascending)
                {
                    lvwColumnSorter.Order = SortOrder.Descending;
                }
                else
                {
                    lvwColumnSorter.Order = SortOrder.Ascending;
                }
            }
            else
            {
                // Set the column number that is to be sorted; default to ascending.
                lvwColumnSorter.SortColumn = e.Column;
                lvwColumnSorter.Order = SortOrder.Ascending;
            }

            // Perform the sort with these new sort options.
            this.lvData.Sort();
        }

        private void lvData_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (lvData.SelectedItems.Count > 0)
                btnAddTemp.Enabled = true;
            else
                btnAddTemp.Enabled = false;
        }
    }
}
