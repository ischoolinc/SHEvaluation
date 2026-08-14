using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text;
using System.Windows.Forms;
using FISCA.Authentication;
using FISCA.Presentation.Controls;
using SHImportExportScoreRank.DAO;
using SHImportExportScoreRank.Permissions;
using SHImportExportScoreRank.UIForm;

namespace SHImportExportScoreRank.DetailContent
{
    [FISCA.Permission.FeatureCode(
        FeatureCode.StudentSchoolYearEntryRankDetail,
        "排名資料")]
    public partial class UCStudentRank
        : FISCA.Presentation.DetailContent
    {
        private BackgroundWorker _bgWorker;
        private bool _isBusy;
        private List<SchoolYearEntryRankRecord> _records =
            new List<SchoolYearEntryRankRecord>();

        public UCStudentRank()
        {
            InitializeComponent();

            this.Group = "排名資料";

            // 先暫時只能檢視，無法新增及刪除。
            btnAdd.Enabled = btnDelete.Enabled = false;
            btnEdit.Enabled = false;

            lvData.SelectedIndexChanged += lvData_SelectedIndexChanged;
            lvData.MouseDoubleClick += lvData_MouseDoubleClick;
            btnEdit.Click += btnEdit_Click;

            FISCA.Features.TryRegister(
                "RankDetailContent",
                x =>
                {
                    if (this.IsDisposed)
                        return;

                    if (this.InvokeRequired)
                    {
                        this.BeginInvoke(new Action(() =>
                        {
                            if (!this.IsDisposed)
                                ReloadData();
                        }));
                    }
                    else
                    {
                        ReloadData();
                    }
                });

            SetupColumns();

            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
        }

        private void SetupColumns()
        {
            lvData.Columns.Clear();

            AddColumn("學年度", 80);
            AddColumn("成績年級", 80);
            AddColumn("成績類型", 90);
            AddColumn("成績項目", 90);
            AddColumn("排名範圍", 110);
            AddColumn("建立時間", 140);
            AddColumn("建立方式", 90);
        }

        private void AddColumn(string text, int width)
        {
            ColumnHeader column = new ColumnHeader();
            column.Text = text;
            column.Width = width;
            lvData.Columns.Add(column);
        }

        protected override void OnPrimaryKeyChanged(EventArgs e)
        {
            ReloadData();
        }

        private void ReloadData()
        {
            if (string.IsNullOrWhiteSpace(this.PrimaryKey))
            {
                _records = new List<SchoolYearEntryRankRecord>();
                BindListView();
                this.Loading = false;
                return;
            }

            LoadDataAsync();
        }

        private void LoadDataAsync()
        {
            if (_bgWorker.IsBusy)
            {
                _isBusy = true;
                return;
            }

            this.Loading = true;
            _bgWorker.RunWorkerAsync(this.PrimaryKey);
        }

        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            string studentID = e.Argument as string;
            RankOverrideDataAccess dataAccess = new RankOverrideDataAccess();
            e.Result = dataAccess.GetSchoolYearEntryRanksByStudentID(studentID);
        }

        private void _bgWorker_RunWorkerCompleted(
            object sender,
            RunWorkerCompletedEventArgs e)
        {
            if (_isBusy)
            {
                _isBusy = false;
                LoadDataAsync();
                return;
            }

            try
            {
                if (e.Error != null)
                {
                    _records = new List<SchoolYearEntryRankRecord>();
                    BindListView();

                    try
                    {
                        FISCA.LogAgent.ApplicationLog.Log(
                            "排名資料",
                            "資料項目查詢失敗",
                            BuildErrorLog(e.Error));
                    }
                    catch
                    {
                    }

                    MsgBox.Show(
                        "讀取排名資料失敗：" + e.Error.Message);
                    return;
                }

                _records = e.Result as List<SchoolYearEntryRankRecord>
                    ?? new List<SchoolYearEntryRankRecord>();
                BindListView();
            }
            finally
            {
                this.Loading = false;
            }
        }

        private void BindListView()
        {
            lvData.BeginUpdate();
            try
            {
                lvData.Items.Clear();

                foreach (SchoolYearEntryRankRecord record in _records)
                {
                    ListViewItem item = new ListViewItem(
                        record.SchoolYear.HasValue
                            ? record.SchoolYear.Value.ToString()
                            : string.Empty);

                    item.SubItems.Add(
                        record.GradeYear.HasValue
                            ? record.GradeYear.Value.ToString()
                            : string.Empty);

                    item.SubItems.Add(record.ScoreType ?? string.Empty);
                    item.SubItems.Add(record.ScoreItem ?? string.Empty);
                    item.SubItems.Add(record.RankType ?? string.Empty);
                    item.SubItems.Add(record.CreateTime ?? string.Empty);
                    item.SubItems.Add(record.CreateMethod ?? string.Empty);
                    item.Tag = record;
                    lvData.Items.Add(item);
                }
            }
            finally
            {
                lvData.EndUpdate();
            }
        }

        private void lvData_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnEdit.Enabled = lvData.SelectedItems.Count == 1;
        }

        private void lvData_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            OpenDetailForm();
        }

        private void btnEdit_Click(object sender, EventArgs e)
        {
            OpenDetailForm();
        }

        private void OpenDetailForm()
        {
            if (lvData.SelectedItems.Count != 1)
                return;

            SchoolYearEntryRankRecord record =
                lvData.SelectedItems[0].Tag as SchoolYearEntryRankRecord;

            if (record == null)
                return;

            StudentRankDetailContext context = new StudentRankDetailContext
            {
                StudentId = this.PrimaryKey,
                SchoolYear = record.SchoolYear,
                GradeYear = record.GradeYear,
                ScoreType = record.ScoreType,
                ScoreItem = record.ScoreItem,
                CreateTime = record.CreateTime,
                CreateMethod = record.CreateMethod
            };

            using (frmStudentRankDetail form = new frmStudentRankDetail(context))
            {
                if (form.ShowDialog() == DialogResult.OK)
                    ReloadData();
            }
        }

        private string BuildErrorLog(Exception ex)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("排名資料查詢失敗");
            sb.AppendLine("目前使用者：" + DSAServices.UserAccount);
            sb.AppendLine("學生系統編號：" + this.PrimaryKey);
            sb.AppendLine(ex.Message);
            sb.AppendLine(ex.StackTrace);
            return sb.ToString();
        }
    }
}
