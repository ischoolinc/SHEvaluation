using System;
using System.Collections.Generic;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using SHImportExportScoreRank.DAO;

namespace SHImportExportScoreRank.UIForm
{
    public partial class frmStudentRankDetail : BaseForm
    {
        private readonly StudentRankDetailContext _context;
        private readonly RankOverrideDataAccess _dataAccess;
        private StudentRankDetailData _detailData;

        public frmStudentRankDetail(StudentRankDetailContext context)
            : this(context, new RankOverrideDataAccess())
        {
        }

        internal frmStudentRankDetail(
            StudentRankDetailContext context,
            RankOverrideDataAccess dataAccess)
        {
            InitializeComponent();

            _context = context ?? new StudentRankDetailContext();
            _dataAccess = dataAccess ?? new RankOverrideDataAccess();
            SetupGridColumns();
        }

        private void frmStudentRankDetail_Load(object sender, EventArgs e)
        {
            LoadData();
        }

        private void SetupGridColumns()
        {
            dgvRankDetail.AutoGenerateColumns = false;
            dgvRankDetail.Columns.Clear();

            AddTextColumn("colScoreCategory", "成績類別", "ScoreCategory", 90);
            AddTextColumn("colRankMethod", "排名方式", "RankMethod", 70);
            AddTextColumn("colScore", "排名分數", "Score", 70);
            AddTextColumn("colRankType", "排名範圍", "RankType", 80);
            AddTextColumn("colRankName", "母群", "RankName", 90);
            AddTextColumn("colRank", "排名", "RankDisplay", 70);
            AddTextColumn("colPR", "PR值", "PR", 60);
            AddTextColumn("colPercentage", "百分比", "Percentage", 70);
        }

        private void AddTextColumn(
            string name,
            string header,
            string dataProperty,
            int fillWeight)
        {
            DataGridViewTextBoxColumn column = new DataGridViewTextBoxColumn
            {
                Name = name,
                HeaderText = header,
                DataPropertyName = dataProperty,
                ReadOnly = true,
                FillWeight = fillWeight
            };

            dgvRankDetail.Columns.Add(column);
        }

        private void LoadData()
        {
            try
            {
                _detailData = _dataAccess.GetStudentRankDetail(_context);
            }
            catch (Exception ex)
            {
                _detailData = new StudentRankDetailData
                {
                    Items = new List<StudentRankDetailItem>()
                };

                ApplyDetailHeaderFromContext(_detailData);

                MsgBox.Show("讀取排名資料失敗：" + ex.Message);
            }

            BindHeader();
            BindDetail();

            if (_detailData == null ||
                _detailData.Items == null ||
                _detailData.Items.Count == 0)
            {
                MsgBox.Show("查無排名明細資料。");
            }
        }

        private void ApplyDetailHeaderFromContext(StudentRankDetailData data)
        {
            if (data == null || _context == null)
                return;

            data.SchoolYear = _context.SchoolYear;
            data.ScoreType = _context.ScoreType ?? string.Empty;
            data.ScoreItem = _context.ScoreItem ?? string.Empty;
            data.CreateType = _context.CreateMethod ?? string.Empty;
            data.CreateTime = _context.CreateTime ?? string.Empty;
            data.BatchName = string.Empty;
            data.Items = data.Items ?? new List<StudentRankDetailItem>();
        }

        private void BindHeader()
        {
            if (_detailData == null)
                return;

            lblSchoolYearValue.Text = _detailData.SchoolYear.HasValue
                ? _detailData.SchoolYear.Value.ToString()
                : string.Empty;
            lblScoreTypeValue.Text = _detailData.ScoreType ?? string.Empty;
            lblScoreItemValue.Text = _detailData.ScoreItem ?? string.Empty;
            lblCreateTypeValue.Text = _detailData.CreateType ?? string.Empty;
            lblCreateTimeValue.Text = _detailData.CreateTime ?? string.Empty;
            lblBatchNameValue.Text = _detailData.BatchName ?? string.Empty;
        }

        private void BindDetail()
        {
            List<StudentRankDetailItem> items =
                _detailData != null && _detailData.Items != null
                    ? _detailData.Items
                    : new List<StudentRankDetailItem>();

            dgvRankDetail.DataSource = null;
            dgvRankDetail.DataSource = items;
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            MsgBox.Show("目前資料庫尚未提供可修改的排名方式欄位，無法儲存。");
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
