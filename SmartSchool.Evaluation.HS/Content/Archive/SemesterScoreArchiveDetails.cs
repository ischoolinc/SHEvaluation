using FISCA.Data;
using SmartSchool.Common;
using SmartSchool.StudentRelated;
using System;
using System.Data;
using System.Linq;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SmartSchool.Evaluation.Content.ScoreEditor
{
    public partial class SemesterScoreArchiveDetails : BaseForm
    {
        string _StudentID;
        string _Uid;
        string _RefEntryUid;

        // 學年度
        string _SchoolYear;
        // 學期
        string _Semester;
        // 修課年級
        string _GradeYear;

        public SemesterScoreArchiveDetails(string refStudentID, string uid, string refEntryUid)
        {
            InitializeComponent();
            _StudentID = refStudentID;
            _Uid = uid;

            //如果是新增 學年度學期可以設定
            if (string.IsNullOrEmpty(_Uid))
            {               
                txtSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
                txtSemester.Text = K12.Data.School.DefaultSemester;


                // 讓所有 TextBox 清空
                textBoxX1.Text = "";
                textBoxX3.Text = "";
                textBoxX4.Text = "";
                textBoxX5.Text = "";
                textBoxX6.Text = "";
                textBoxX7.Text = "";
                textBoxX8.Text = "";
                textBoxX9.Text = "";
                textBoxX10.Text = "";
                textBoxX11.Text = "";
                textBoxX12.Text = "";
                textBoxX13.Text = "";
                // ...所有分項成績 textbox 清空

                // 讓 dataGridView_Archive 全部清空
                dataGridView_Archive.Rows.Clear();

            }            

            _RefEntryUid = refEntryUid;
            dataGridView_Archive.Columns["ColCode"].Visible = false;
            labelX10.Text = Student.Instance.Items[_StudentID].ClassName + (Student.Instance.Items[_StudentID].ClassName != "" && Student.Instance.Items[_StudentID].SeatNo != "" ? " " + Student.Instance.Items[_StudentID].SeatNo : "") +
    " " + Student.Instance.Items[_StudentID].Name + (Student.Instance.Items[_StudentID].StudentNumber == "" ? "" : " (" + Student.Instance.Items[_StudentID].StudentNumber + ")");
        }

        private void SemesterScoreArchive_Load(object sender, EventArgs e)
        {
            if (!string.IsNullOrEmpty(_Uid) && !string.IsNullOrEmpty(_RefEntryUid))
            {
                QueryHelper qh = new QueryHelper();
                string sqlEntry = @"SELECT
uid
	, ref_student_id
	,school_year
	, semester
    , grade_year
	, array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''學業'']/@成績', xmlparse(content score_info)), '')::text AS 學業成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''體育'']/@成績', xmlparse(content score_info)), '')::text AS 體育成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''健康與護理'']/@成績', xmlparse(content score_info)), '')::text AS 健康與護理成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''國防通識'']/@成績', xmlparse(content score_info)), '')::text AS 國防通識成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''實習科目'']/@成績', xmlparse(content score_info)), '')::text AS 實習科目成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''專業科目'']/@成績', xmlparse(content score_info)), '')::text AS 專業科目成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''學業(原始)'']/@成績', xmlparse(content score_info)), '')::text AS 原始學業成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''體育(原始)'']/@成績', xmlparse(content score_info)), '')::text AS 原始體育成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''健康與護理(原始)'']/@成績', xmlparse(content score_info)), '')::text AS 原始健康與護理成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''國防通識(原始)'']/@成績', xmlparse(content score_info)), '')::text AS 原始國防通識成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''實習科目(原始)'']/@成績', xmlparse(content score_info)), '')::text AS 原始實習科目成績
    , array_to_string(xpath('//SemesterEntryScore/Entry[@分項=''專業科目(原始)'']/@成績', xmlparse(content score_info)), '')::text AS 原始專業科目成績
    , last_update
FROM
	$semester_entry_score_archive
WHERE ref_student_id = {0} AND uid = {1}";

                sqlEntry = string.Format(sqlEntry, _StudentID, _RefEntryUid);
                try
                {
                    if (_RefEntryUid != "0")
                    {
                        DataTable dtEntry = qh.Select(sqlEntry);
                        DataRow dataRow = dtEntry.Rows[0];
                        labelX19.Text = "封存時間：" + dataRow["last_update"].ToString();
                        txtSchoolYear.Text = dataRow["school_year"].ToString();
                        txtSemester.Text = dataRow["semester"].ToString();
                        txtGradeYear.Text = dataRow["grade_year"].ToString();
                        textBoxX1.Text = dataRow["學業成績"].ToString();
                        textBoxX7.Text = dataRow["專業科目成績"].ToString();
                        textBoxX6.Text = dataRow["實習科目成績"].ToString();
                        textBoxX3.Text = dataRow["體育成績"].ToString();
                        textBoxX5.Text = dataRow["健康與護理成績"].ToString();
                        textBoxX4.Text = dataRow["國防通識成績"].ToString();
                        textBoxX12.Text = dataRow["原始體育成績"].ToString();
                        textBoxX11.Text = dataRow["原始健康與護理成績"].ToString();
                        textBoxX13.Text = dataRow["原始學業成績"].ToString();
                        textBoxX10.Text = dataRow["原始國防通識成績"].ToString();
                        textBoxX8.Text = dataRow["原始專業科目成績"].ToString();
                        textBoxX9.Text = dataRow["原始實習科目成績"].ToString();

                        // 儲存時使用
                        _SchoolYear = dataRow["school_year"].ToString();
                        _Semester = dataRow["semester"].ToString();
                        _GradeYear = dataRow["grade_year"].ToString();
                    }
                }
                catch (Exception ex)
                {
                    MsgBox.Show("取得學生學期分項成績封存資訊錯誤。");
                }


                string sqlSubj = @"SELECT
	sems_subj_score_ext.ref_student_id
	, sems_subj_score_ext.grade_year
	, sems_subj_score_ext.semester
	, sems_subj_score_ext.school_year
	, sems_subj_score_ext.last_update
    , array_to_string(xpath('//Subject/@領域', subj_score_ele), '')::text AS 領域
	, array_to_string(xpath('//Subject/@開課分項類別', subj_score_ele), '')::text AS 分項類別
	, array_to_string(xpath('//Subject/@科目', subj_score_ele), '')::text AS 科目
	, array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '')::text AS 科目級別
	, array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')::text AS 學分數
	, array_to_string(xpath('//Subject/@修課校部訂', subj_score_ele), '')::text AS 校部訂
	, array_to_string(xpath('//Subject/@修課必選修', subj_score_ele), '')::text AS 必選修
	, array_to_string(xpath('//Subject/@是否取得學分', subj_score_ele), '')::text AS 取得學分
	, array_to_string(xpath('//Subject/@原始成績', subj_score_ele), '')::text AS 原始成績
	, array_to_string(xpath('//Subject/@補考成績', subj_score_ele), '')::text AS 補考成績
	, array_to_string(xpath('//Subject/@重修成績', subj_score_ele), '')::text AS 重修成績
	, array_to_string(xpath('//Subject/@擇優採計成績', subj_score_ele), '')::text AS 手動調整成績
	, array_to_string(xpath('//Subject/@學年調整成績', subj_score_ele), '')::text AS 學年調整成績
	, array_to_string(xpath('//Subject/@不計學分', subj_score_ele), '')::text AS 不計學分
	, array_to_string(xpath('//Subject/@不需評分', subj_score_ele), '')::text AS 不需評分
	, array_to_string(xpath('//Subject/@註記', subj_score_ele), '')::text AS 註記
	, array_to_string(xpath('//Subject/@修課及格標準', subj_score_ele), '')::text AS 修課及格標準
	, array_to_string(xpath('//Subject/@修課補考標準', subj_score_ele), '')::text AS 修課補考標準
	, array_to_string(xpath('//Subject/@修課直接指定總成績', subj_score_ele), '')::text AS 修課直接指定總成績
	, array_to_string(xpath('//Subject/@修課備註', subj_score_ele), '')::text AS 修課備註
	, array_to_string(xpath('//Subject/@修課科目代碼', subj_score_ele), '')::text AS 修課科目代碼
	, array_to_string(xpath('//Subject/@是否補修成績', subj_score_ele), '')::text AS 是否補修成績
	, array_to_string(xpath('//Subject/@重修學年度', subj_score_ele), '')::text AS 重修學年度
	, array_to_string(xpath('//Subject/@重修學期', subj_score_ele), '')::text AS 重修學期 
	, array_to_string(xpath('//Subject/@補修學年度', subj_score_ele), '')::text AS 補修學年度
	, array_to_string(xpath('//Subject/@補修學期', subj_score_ele), '')::text AS 補修學期 
	, array_to_string(xpath('//Subject/@免修', subj_score_ele), '')::text AS 免修
	, array_to_string(xpath('//Subject/@抵免', subj_score_ele), '')::text AS 抵免
	, array_to_string(xpath('//Subject/@指定學年科目名稱', subj_score_ele), '')::text AS 指定學年科目名稱
FROM (
		SELECT 
			$semester_subject_score_archive.*
			, 	unnest(xpath('//SemesterSubjectScoreInfo/Subject', xmlparse(content score_info))) as subj_score_ele
		FROM 
			$semester_subject_score_archive
		WHERE ref_student_id = {0} AND uid={1} 
	) as sems_subj_score_ext";
                sqlSubj = string.Format(sqlSubj, _StudentID, _Uid);

                try
                {
                    DataTable dtSubj = qh.Select(sqlSubj);
                    //dataGridView_Archive.DataSource = dt;
                    if (dtSubj == null || dtSubj.Rows.Count == 0)
                    {
                        dataGridView_Archive.Rows.Clear();
                        return;
                    }

                    int i = 0;
                    //dataGridView_Archive.RowCount = dtSubj.Rows.Count;

                    dataGridView_Archive.Rows.Clear();

                    foreach (DataRow dr in dtSubj.Rows)
                    {
                        i = dataGridView_Archive.Rows.Add();
                        dataGridView_Archive.Rows[i].Cells[0].Value = dr["領域"].ToString();
                        dataGridView_Archive.Rows[i].Cells[1].Value = dr["分項類別"].ToString();
                        dataGridView_Archive.Rows[i].Cells[2].Value = dr["科目"].ToString();
                        dataGridView_Archive.Rows[i].Cells[3].Value = dr["科目級別"].ToString();
                        dataGridView_Archive.Rows[i].Cells[4].Value = dr["學分數"].ToString();
                        dataGridView_Archive.Rows[i].Cells[5].Value = dr["校部訂"].ToString() == "部訂" ? "部定" : dr["校部訂"].ToString();
                        dataGridView_Archive.Rows[i].Cells[6].Value = dr["必選修"].ToString();
                        dataGridView_Archive.Rows[i].Cells[7].Value = dr["取得學分"].ToString() == "是" ? "是" : "否";
                        dataGridView_Archive.Rows[i].Cells[8].Value = dr["原始成績"].ToString();
                        dataGridView_Archive.Rows[i].Cells[9].Value = dr["補考成績"].ToString();
                        dataGridView_Archive.Rows[i].Cells[10].Value = dr["重修成績"].ToString();
                        dataGridView_Archive.Rows[i].Cells[11].Value = dr["手動調整成績"].ToString();
                        dataGridView_Archive.Rows[i].Cells[12].Value = dr["學年調整成績"].ToString();
                        dataGridView_Archive.Rows[i].Cells[13].Value = dr["不計學分"];
                        dataGridView_Archive.Rows[i].Cells[14].Value = dr["不需評分"];
                        dataGridView_Archive.Rows[i].Cells[15].Value = dr["註記"].ToString();
                        dataGridView_Archive.Rows[i].Cells[16].Value = dr["修課及格標準"].ToString();
                        dataGridView_Archive.Rows[i].Cells[17].Value = dr["修課補考標準"].ToString();
                        dataGridView_Archive.Rows[i].Cells[18].Value = dr["修課直接指定總成績"].ToString();
                        dataGridView_Archive.Rows[i].Cells[19].Value = dr["修課備註"].ToString();
                        dataGridView_Archive.Rows[i].Cells[20].Value = dr["修課科目代碼"].ToString();
                        dataGridView_Archive.Rows[i].Cells[21].Value = dr["是否補修成績"].ToString() == "是" ? "是" : "否";
                        dataGridView_Archive.Rows[i].Cells[22].Value = dr["補修學年度"].ToString();
                        dataGridView_Archive.Rows[i].Cells[23].Value = dr["補修學期"].ToString();
                        dataGridView_Archive.Rows[i].Cells[24].Value = dr["重修學年度"].ToString();
                        dataGridView_Archive.Rows[i].Cells[25].Value = dr["重修學期"].ToString();
                        dataGridView_Archive.Rows[i].Cells[26].Value = dr["免修"].ToString() == "是" ? "是" : "否";
                        dataGridView_Archive.Rows[i].Cells[27].Value = dr["抵免"].ToString() == "是" ? "是" : "否";
                        dataGridView_Archive.Rows[i].Cells[28].Value = dr["指定學年科目名稱"].ToString();

                    }
                }
                catch (Exception ex)
                {
                    MsgBox.Show("取得學生學期科目成績封存資訊錯誤。");
                }
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            btnSave.Enabled = false;

            // 資料檢查
            // 檢查 txtSchoolYer,txtSemester, txtGradeYear 是否為數字
            if (!int.TryParse(txtSchoolYear.Text, out int schoolYear) || !int.TryParse(txtSemester.Text, out int semester) || !int.TryParse(txtGradeYear.Text, out int gradeYear))
            {
                MsgBox.Show("學年度、學期或修課年級必須為數字。");
                btnSave.Enabled = true;
                return;
            }

            // 檢查 dataGridView_Archive 所有 cell，當ErrorText !="" 時，表示有錯誤， return
            foreach (DataGridViewRow row in dataGridView_Archive.Rows)
            {
                if (row.IsNewRow) continue; // 跳過新增列
                // 檢查每一個 cell 的 ErrorText
                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (!string.IsNullOrEmpty(cell.ErrorText))
                    {
                        MsgBox.Show("請修正所有錯誤後再儲存。");
                        btnSave.Enabled = true;
                        return;
                    }
                }
            }

            _SchoolYear = txtSchoolYear.Text;
            _Semester = txtSemester.Text;
            _GradeYear = txtGradeYear.Text;

            // 產生學期分項成績 XML，新增至資料庫 $semester_entry_score_archive
            string ref_sems_entry_uid = "";
            // 分項成績
            XElement SemesterEntryScoreXML = new XElement("SemesterEntryScore",
               new XElement("Entry", new XAttribute("分項", "學業"), new XAttribute("成績", textBoxX1.Text)),
               new XElement("Entry", new XAttribute("分項", "專業科目"), new XAttribute("成績", textBoxX7.Text)),
               new XElement("Entry", new XAttribute("分項", "實習科目"), new XAttribute("成績", textBoxX6.Text)),
               new XElement("Entry", new XAttribute("分項", "體育"), new XAttribute("成績", textBoxX3.Text)),
               new XElement("Entry", new XAttribute("分項", "健康與護理"), new XAttribute("成績", textBoxX5.Text)),
               new XElement("Entry", new XAttribute("分項", "國防通識"), new XAttribute("成績", textBoxX4.Text)),
               new XElement("Entry", new XAttribute("分項", "學業(原始)"), new XAttribute("成績", textBoxX13.Text)),
               new XElement("Entry", new XAttribute("分項", "體育(原始)"), new XAttribute("成績", textBoxX12.Text)),
               new XElement("Entry", new XAttribute("分項", "健康與護理(原始)"), new XAttribute("成績", textBoxX11.Text)),
               new XElement("Entry", new XAttribute("分項", "國防通識(原始)"), new XAttribute("成績", textBoxX10.Text)),
               new XElement("Entry", new XAttribute("分項", "實習科目(原始)"), new XAttribute("成績", textBoxX9.Text)),
               new XElement("Entry", new XAttribute("分項", "專業科目(原始)"), new XAttribute("成績", textBoxX8.Text))
           );


            string InsertSemesterEntryScoreArchive = string.Format(@"
            INSERT INTO
                $semester_entry_score_archive (
                    ref_student_id,
                    school_year,
                    semester,
                    grade_year,
                    score_info
                )
            VALUES
                (
                    {0},
                    {1},
                    {2},
                    {3},
                    '{4}'
                ) RETURNING uid 
            ", _StudentID, _SchoolYear, _Semester, _GradeYear, SemesterEntryScoreXML.ToString());

            // 取得 ref_sems_entry_uid
            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(InsertSemesterEntryScoreArchive);
            if (dt != null && dt.Rows.Count > 0)
            {
                ref_sems_entry_uid = dt.Rows[0][0].ToString();
            }


            //  處理學期科目成績
            XElement SemesterSubjectScoreXML = new XElement("SemesterSubjectScoreInfo");

            foreach (DataGridViewRow row in dataGridView_Archive.Rows)
            {
                if (row.IsNewRow) continue; // 跳過新增列

                // 校部定回寫轉部訂
                string RequirdBy = "校訂";
                if (row.Cells[5].Value != null && row.Cells[5].Value.ToString() == "部定")
                    RequirdBy = "部訂";


                XElement subject = new XElement("Subject",
                    new XAttribute("領域", row.Cells[0].Value?.ToString() ?? ""),
                    new XAttribute("開課分項類別", row.Cells[1].Value?.ToString() ?? ""),
                    new XAttribute("科目", row.Cells[2].Value?.ToString() ?? ""),
                    new XAttribute("科目級別", row.Cells[3].Value?.ToString() ?? ""),
                    new XAttribute("開課學分數", row.Cells[4].Value?.ToString() ?? ""),
                    new XAttribute("修課校部訂", RequirdBy),
                    new XAttribute("修課必選修", row.Cells[6].Value?.ToString() ?? ""),
                    new XAttribute("是否取得學分", row.Cells[7].Value?.ToString() ?? "否"),
                    new XAttribute("原始成績", row.Cells[8].Value?.ToString() ?? ""),
                    new XAttribute("補考成績", row.Cells[9].Value?.ToString() ?? ""),
                    new XAttribute("重修成績", row.Cells[10].Value?.ToString() ?? ""),
                    new XAttribute("擇優採計成績", row.Cells[11].Value?.ToString() ?? ""),
                    new XAttribute("學年調整成績", row.Cells[12].Value?.ToString() ?? ""),
                    new XAttribute("不計學分", row.Cells[13].Value?.ToString() ?? "否"),
                    new XAttribute("不需評分", row.Cells[14].Value?.ToString() ?? "否"),
                    new XAttribute("註記", row.Cells[15].Value?.ToString() ?? ""),
                    new XAttribute("修課及格標準", row.Cells[16].Value?.ToString() ?? ""),
                    new XAttribute("修課補考標準", row.Cells[17].Value?.ToString() ?? ""),
                    new XAttribute("修課直接指定總成績", row.Cells[18].Value?.ToString() ?? ""),
                    new XAttribute("修課備註", row.Cells[19].Value?.ToString() ?? ""),
                    new XAttribute("修課科目代碼", row.Cells[20].Value?.ToString() ?? ""),
                    new XAttribute("是否補修成績", row.Cells[21].Value?.ToString() ?? "否"),
                    new XAttribute("補修學年度", row.Cells[22].Value?.ToString() ?? ""),
                    new XAttribute("補修學期", row.Cells[23].Value?.ToString() ?? ""),
                    new XAttribute("重修學年度", row.Cells[24].Value?.ToString() ?? ""),
                    new XAttribute("重修學期", row.Cells[25].Value?.ToString() ?? ""),
                    new XAttribute("免修", row.Cells[26].Value?.ToString() ?? "否"),
                    new XAttribute("抵免", row.Cells[27].Value?.ToString() ?? "否"),
                    new XAttribute("指定學年科目名稱", row.Cells[28].Value?.ToString() ?? "")
                );
                SemesterSubjectScoreXML.Add(subject);
            }

            // 更新學期科目成績 XML，新增至資料庫 $semester_subject_score_archive
            string InsertSemesterSubjectScoreArchive = string.Format(@"
            INSERT INTO
                $semester_subject_score_archive (
                    ref_student_id,
                    school_year,
                    semester,
                    grade_year,
                    score_info,
                    ref_sems_entry_uid
                )
            VALUES
                (
                    {0},
                    {1},
                    {2},
                    {3},
                    '{4}',
                    {5}
                ) RETURNING uid;
            ", _StudentID, _SchoolYear, _Semester, _GradeYear, SemesterSubjectScoreXML.ToString(), ref_sems_entry_uid);

            DataTable dtSubj = qh.Select(InsertSemesterSubjectScoreArchive);
            if (dtSubj != null && dtSubj.Rows.Count > 0)
            {
                // _Uid = dtSubj.Rows[0][0].ToString();
            }
            this.DialogResult = DialogResult.OK;

            btnSave.Enabled = true;
        }

        private void dataGridView_Archive_CellValidating(object sender, DataGridViewCellValidatingEventArgs e)
        {
            string input = e.FormattedValue?.ToString().Trim() ?? "";

            // 先清空
            dataGridView_Archive.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "";

            // 整數欄位
            if (new int[] { 3, 4, 22, 23, 24, 25 }.Contains(e.ColumnIndex))
            {
                if (!string.IsNullOrEmpty(input) && !int.TryParse(input, out int _))
                {
                    dataGridView_Archive.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "此欄位必須為整數！";
                }
            }
            // 小數欄位
            else if (new int[] { 8, 9, 10, 11, 12 }.Contains(e.ColumnIndex))
            {
                if (!string.IsNullOrEmpty(input) && !decimal.TryParse(input, out decimal _))
                {
                    dataGridView_Archive.Rows[e.RowIndex].Cells[e.ColumnIndex].ErrorText = "此欄位必須為數字！";
                }
            }
        }

        private void dataGridView_Archive_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Delete)
            {
                foreach (DataGridViewRow row in dataGridView_Archive.SelectedRows)
                {
                    if (!row.IsNewRow)
                        dataGridView_Archive.Rows.Remove(row);
                }
            }
        }
    }
}
