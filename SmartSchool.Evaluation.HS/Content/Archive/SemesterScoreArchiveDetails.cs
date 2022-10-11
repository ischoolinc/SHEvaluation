using System;
using SmartSchool.Common;
using SmartSchool.StudentRelated;
using FISCA.Data;
using System.Data;

namespace SmartSchool.Evaluation.Content.ScoreEditor
{
    public partial class SemesterScoreArchiveDetails : BaseForm
    {
        string _StudentID;
        string _Uid;
        string _RefEntryUid;


        public SemesterScoreArchiveDetails(string refStudentID, string uid, string refEntryUid)
        {
            InitializeComponent();
            _StudentID = refStudentID;
            _Uid = uid;
            _RefEntryUid = refEntryUid;
            dataGridView_Archive.Columns["ColCode"].Visible = false;
            labelX10.Text = Student.Instance.Items[_StudentID].ClassName + (Student.Instance.Items[_StudentID].ClassName != "" && Student.Instance.Items[_StudentID].SeatNo != "" ? " " + Student.Instance.Items[_StudentID].SeatNo : "") +
    " " + Student.Instance.Items[_StudentID].Name + (Student.Instance.Items[_StudentID].StudentNumber == "" ? "" : " (" + Student.Instance.Items[_StudentID].StudentNumber + ")");
        }

        private void SemesterScoreArchive_Load(object sender, EventArgs e)
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
                    textBoxX2.Text = dataRow["school_year"].ToString();
                    textBoxX15.Text = dataRow["semester"].ToString();
                    labelX9.Text = "修課年級：" + dataRow["grade_year"].ToString();
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
                dataGridView_Archive.RowCount = dtSubj.Rows.Count;

                foreach (DataRow dr in dtSubj.Rows)
                {
                    dataGridView_Archive.Rows[i].Cells[0].Value = dr["領域"].ToString();
                    dataGridView_Archive.Rows[i].Cells[1].Value = dr["分項類別"].ToString();
                    dataGridView_Archive.Rows[i].Cells[2].Value = dr["科目"].ToString();
                    dataGridView_Archive.Rows[i].Cells[3].Value = dr["科目級別"].ToString();
                    dataGridView_Archive.Rows[i].Cells[4].Value = dr["學分數"].ToString();
                    dataGridView_Archive.Rows[i].Cells[5].Value = dr["校部訂"].ToString() == "部訂" ? "部定": dr["校部訂"].ToString();
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
                    i++;
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show("取得學生學期科目成績封存資訊錯誤。");
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
