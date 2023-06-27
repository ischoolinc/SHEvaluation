using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using FISCA.Data;
using System.Xml.Linq;

namespace SmartSchool.Evaluation.Content.ScoreEditor
{
    public partial class frmPickSubject : BaseForm
    {
        //SubjectInfo pickSubject = null;
        List<SubjectInfo> pickSubjects;
        string StudentID = "";

        // 學生課程規畫表 XML
        XElement GraduationPlanXml = null;

        // 課程規劃表名稱
        string GraduationPlanName = "";

        bool isNewRow = false;

        // 已經有的科目名稱與級別
        List<string> hasSubjectNameAndLevel = new List<string>();

        public frmPickSubject()
        {
            InitializeComponent();
            pickSubjects = new List<SubjectInfo>();
        }

        private void frmPickSubject_Load(object sender, EventArgs e)
        {
            // 取得學生課程規畫XML
            LoadGraduationPlanXml();
            LoadXMLToDataGridView();
        }

        // 取得選取的科目
        public List<SubjectInfo> GetPickSubjects()
        {
            return pickSubjects;
        }

        // 設定學生編號
        public void SetStudentID(string id)
        {
            StudentID = id;
        }

        // 讀取學生課程規劃 XML 
        private void LoadGraduationPlanXml()
        {
            QueryHelper qh = new QueryHelper();
            string sql = string.Format(@"    
                WITH ROW AS (
                    SELECT
                        {0} AS student_id
                ),
                target_student AS(
                    SELECT
                        ROW.student_id,
                        graduation_plan.id AS graduation_plan_id
                    FROM
                        ROW
                        INNER JOIN student
                            ON ROW.student_id = student.id
                        INNER JOIN class
                            ON student.ref_class_id = class.id
                        INNER JOIN graduation_plan 
                            ON graduation_plan.id = COALESCE(
                            student.ref_graduation_plan_id,
                            class.ref_graduation_plan_id
                        )
                ),
                graduation_plan_expand AS(
                    SELECT
                        graduation_plan.name,
                        content
                    FROM
                        graduation_plan
                        INNER JOIN target_student 
                            ON graduation_plan.id = target_student.graduation_plan_id
                )
                SELECT
                    *
                FROM
                    graduation_plan_expand
            ", StudentID);

            DataTable dt = qh.Select(sql);
            if (dt.Rows.Count > 0)
            {
                try
                {

                    // 取得課程規劃表名稱
                    GraduationPlanName = dt.Rows[0]["name"] + "";

                    // 讀取課程規劃 XML
                    GraduationPlanXml = XElement.Parse(dt.Rows[0]["content"] + "");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
                this.labelGPName.Text = GraduationPlanName;
            }
        }

        // 將 XML 資料填入畫面 DataGridView
        private void LoadXMLToDataGridView()
        {
            if (GraduationPlanXml != null)
            {
                // 將選取課程規畫表用RowIndex切成各個科目
                Dictionary<string, List<XElement>> dataDict = new Dictionary<string, List<XElement>>();
                foreach (XElement elm in GraduationPlanXml.Elements("Subject"))
                {
                    string idx = elm.Element("Grouping").Attribute("RowIndex").Value;

                    if (!dataDict.ContainsKey(idx))
                        dataDict.Add(idx, new List<XElement>());

                    dataDict[idx].Add(elm);
                }

                // 填入格子
                dgvMain.Rows.Clear();

                // 填入資料
                foreach (string idx in dataDict.Keys)
                {
                    int rowIdx = dgvMain.Rows.Add();
                    XElement firstElm = null;
                    if (dataDict[idx].Count > 0)
                    {
                        firstElm = dataDict[idx][0];
                    }

                    // 將資料存入 Tag
                    dgvMain.Rows[rowIdx].Tag = firstElm;

                    dgvMain.Rows[rowIdx].Cells[colMainDomainName.Index].Value = firstElm.Attribute("Domain").Value;
                    dgvMain.Rows[rowIdx].Cells[colMainEntry.Index].Value = firstElm.Attribute("Entry").Value;
                    dgvMain.Rows[rowIdx].Cells[colMainSubjectName.Index].Value = firstElm.Attribute("SubjectName").Value;

                    if (firstElm.Attribute("RequiredBy").Value == "部訂")
                    {
                        dgvMain.Rows[rowIdx].Cells[colMainRequiredBy.Index].Value = "部定";
                    }
                    else
                        dgvMain.Rows[rowIdx].Cells[colMainRequiredBy.Index].Value = firstElm.Attribute("RequiredBy").Value;

                    dgvMain.Rows[rowIdx].Cells[colMainRequired.Index].Value = firstElm.Attribute("Required").Value;


                    // 不計學分
                    if (firstElm.Attribute("NotIncludedInCredit").Value.ToUpper() == "TRUE")
                        dgvMain.Rows[rowIdx].Cells["不計學分"].Value = "是";
                    else
                        dgvMain.Rows[rowIdx].Cells["不計學分"].Value = "否";

                    // 不需評分
                    if (firstElm.Attribute("NotIncludedInCalc").Value.ToUpper() == "TRUE")
                        dgvMain.Rows[rowIdx].Cells["不需評分"].Value = "是";
                    else
                        dgvMain.Rows[rowIdx].Cells["不需評分"].Value = "否";

                    // 課程代碼
                    dgvMain.Rows[rowIdx].Cells["課程代碼"].Value = firstElm.Attribute("課程代碼").Value;

                    // 指定學年科目名稱
                    if (firstElm.Attribute("指定學年科目名稱") != null)
                    {
                        dgvMain.Rows[rowIdx].Cells[colMainSchoolYearGroupName.Index].Value = firstElm.Attribute("指定學年科目名稱").Value;
                    }

                    // 報部科目名稱
                    if (firstElm.Attribute("報部科目名稱") != null)
                    {
                        dgvMain.Rows[rowIdx].Cells[colMainOfficialSubjectName.Index].Value = firstElm.Attribute("報部科目名稱").Value;
                    }

                    // 解析學分
                    foreach (XElement elmD in dataDict[idx])
                    {
                        try
                        {
                            if (elmD.Attribute("GradeYear").Value == "1" && elmD.Attribute("Semester").Value == "1")
                            {
                                dgvMain.Rows[rowIdx].Cells[colMain1_1.Index].Tag = elmD;
                                dgvMain.Rows[rowIdx].Cells[colMain1_1.Index].Value = elmD.Attribute("Credit").Value;
                                if (CheckHasSubjectNameLevel(elmD))
                                    dgvMain.Rows[rowIdx].Cells[colMain1_1.Index].Style.BackColor = Color.LightGray;

                            }

                            if (elmD.Attribute("GradeYear").Value == "1" && elmD.Attribute("Semester").Value == "2")
                            {
                                dgvMain.Rows[rowIdx].Cells[colMain1_2.Index].Tag = elmD;
                                dgvMain.Rows[rowIdx].Cells[colMain1_2.Index].Value = elmD.Attribute("Credit").Value;
                                if (CheckHasSubjectNameLevel(elmD))
                                    dgvMain.Rows[rowIdx].Cells[colMain1_2.Index].Style.BackColor = Color.LightGray;
                            }

                            if (elmD.Attribute("GradeYear").Value == "2" && elmD.Attribute("Semester").Value == "1")
                            {
                                dgvMain.Rows[rowIdx].Cells[colMain2_1.Index].Tag = elmD;
                                dgvMain.Rows[rowIdx].Cells[colMain2_1.Index].Value = elmD.Attribute("Credit").Value;
                                if (CheckHasSubjectNameLevel(elmD))
                                    dgvMain.Rows[rowIdx].Cells[colMain2_1.Index].Style.BackColor = Color.LightGray;
                            }



                            if (elmD.Attribute("GradeYear").Value == "2" && elmD.Attribute("Semester").Value == "2")
                            {
                                dgvMain.Rows[rowIdx].Cells[colMain2_2.Index].Tag = elmD;
                                dgvMain.Rows[rowIdx].Cells[colMain2_2.Index].Value = elmD.Attribute("Credit").Value;
                                if (CheckHasSubjectNameLevel(elmD))
                                    dgvMain.Rows[rowIdx].Cells[colMain2_2.Index].Style.BackColor = Color.LightGray;
                            }



                            if (elmD.Attribute("GradeYear").Value == "3" && elmD.Attribute("Semester").Value == "1")
                            {
                                dgvMain.Rows[rowIdx].Cells[colMain3_1.Index].Tag = elmD;
                                dgvMain.Rows[rowIdx].Cells[colMain3_1.Index].Value = elmD.Attribute("Credit").Value;
                                if (CheckHasSubjectNameLevel(elmD))
                                    dgvMain.Rows[rowIdx].Cells[colMain3_1.Index].Style.BackColor = Color.LightGray;

                            }

                            if (elmD.Attribute("GradeYear").Value == "3" && elmD.Attribute("Semester").Value == "2")
                            {
                                dgvMain.Rows[rowIdx].Cells[colMain3_2.Index].Tag = elmD;
                                dgvMain.Rows[rowIdx].Cells[colMain3_2.Index].Value = elmD.Attribute("Credit").Value;
                                if (CheckHasSubjectNameLevel(elmD))
                                    dgvMain.Rows[rowIdx].Cells[colMain3_2.Index].Style.BackColor = Color.LightGray;
                            }

                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }


                    // 沒有學分呈現淺灰
                    foreach (DataGridViewCell cell in dgvMain.Rows[rowIdx].Cells)
                    {
                        if (cell.ColumnIndex >= 5 && cell.ColumnIndex <= 10)
                        {
                            if (string.IsNullOrEmpty("" + cell.Value))
                                cell.Style.BackColor = Color.LightGray;
                        }
                    }
                }
            }
        }

        // 檢查科目名稱與級別是否存在
        private bool CheckHasSubjectNameLevel(XElement elm)
        {
            bool value = false;
            if (elm.Attribute("SubjectName") != null && elm.Attribute("Level") != null)
            {
                if (hasSubjectNameAndLevel.Contains(elm.Attribute("SubjectName").Value + "_" + elm.Attribute("Level").Value))
                    value = true;
            }

            return value;
        }

        public void SetIsNewRow(bool value)
        {
            isNewRow = value;
        }


        private void dgvMain_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // 選學分範圍
            if (e.RowIndex > -1 && (e.ColumnIndex >= 5 && e.ColumnIndex <= 10))
            {
                // Console.WriteLine(e.ColumnIndex);

                if (!isNewRow)
                {
                    //只能選一筆
                    foreach (DataGridViewRow row in dgvMain.Rows)
                    {
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.ColumnIndex >= 5 && cell.ColumnIndex <= 10)
                            {
                                if (cell.ColumnIndex != e.ColumnIndex)
                                {
                                    if (cell.Style.BackColor != Color.LightGray)
                                        cell.Style.BackColor = Color.White;
                                }
                            }
                        }
                    }

                }

                // 不能選跳過
                if (dgvMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor == Color.LightGray)
                    return;

                if (dgvMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor == Color.Yellow)
                {
                    dgvMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.White;
                    return;
                }

                if (dgvMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor.IsEmpty || dgvMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor == Color.White)
                {
                    dgvMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor = Color.Yellow;
                    return;
                }

            }

        }

        // 設定已有科目名稱級別
        public void SetHasSubjectNameAndLevel(List<string> value)
        {
            hasSubjectNameAndLevel = value;
        }

        private void btnAddSubject_Click(object sender, EventArgs e)
        {
            pickSubjects.Clear();
            // 加入科目
            foreach (DataGridViewRow row in dgvMain.Rows)
            {
                if (row.IsNewRow)
                    continue;

                foreach (DataGridViewCell cell in row.Cells)
                {
                    if (cell.ColumnIndex >= 5 && cell.ColumnIndex <= 10)
                    {
                        if (cell.Style.BackColor == Color.Yellow)
                        {
                            if (cell.Value != null)
                            {

                                SubjectInfo ss = new SubjectInfo();
                                ss.SubjectName = dgvMain.Rows[cell.RowIndex].Cells[0].Value + "";
                                ss.Domain = dgvMain.Rows[cell.RowIndex].Cells[1].Value + "";
                                ss.Entry = dgvMain.Rows[cell.RowIndex].Cells[2].Value + "";
                                ss.RequiredBy = dgvMain.Rows[cell.RowIndex].Cells[3].Value + "";
                                ss.Required = dgvMain.Rows[cell.RowIndex].Cells[4].Value + "";
                                ss.Credit = dgvMain.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Value + "";
                                ss.CourseCode = dgvMain.Rows[cell.RowIndex].Cells["課程代碼"].Value + "";
                                ss.NotIncludedInCalc = dgvMain.Rows[cell.RowIndex].Cells["不需評分"].Value + "";
                                ss.NotIncludedInCredit = dgvMain.Rows[cell.RowIndex].Cells["不計學分"].Value + "";

                                // 指定學年科目名稱
                                if (dgvMain.Rows[cell.RowIndex].Cells[colMainSchoolYearGroupName.Index].Value != null)
                                {
                                    ss.SchoolYearSubjectName = dgvMain.Rows[cell.RowIndex].Cells[colMainSchoolYearGroupName.Index].Value + "";
                                }

                                // 報部科目名稱
                                if (dgvMain.Rows[cell.RowIndex].Cells[colMainOfficialSubjectName.Index].Value != null)
                                {
                                    ss.DeptSubjectName = dgvMain.Rows[cell.RowIndex].Cells[colMainOfficialSubjectName.Index].Value + "";
                                }

                                // 解析級別
                                XElement elm = dgvMain.Rows[cell.RowIndex].Cells[cell.ColumnIndex].Tag as XElement;
                                if (elm != null)
                                {
                                    if (elm.Attribute("Level") != null)
                                        ss.SubjectLevel = elm.Attribute("Level").Value + "";
                                }


                                pickSubjects.Add(ss);
                            }
                        }
                    }
                }
            }

            Console.WriteLine(pickSubjects.Count);
            this.DialogResult = DialogResult.OK;
        }
    }
}
