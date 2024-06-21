using System;
using System.Collections.Generic;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Xml;
using DevComponents.DotNetBar.Controls;
using FISCA.DSAUtil;
using SmartSchool.ApplicationLog;
using SmartSchool.Common;
using SmartSchool.Evaluation.GraduationPlan;
using SmartSchool.Feature.Score;
using SmartSchool.StudentRelated;
using SmartSchool.Customization.Data;
using System.Runtime.Versioning;
//using SHCourseGroupCodeDAL;

namespace SmartSchool.Evaluation.Content.ScoreEditor
{
    public partial class SemesterScoreEditor : BaseForm
    {
        string _SubjectScoreID = "";
        string _EntryScoreID1 = "";
        string _EntryScoreID2 = "";
        string _StudentID;
        private Dictionary<Control, string> _entryScoreBase;
        private bool _scoreUpdating = false;
        private Dictionary<Control, ErrorProvider> _errorProviderDictionary = new Dictionary<Control, ErrorProvider>();

        //Log用來比較前後資料差異的變數
        private Dictionary<string, string> _beforeData = new Dictionary<string, string>();
        private Dictionary<string, string> _afterData = new Dictionary<string, string>();
        private DSXmlHelper _beforeXml = new DSXmlHelper("BeforeXml");
        private DSXmlHelper _afterXml = new DSXmlHelper("AfterXml");
        private EntityAction _entityAction = EntityAction.Insert;

        //排名相關
        private SubjectScoreToolTipProvider _subject_rating;
        private EntryScoreToolTipProvider _score_rating;
        private EntryScoreToolTipProvider _moral_rating;
        private const int SubjectColumn = 2, SubjectLevel = 3;

        public SemesterScoreEditor(string refStudentID)
        {
            InitializeComponent();
            _StudentID = refStudentID;
            _beforeXml.AddElement("SubjectGradeYear");
            _beforeXml.AddElement("SubjectCollection");
            _beforeXml.AddElement("EntryCollection");
            _beforeXml.AddElement("LockScore");
            _afterXml.AddElement("SubjectGradeYear");
            _afterXml.AddElement("SubjectCollection");
            _afterXml.AddElement("EntryCollection");
            _afterXml.AddElement("LockScore");
            labelX10.Text = Student.Instance.Items[_StudentID].ClassName + (Student.Instance.Items[_StudentID].ClassName != "" && Student.Instance.Items[_StudentID].SeatNo != "" ? " " + Student.Instance.Items[_StudentID].SeatNo : "") +
                " " + Student.Instance.Items[_StudentID].Name + (Student.Instance.Items[_StudentID].StudentNumber == "" ? "" : " (" + Student.Instance.Items[_StudentID].StudentNumber + ")");
            //for (int s = 3; s > 0; s--)
            //{
            //    comboBoxEx1.Items.Add(CurrentUser.Instance.SchoolYear - s);
            //}
            for (int i = -3; i <= 3; i++)
            {
                comboBoxEx1.Items.Add(CurrentUser.Instance.SchoolYear - i);
            }
            comboBoxEx2.Items.AddRange(new object[] { "1", "2" });
            cboAttendGradeYear.Items.AddRange(new object[] { "1", "2", "3", "4" });

            ValidateAll();
        }

        private void SemesterScoreEditor_Load(object sender, EventArgs e)
        {

        }
        public SemesterScoreEditor(string schoolYear, string semester, string refStudentID)
        {
            InitializeComponent();
            _StudentID = refStudentID;
            _beforeXml.AddElement("SubjectGradeYear");
            _beforeXml.AddElement("SubjectCollection");
            _beforeXml.AddElement("EntryCollection");
            _beforeXml.AddElement("LockScore");
            _afterXml.AddElement("SubjectGradeYear");
            _afterXml.AddElement("SubjectCollection");
            _afterXml.AddElement("EntryCollection");
            _afterXml.AddElement("LockScore");
            _entityAction = EntityAction.Update;
            labelX10.Text = Student.Instance.Items[_StudentID].ClassName + (Student.Instance.Items[_StudentID].ClassName != "" && Student.Instance.Items[_StudentID].SeatNo != "" ? " " + Student.Instance.Items[_StudentID].SeatNo : "") +
                " " + Student.Instance.Items[_StudentID].Name + (Student.Instance.Items[_StudentID].StudentNumber == "" ? "" : " (" + Student.Instance.Items[_StudentID].StudentNumber + ")");
            this.comboBoxEx1.Text = schoolYear;
            this.comboBoxEx2.Text = semester;
            comboBoxEx1.Enabled = comboBoxEx2.Enabled = false;
            ReLoad(null, null);
            cboAttendGradeYear.Items.AddRange(new object[] { "1", "2", "3", "4" });
            ValidateAll();

            btnSave.Visible = CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            dataGridViewX1.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            cboAttendGradeYear.Enabled = CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            btnCalcScore.Enabled = CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX1.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX2.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX3.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX4.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX5.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX6.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX7.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX8.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX9.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX10.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX11.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX12.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;
            textBoxX13.ReadOnly = !CurrentUser.Acl[SemesterScorePalmerworm.FeatureCode].Editable;

        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void ReLoad(object sender, EventArgs e)
        {
            _scoreUpdating = true;

            _SubjectScoreID = "";
            _EntryScoreID1 = "";
            _EntryScoreID2 = "";

            if (!ValidateSchoolYearSemester())
                return;

            string gradeYear = "";

            #region 科目成績
            dataGridViewX1.Rows.Clear();
            DSXmlHelper _SubjectResponse = QueryScore.GetSemesterSubjectScoreBySemester(true, (comboBoxEx1.Text), (comboBoxEx2.Text), _StudentID);
            if (_SubjectResponse.GetElement("SemesterSubjectScore") != null)
            {
                _SubjectScoreID = _SubjectResponse.GetElement("SemesterSubjectScore").GetAttribute("ID");
                int t;
                if (int.TryParse(_SubjectResponse.GetText("SemesterSubjectScore/GradeYear"), out t))
                    gradeYear = "" + t;


                #region 建立排名相關物件
                DSXmlHelper rating = _SubjectResponse;
                _subject_rating = new SubjectScoreToolTipProvider(rating.GetElement("SemesterSubjectScore/ClassRating"),
                    rating.GetElement("SemesterSubjectScore/DeptRating"),
                    rating.GetElement("SemesterSubjectScore/YearRating"));
                #endregion

                foreach (XmlElement var in _SubjectResponse.GetElements("SemesterSubjectScore/ScoreInfo/SemesterSubjectScoreInfo/Subject"))
                {
                    _beforeXml.AddElement("SubjectCollection", var);
                    DataGridViewRow row = new DataGridViewRow();
                    row.CreateCells(dataGridViewX1,
                        "", // 選科目按鈕
                        var.GetAttribute("領域"), // 領域名稱
                        var.GetAttribute("開課分項類別"),
                        var.GetAttribute("科目"),
                        var.GetAttribute("科目級別"),
                        var.GetAttribute("開課學分數"),
                        var.GetAttribute("修課校部訂") == "部訂" ? "部定" : var.GetAttribute("修課校部訂"),
                        var.GetAttribute("修課必選修"),
                        var.GetAttribute("是否取得學分") == "是",
                        var.GetAttribute("原始成績"),
                        var.GetAttribute("補考成績"),
                        var.GetAttribute("重修成績"),
                        var.GetAttribute("擇優採計成績"),
                        var.GetAttribute("學年調整成績"),
                        var.GetAttribute("不計學分") == "是",
                        var.GetAttribute("不需評分") == "是",
                        var.GetAttribute("註記"),
                        var.GetAttribute("修課及格標準"),
                        var.GetAttribute("修課補考標準"),
                        var.GetAttribute("修課直接指定總成績"),
                        var.GetAttribute("修課備註"),
                        var.GetAttribute("是否補修成績") == "是",
                        var.GetAttribute("補修學年度"),
                        var.GetAttribute("補修學期"),
                        var.GetAttribute("重修學年度"),
                        var.GetAttribute("重修學期"),
                        var.GetAttribute("免修") == "是",
                        var.GetAttribute("抵免") == "是",
                        var.GetAttribute("指定學年科目名稱"),
                        var.GetAttribute("修課科目代碼"),
                        var.GetAttribute("報部科目名稱"),
                        var.GetAttribute("是否重讀") == "是"
                        );
                    row.Cells[SubjectColumn].ToolTipText = GetSubjectScorePlace(row);

                    dataGridViewX1.Rows.Add(row);
                }
                foreach (DataGridViewRow row in dataGridViewX1.Rows)
                {
                    CompareSubjectInfo(row);
                }
            }
            #endregion
            #region 分項成績
            textBoxX1.Text = "";
            textBoxX2.Text = "";
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
            DSXmlHelper _EntryResponse = QueryScore.GetSemesterEntryScoreBySemester(true, (comboBoxEx1.Text), (comboBoxEx2.Text), _StudentID);
            if (_EntryResponse.GetElement("SemesterEntryScore") != null)
            {
                int t;
                if (int.TryParse(_EntryResponse.GetText("SemesterEntryScore/GradeYear"), out t))
                    gradeYear = "" + t;
                foreach (XmlElement var in _EntryResponse.GetElements("SemesterEntryScore"))
                {
                    string id = var.GetAttribute("ID");
                    switch (var.SelectSingleNode("EntryGroup").InnerText)
                    {
                        case "學習":
                            _EntryScoreID1 = id;
                            _score_rating = new EntryScoreToolTipProvider("學業", var); //排名資訊
                            break;
                        case "行為":
                            _EntryScoreID2 = id;
                            _moral_rating = new EntryScoreToolTipProvider("德行", var); //排名資訊
                            break;
                        default:
                            throw new Exception("拎唄謀洗鰓機雷EntryGroup： 　\"" + var.SelectSingleNode("EntryGroup").InnerText + "\"");
                    }
                    foreach (XmlNode score in var.SelectNodes("ScoreInfo/SemesterEntryScore/Entry"))
                    {
                        XmlElement element = (XmlElement)score;
                        _beforeXml.AddElement("EntryCollection", element);

                        #region 依分項填入格子
                        switch (element.GetAttribute("分項"))
                        {
                            case "學業":
                                textBoxX1.Text = element.GetAttribute("成績");

                                _score_rating.SetTooltip(textBoxX1);
                                _score_rating.SetTooltip(labelX3);
                                break;
                            case "體育":
                                textBoxX3.Text = element.GetAttribute("成績");
                                break;
                            case "國防通識":
                                textBoxX4.Text = element.GetAttribute("成績");
                                break;
                            case "健康與護理":
                                textBoxX5.Text = element.GetAttribute("成績");
                                break;
                            case "實習科目":
                                textBoxX6.Text = element.GetAttribute("成績");
                                break;
                            case "專業科目":
                                textBoxX7.Text = element.GetAttribute("成績");
                                break;
                            case "學業(原始)":
                                textBoxX13.Text = element.GetAttribute("成績");
                                break;
                            case "體育(原始)":
                                textBoxX12.Text = element.GetAttribute("成績");
                                break;
                            case "國防通識(原始)":
                                textBoxX10.Text = element.GetAttribute("成績");
                                break;
                            case "健康與護理(原始)":
                                textBoxX11.Text = element.GetAttribute("成績");
                                break;
                            case "實習科目(原始)":
                                textBoxX9.Text = element.GetAttribute("成績");
                                break;
                            case "專業科目(原始)":
                                textBoxX8.Text = element.GetAttribute("成績");
                                break;
                            case "德行":
                                textBoxX2.Text = element.GetAttribute("成績");

                                _moral_rating.SetTooltip(textBoxX2);
                                _moral_rating.SetTooltip(labelX4);

                                bool lockScore = false;
                                bool.TryParse(element.GetAttribute("鎖定"), out lockScore);
                                if (lockScore)
                                {
                                    buttonItem4.Checked = true;
                                }
                                else
                                {
                                    buttonItem3.Checked = true;
                                }
                                _beforeXml.AddElement("LockScore", "Lock", element.GetAttribute("鎖定"));
                                break;
                            default:
                                //throw new Exception("拎唄謀洗鰓機雷分項： " + element.GetAttribute("分項"));
                                break;
                        }
                        #endregion
                    }
                }
            }
            _entryScoreBase = new Dictionary<Control, string>();
            _entryScoreBase.Add(textBoxX1, textBoxX1.Text);
            _entryScoreBase.Add(textBoxX2, textBoxX2.Text);
            _entryScoreBase.Add(textBoxX3, textBoxX3.Text);
            _entryScoreBase.Add(textBoxX4, textBoxX4.Text);
            _entryScoreBase.Add(textBoxX5, textBoxX5.Text);
            _entryScoreBase.Add(textBoxX6, textBoxX6.Text);
            _entryScoreBase.Add(textBoxX7, textBoxX7.Text);
            _entryScoreBase.Add(textBoxX8, textBoxX8.Text);
            _entryScoreBase.Add(textBoxX9, textBoxX9.Text);
            _entryScoreBase.Add(textBoxX10, textBoxX10.Text);
            _entryScoreBase.Add(textBoxX11, textBoxX11.Text);
            _entryScoreBase.Add(textBoxX12, textBoxX12.Text);
            _entryScoreBase.Add(textBoxX13, textBoxX13.Text);

            _scoreUpdating = false;
            #endregion
            int tryint;
            if (int.TryParse(gradeYear, out tryint))
            {
                cboAttendGradeYear.Text = gradeYear;
                _beforeXml.AddElement("SubjectGradeYear", "GradeYear", cboAttendGradeYear.Text);
            }
        }

        private void CompareSubjectInfo(DataGridViewRow row)
        {
            if (!row.IsNewRow)
            {
                #region 當課程資訊定義取自課程規畫表時
                //將自課程規畫表取得的資料顯示於TOOLTIP
                if (GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(_StudentID) != null)
                {
                    #region 比對各項目資料
                    GraduationPlanInfo gplan = GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(_StudentID);

                    // 科目3,級別4
                    GraduationPlanSubject subject = gplan.GetSubjectInfo("" + row.Cells[3].Value, "" + row.Cells[4].Value);

                    int index = 0; // 領域名稱檢查規則，因為舊資料中都會缺少領域欄位，故先不進行檢查，允許空值 --2022.10.05 俊緯
                    //if ("" + row.Cells[index].Value != subject.Domain)
                    //{
                    //    row.Cells[index].ToolTipText = "在課程規劃表 \"" + gplan.Name + "\"中\n值為: " + subject.Domain;
                    //    row.Cells[index].Style.BackColor = Color.Gainsboro;
                    //    row.Cells[index].Style.ForeColor = Color.DimGray;
                    //}
                    //else
                    //{
                    //    row.Cells[index].ToolTipText = "";
                    //    row.Cells[index].Style.BackColor = dataGridViewX1.DefaultCellStyle.BackColor;
                    //    row.Cells[index].Style.ForeColor = dataGridViewX1.DefaultCellStyle.ForeColor;
                    //}

                    //  分項 2
                    index = 2;
                    if ("" + row.Cells[index].Value != subject.Entry)
                    {
                        row.Cells[index].ToolTipText = "在課程規劃表 \"" + gplan.Name + "\"中\n值為: " + subject.Entry;
                        row.Cells[index].Style.BackColor = Color.Gainsboro;
                        row.Cells[index].Style.ForeColor = Color.DimGray;
                    }
                    else
                    {
                        row.Cells[index].ToolTipText = "";
                        row.Cells[index].Style.BackColor = dataGridViewX1.DefaultCellStyle.BackColor;
                        row.Cells[index].Style.ForeColor = dataGridViewX1.DefaultCellStyle.ForeColor;
                    }

                    // 學分 5
                    index = 5;
                    if ("" + row.Cells[index].Value != subject.Credit)
                    {
                        row.Cells[index].ToolTipText = "在課程規劃表 \"" + gplan.Name + "\"中\n值為: " + subject.Credit;
                        row.Cells[index].Style.BackColor = Color.Gainsboro;
                        row.Cells[index].Style.ForeColor = Color.DimGray;
                    }
                    else
                    {
                        row.Cells[index].ToolTipText = "";
                        row.Cells[index].Style.BackColor = dataGridViewX1.DefaultCellStyle.BackColor;
                        row.Cells[index].Style.ForeColor = dataGridViewX1.DefaultCellStyle.ForeColor;
                    }

                    // 校部定 6
                    index = 6;
                    string req = subject.RequiredBy;
                    if (req == "部訂") req = "部定";
                    if ("" + row.Cells[index].Value != req)
                    {
                        row.Cells[index].ToolTipText = "在課程規劃表 \"" + gplan.Name + "\"中\n值為: " + req;
                        row.Cells[index].Style.BackColor = Color.Gainsboro;
                        row.Cells[index].Style.ForeColor = Color.DimGray;
                    }
                    else
                    {
                        row.Cells[index].ToolTipText = "";
                        row.Cells[index].Style.BackColor = dataGridViewX1.DefaultCellStyle.BackColor;
                        row.Cells[index].Style.ForeColor = dataGridViewX1.DefaultCellStyle.ForeColor;
                    }

                    // 必選修 7
                    index = 7;
                    if ("" + row.Cells[index].Value != subject.Required)
                    {
                        row.Cells[index].ToolTipText = "在課程規劃表 \"" + gplan.Name + "\"中\n值為: " + subject.Required;
                        row.Cells[index].Style.BackColor = Color.Gainsboro;
                        row.Cells[index].Style.ForeColor = Color.DimGray;
                    }
                    else
                    {
                        row.Cells[index].ToolTipText = "";
                        row.Cells[index].Style.BackColor = dataGridViewX1.DefaultCellStyle.BackColor;
                        row.Cells[index].Style.ForeColor = dataGridViewX1.DefaultCellStyle.ForeColor;
                    }


                    // 不計學分14
                    index = 14;
                    if ((row.Cells[index].Value != null && (bool)row.Cells[index].Value) != subject.NotIncludedInCredit)
                    {
                        row.Cells[index].ToolTipText = "在課程規劃表 \"" + gplan.Name + "\"中\n值為: " + (subject.NotIncludedInCredit ? "是" : "否");
                        row.Cells[index].Style.BackColor = Color.Gainsboro;
                        row.Cells[index].Style.ForeColor = Color.DimGray;
                    }
                    else
                    {
                        row.Cells[index].ToolTipText = "";
                        row.Cells[index].Style.BackColor = dataGridViewX1.DefaultCellStyle.BackColor;
                        row.Cells[index].Style.ForeColor = dataGridViewX1.DefaultCellStyle.ForeColor;
                    }

                    // 不需評分 15
                    index = 15;
                    if ((row.Cells[index].Value != null && (bool)row.Cells[index].Value) != subject.NotIncludedInCalc)
                    {
                        row.Cells[index].ToolTipText = "在課程規劃表 \"" + gplan.Name + "\"中\n值為: " + (subject.NotIncludedInCalc ? "是" : "否");
                        row.Cells[index].Style.BackColor = Color.Gainsboro;
                        row.Cells[index].Style.ForeColor = Color.DimGray;
                    }
                    else
                    {
                        row.Cells[index].ToolTipText = "";
                        row.Cells[index].Style.BackColor = dataGridViewX1.DefaultCellStyle.BackColor;
                        row.Cells[index].Style.ForeColor = dataGridViewX1.DefaultCellStyle.ForeColor;
                    }
                    #endregion
                }
                else
                {
                    //改變顏色
                    //foreach (int index in new int[] { 0, 1, 4, 5, 6, 13, 14 }) // 領域名稱檢查規則，因為舊資料中都會缺少領域欄位，故先不進行檢查，允許空值 --2022.10.05 俊緯
                    //foreach (int index in new int[] { 1, 4, 5, 6, 13, 14 }) // 2023/5/18，前面多一個功能，所以要往後移一個
                    foreach (int index in new int[] { 2, 5, 6, 7, 14, 15 })
                    {
                        row.Cells[index].ToolTipText = "學生課程規劃表未設定，無法比較與課程規畫表差異。";
                        row.Cells[index].Style.BackColor = Color.Gainsboro;
                        row.Cells[index].Style.ForeColor = Color.DimGray;
                    }
                }
                #endregion
            }
            else
            {
                foreach (DataGridViewCell cell in row.Cells)
                {
                    cell.ToolTipText = "";
                    cell.Style.BackColor = dataGridViewX1.DefaultCellStyle.BackColor;
                    cell.Style.ForeColor = dataGridViewX1.DefaultCellStyle.ForeColor;
                }
            }
        }

        private void comboBoxEx3_TextChanged(object sender, EventArgs e)
        {
            int s = 0;
            if (!int.TryParse(cboAttendGradeYear.Text, out s))
            {
                errorGradeYear.Icon = Properties.Resources.error;
                errorGradeYear.SetError(cboAttendGradeYear, "必須輸入數字");
            }
            else
                errorGradeYear.Clear();
            if (!int.TryParse(cboAttendGradeYear.Text, out s))
            {
                errorGradeYear.Icon = Properties.Resources.error;
                errorGradeYear.SetError(cboAttendGradeYear, "修課年級必須填寫");
            }
            else
                errorGradeYear.Clear();
        }

        private void dataGridViewX1_DataError(object sender, DataGridViewDataErrorEventArgs e)
        {
            e.Cancel = true;
            DataGridViewCell cell = dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];
            string message = "儲存格值：" + cell.Value + "。\n發生錯誤： " + e.Exception.Message + "。";
            /// 2022-01 Cynthia 因分項類別的item移除了體育、國防通識、健康與護理，為了讓舊資料不要出現紅點，故增加一層判斷。
            if (e.ColumnIndex != 1 && cell.Value.ToString() != "體育" && cell.Value.ToString() != "國防通識" && cell.Value.ToString() != "健康與護理")
                if (cell.ErrorText != message)
                {
                    cell.ErrorText = message;
                    dataGridViewX1.UpdateCellErrorText(e.ColumnIndex, e.RowIndex);
                }

            // 2022-01 Cynthia 且為了讓分項類別可以呈現舊資料，只好把分項類別加回去。
            // 若不加回去則會顯示為「學業」。
            DataGridViewComboBoxColumn comboColumn;
            switch (e.ColumnIndex)
            {
                case 1:
                    comboColumn = ((DataGridViewComboBoxColumn)dataGridViewX1.Columns["ColEntry"]);
                    if (!comboColumn.Items.Contains("體育") && cell.Value.ToString() == "體育")
                    {
                        comboColumn.Items.Add("體育");
                    }
                    if (!comboColumn.Items.Contains("國防通識") && cell.Value.ToString() == "國防通識")
                    {
                        comboColumn.Items.Add("國防通識");
                    }
                    if (!comboColumn.Items.Contains("健康與護理") && cell.Value.ToString() == "健康與護理")
                    {
                        comboColumn.Items.Add("健康與護理");
                    }
                    break;
            }
        }

        private void dataGridViewX1_CellEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridViewX1.SelectedCells.Count == 1)
                dataGridViewX1.BeginEdit(true);
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            dataGridViewX1.EndEdit();
            if (!ValidateAll())
            {
                MsgBox.Show("資料有誤，請修正。");
                return;
            }
            #region 新增修改科目成績資料
            XmlElement subjectScoreInfo = CreateSubjectScoreElement();
            if (_SubjectScoreID != "")
            {
                if (subjectScoreInfo.SelectNodes("Subject").Count > 0)
                    EditScore.UpdateSemesterSubjectScore(_SubjectScoreID, cboAttendGradeYear.Text, subjectScoreInfo);
                else
                    RemoveScore.DeleteSemesterSubjectScore(_SubjectScoreID);
            }
            else
            {
                if (subjectScoreInfo.SelectNodes("Subject").Count > 0)
                    AddScore.InsertSemesterSubjectScore(_StudentID, comboBoxEx1.Text, comboBoxEx2.Text, cboAttendGradeYear.Text, subjectScoreInfo);
            }
            #endregion

            _afterXml.AddElement("LockScore", "Lock", "" + buttonItem4.Checked);
            #region 新增修改分項成績資料

            XmlDocument doc = new XmlDocument();
            XmlElement entryScoreInfo1 = doc.CreateElement("SemesterEntryScore");
            XmlElement entryScoreInfo2 = doc.CreateElement("SemesterEntryScore");
            XmlElement entryElement;
            double test = 0;
            #region 行為類
            if (double.TryParse(textBoxX2.Text, out test))
            {
                entryElement = doc.CreateElement("Entry");
                entryElement.SetAttribute("分項", "德行");
                entryElement.SetAttribute("成績", textBoxX2.Text);
                entryElement.SetAttribute("鎖定", "" + buttonItem4.Checked);
                entryScoreInfo1.AppendChild(entryElement);
                _afterXml.AddElement("EntryCollection", entryElement);
                if (_EntryScoreID2 != "")
                {
                    EditScore.UpdateSemesterEntryScore(_EntryScoreID2, cboAttendGradeYear.Text, entryScoreInfo1);
                }
                else
                    AddScore.InsertSemesterEntryScore(_StudentID, comboBoxEx1.Text, comboBoxEx2.Text, cboAttendGradeYear.Text, "行為", entryScoreInfo1);
            }
            else
            {
                if (_EntryScoreID2 != "")
                    RemoveScore.DeleteSemesterEntityScore(_EntryScoreID2);
            }
            #endregion
            #region 學習類
            if (double.TryParse(textBoxX1.Text, out test)
                || double.TryParse(textBoxX3.Text, out test)
                || double.TryParse(textBoxX5.Text, out test)
                || double.TryParse(textBoxX4.Text, out test)
                || double.TryParse(textBoxX6.Text, out test)
                || double.TryParse(textBoxX7.Text, out test)
                || double.TryParse(textBoxX8.Text, out test)
                || double.TryParse(textBoxX9.Text, out test)
                || double.TryParse(textBoxX10.Text, out test)
                || double.TryParse(textBoxX11.Text, out test)
                || double.TryParse(textBoxX12.Text, out test)
                || double.TryParse(textBoxX13.Text, out test)
                )
            {
                if (double.TryParse(textBoxX1.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "學業");
                    entryElement.SetAttribute("成績", textBoxX1.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX3.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "體育");
                    entryElement.SetAttribute("成績", textBoxX3.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX5.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "健康與護理");
                    entryElement.SetAttribute("成績", textBoxX5.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX4.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "國防通識");
                    entryElement.SetAttribute("成績", textBoxX4.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX6.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "實習科目");
                    entryElement.SetAttribute("成績", textBoxX6.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX7.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "專業科目");
                    entryElement.SetAttribute("成績", textBoxX7.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX13.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "學業(原始)");
                    entryElement.SetAttribute("成績", textBoxX13.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX12.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "體育(原始)");
                    entryElement.SetAttribute("成績", textBoxX12.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX11.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "健康與護理(原始)");
                    entryElement.SetAttribute("成績", textBoxX11.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX10.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "國防通識(原始)");
                    entryElement.SetAttribute("成績", textBoxX10.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX9.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "實習科目(原始)");
                    entryElement.SetAttribute("成績", textBoxX9.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (double.TryParse(textBoxX8.Text, out test))
                {
                    entryElement = doc.CreateElement("Entry");
                    entryElement.SetAttribute("分項", "專業科目(原始)");
                    entryElement.SetAttribute("成績", textBoxX8.Text);
                    entryScoreInfo2.AppendChild(entryElement);
                    _afterXml.AddElement("EntryCollection", entryElement);
                }
                if (_EntryScoreID1 != "")
                {
                    EditScore.UpdateSemesterEntryScore(_EntryScoreID1, cboAttendGradeYear.Text, entryScoreInfo2);
                }
                else
                    AddScore.InsertSemesterEntryScore(_StudentID, comboBoxEx1.Text, comboBoxEx2.Text, cboAttendGradeYear.Text, "學習", entryScoreInfo2);
            }
            else
            {
                if (_EntryScoreID1 != "")
                    RemoveScore.DeleteSemesterEntityScore(_EntryScoreID1);
            }
            #endregion
            #endregion

            #region 處理Log

            StringBuilder desc = new StringBuilder("");
            desc.AppendLine("學生姓名：" + Student.Instance.Items[_StudentID].Name + " ");
            desc.AppendLine(comboBoxEx1.Text + " 學年度 第" + comboBoxEx2.Text + "學期 ");

            #region 處理修課年級log
            if (_beforeXml.GetText("SubjectGradeYear/GradeYear") != _afterXml.GetText("SubjectGradeYear/GradeYear"))
                desc.AppendLine("修課年級由「" + _beforeXml.GetText("SubjectGradeYear/GradeYear") + "」變更為「" + _afterXml.GetText("SubjectGradeYear/GradeYear") + "」");
            #endregion

            #region 處理鎖定成績
            if (_beforeXml.GetText("LockScore/Lock") != _afterXml.GetText("LockScore/Lock"))
                desc.AppendLine("鎖定德行分項成績由「" + _beforeXml.GetText("LockScore/Lock") + "」變更為「" + _afterXml.GetText("LockScore/Lock") + "」");
            #endregion

            #region 處理科目成績log

            //設定分隔符號
            string splitSign = "#$@%#";

            foreach (XmlElement var in _beforeXml.GetElements("SubjectCollection/Subject"))
            {
                string prefix = var.GetAttribute("科目") + splitSign + var.GetAttribute("科目級別") + splitSign;

                foreach (XmlAttribute attr in var.Attributes)
                {
                    if (!_beforeData.ContainsKey(prefix + attr.Name))
                        _beforeData.Add(prefix + attr.Name, attr.InnerText);
                }
            }

            foreach (XmlElement var in _afterXml.GetElements("SubjectCollection/Subject"))
            {
                string prefix = var.GetAttribute("科目") + splitSign + var.GetAttribute("科目級別") + splitSign;

                foreach (XmlAttribute attr in var.Attributes)
                {
                    if (!_afterData.ContainsKey(prefix + attr.Name))
                        _afterData.Add(prefix + attr.Name, attr.InnerText);
                }
            }

            desc.AppendLine("科目成績：");

            string delItem = "";
            foreach (string var in _beforeData.Keys)
            {
                string[] splitWord = var.Split(new string[] { splitSign }, StringSplitOptions.None);

                if (!_afterData.ContainsKey(var))
                {
                    if (splitWord[0] + splitSign + splitWord[1] != delItem)
                        desc.AppendLine("刪除科目「" + splitWord[0] + ((splitWord[1] == "") ? "" : " " + GetNumber(int.Parse(splitWord[1]))) + "」");
                    delItem = splitWord[0] + splitSign + splitWord[1];
                }
                else
                {
                    if (_beforeData[var] != _afterData[var])
                    {
                        desc.AppendLine("科目「" + splitWord[0] + ((splitWord[1] == "") ? "" : " " + GetNumber(int.Parse(splitWord[1]))) + "」的欄位「" + splitWord[2] + "」由「" + _beforeData[var] + "」變更為「" + _afterData[var] + "」");
                    }
                    _afterData.Remove(var);
                }
            }

            string newItem = "";
            foreach (string var in _afterData.Keys)
            {
                string[] splitWord = var.Split(new string[] { splitSign }, StringSplitOptions.None);

                if (splitWord[0] + splitSign + splitWord[1] != newItem)
                    desc.AppendLine("新增科目「" + splitWord[0] + ((splitWord[1] == "") ? "" : " " + GetNumber(int.Parse(splitWord[1]))) + "」");
                newItem = splitWord[0] + splitSign + splitWord[1];
                if (_afterData[var] != "")
                    desc.AppendLine("科目「" + splitWord[0] + ((splitWord[1] == "") ? "" : " " + GetNumber(int.Parse(splitWord[1]))) + "」的欄位「" + splitWord[2] + "」為「" + _afterData[var] + "」");
            }

            _beforeData.Clear();
            _afterData.Clear();

            #endregion

            #region 處理分項成績log

            foreach (XmlElement var in _beforeXml.GetElements("EntryCollection/Entry"))
            {
                if (!_beforeData.ContainsKey(var.GetAttribute("分項")))
                    _beforeData.Add(var.GetAttribute("分項"), var.GetAttribute("成績"));
                if (!_afterData.ContainsKey(var.GetAttribute("分項")))
                    _afterData.Add(var.GetAttribute("分項"), "");
            }

            foreach (XmlElement var in _afterXml.GetElements("EntryCollection/Entry"))
            {
                if (!_afterData.ContainsKey(var.GetAttribute("分項")))
                    _afterData.Add(var.GetAttribute("分項"), var.GetAttribute("成績"));
                else
                    _afterData[var.GetAttribute("分項")] = var.GetAttribute("成績");
                if (!_beforeData.ContainsKey(var.GetAttribute("分項")))
                    _beforeData.Add(var.GetAttribute("分項"), "");
            }

            desc.AppendLine("分項成績：");

            foreach (string var in _afterData.Keys)
            {
                if (_beforeData[var] != _afterData[var])
                    desc.AppendLine("「" + var + "成績」由「" + _beforeData[var] + "」變更為「" + _afterData[var] + "」");
            }

            #endregion

            CurrentUser.Instance.AppLog.Write(EntityType.Student, _entityAction, _StudentID, desc.ToString(), Text, _afterXml.GetRawXml());

            #endregion

            EventHub.Instance.InvokScoreChanged(_StudentID);
            this.Close();
        }

        private XmlElement CreateSubjectScoreElement()
        {
            XmlDocument doc = new XmlDocument();
            XmlElement subjectScoreInfo = doc.CreateElement("SemesterSubjectScoreInfo");

            _afterXml.AddElement("SubjectGradeYear", "GradeYear", cboAttendGradeYear.Text);

            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow)
                    break;
                XmlElement subjectElement = doc.CreateElement("Subject");
                subjectElement.SetAttribute("領域", "" + row.Cells[1].Value);
                subjectElement.SetAttribute("開課分項類別", "" + row.Cells[2].Value);
                subjectElement.SetAttribute("科目", "" + row.Cells[3].Value);
                subjectElement.SetAttribute("科目級別", "" + row.Cells[4].Value);
                subjectElement.SetAttribute("開課學分數", "" + row.Cells[5].Value);
                subjectElement.SetAttribute("修課校部訂", "" + row.Cells[6].Value == "部定" ? "部訂" : row.Cells[6].Value.ToString());
                subjectElement.SetAttribute("修課必選修", "" + row.Cells[7].Value);
                subjectElement.SetAttribute("是否取得學分", (row.Cells[8].Value != null && (bool)row.Cells[8].Value) ? "是" : "否");
                subjectElement.SetAttribute("原始成績", "" + row.Cells[9].Value);
                subjectElement.SetAttribute("補考成績", "" + row.Cells[10].Value);
                subjectElement.SetAttribute("重修成績", "" + row.Cells[11].Value);
                subjectElement.SetAttribute("擇優採計成績", "" + row.Cells[12].Value);
                subjectElement.SetAttribute("學年調整成績", "" + row.Cells[13].Value);
                subjectElement.SetAttribute("不計學分", (row.Cells[14].Value != null && (bool)row.Cells[14].Value) ? "是" : "否");
                subjectElement.SetAttribute("不需評分", (row.Cells[15].Value != null && (bool)row.Cells[15].Value) ? "是" : "否");
                subjectElement.SetAttribute("註記", "" + row.Cells[16].Value);
                subjectElement.SetAttribute("修課及格標準", "" + row.Cells[colPassingStandard.Index].Value);
                subjectElement.SetAttribute("修課補考標準", "" + row.Cells[colMakeupStandard.Index].Value);
                subjectElement.SetAttribute("修課直接指定總成績", "" + row.Cells[colDesignateFinalScore.Index].Value);
                subjectElement.SetAttribute("修課備註", "" + row.Cells[colRemark.Index].Value);
                subjectElement.SetAttribute("修課科目代碼", "" + row.Cells[colCourseCode.Index].Value);

                subjectElement.SetAttribute("是否補修成績", (row.Cells[colIsMakeupScore.Index].Value != null && (bool)row.Cells[colIsMakeupScore.Index].Value) ? "是" : "否");
                subjectElement.SetAttribute("重修學年度", "" + row.Cells[colRetakeSchoolYear.Index].Value);
                subjectElement.SetAttribute("重修學期", "" + row.Cells[colRetakeSemester.Index].Value);
                subjectElement.SetAttribute("補修學年度", "" + row.Cells[colSScoreSchoolYear.Index].Value);
                subjectElement.SetAttribute("補修學期", "" + row.Cells[colSScoreSemester.Index].Value);
                subjectElement.SetAttribute("免修", (row.Cells[colScoreN1.Index].Value != null && (bool)row.Cells[colScoreN1.Index].Value) ? "是" : "否");
                subjectElement.SetAttribute("抵免", (row.Cells[colScoreN2.Index].Value != null && (bool)row.Cells[colScoreN2.Index].Value) ? "是" : "否");
                subjectElement.SetAttribute("指定學年科目名稱", "" + row.Cells[ColSpecifySubjectName.Index].Value);

                // 回填報部科目名稱
                subjectElement.SetAttribute("報部科目名稱", "" + row.Cells[colDSubjectName.Index].Value);

                subjectElement.SetAttribute("是否重讀", (row.Cells[colReread.Index].Value != null && (bool)row.Cells[colReread.Index].Value) ? "是" : "否");

                subjectScoreInfo.AppendChild(subjectElement);

                _afterXml.AddElement("SubjectCollection", subjectElement);
            }
            return subjectScoreInfo;
        }

        #region 資料驗證相關
        private bool ValidateAll()
        {
            errorSchoolYear.Clear();
            errorSemester.Clear();
            errorGradeYear.Clear();
            bool validatePass = true;
            #region 檢查輸入欄位值
            int s = 0;
            if (!int.TryParse(comboBoxEx1.Text, out s))
            {
                validatePass &= false;
                errorSchoolYear.Icon = Properties.Resources.error;
                errorSchoolYear.SetError(comboBoxEx1, "必須輸入數字");
            }
            if (!int.TryParse(comboBoxEx2.Text, out s))
            {
                validatePass &= false;
                errorSemester.Icon = Properties.Resources.error;
                errorSemester.SetError(comboBoxEx2, "必須輸入數字");
            }
            if (!int.TryParse(cboAttendGradeYear.Text, out s))
            {
                validatePass &= false;
                errorGradeYear.Icon = Properties.Resources.error;
                errorGradeYear.SetError(cboAttendGradeYear, "必須輸入數字");
            }
            #endregion
            #region 檢查空值
            if (comboBoxEx1.Text == "")
            {
                validatePass &= false;
                errorSchoolYear.Icon = Properties.Resources.error;
                errorSchoolYear.SetError(comboBoxEx1, "必須填寫");
            }
            if (comboBoxEx2.Text == "")
            {
                validatePass &= false;
                errorSemester.Icon = Properties.Resources.error;
                errorSemester.SetError(comboBoxEx2, "必須填寫");
            }
            if (!int.TryParse(cboAttendGradeYear.Text, out s))
            {
                validatePass &= false;
                errorGradeYear.Icon = Properties.Resources.error;
                errorGradeYear.SetError(cboAttendGradeYear, "修課年級必須填寫");
            }
            #endregion
            #region 檢查DataGridView資料正確
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                if (row.IsNewRow)
                    break;
                validatePass &= ValidateRow(row.Index);
            }
            #endregion


            // 檢查科目名稱+級別是否重複
            List<string> chkSubjectLevel = new List<string>();
            foreach (DataGridViewRow row in dataGridViewX1.Rows)
            {
                string key = row.Cells[3].Value + "_" + row.Cells[4].Value;
                if (chkSubjectLevel.Contains(key))
                {
                    row.Cells[3].ErrorText = "科目名稱+級別重複";
                    row.Cells[4].ErrorText = "科目名稱+級別重複";
                    validatePass &= false;
                }
                else
                {
                    chkSubjectLevel.Add(key);
                }
            }
            return validatePass;
        }

        private bool ValidateRow(int p)
        {
            bool validatePass = true;
            DataGridViewRow row = dataGridViewX1.Rows[p];
            if (row.IsNewRow) return true;
            CompareSubjectInfo(row);

            //foreach (int i in new int[] { 0, 1, 2, 4, 5, 6 })
            //foreach (int i in new int[] { 1, 2, 4, 5, 6 }) // 領域名稱檢查規則，因為舊資料中都會缺少領域欄位，故先不進行檢查，允許空值 --2022.10.05 俊緯
            foreach (int i in new int[] { 2, 3, 5, 6, 7 })   // 2023/5/19，因選科目功能往後移  CT
            {
                row.Cells[i].ErrorText = "";
                if ("" + row.Cells[i].Value == "")
                {
                    validatePass &= false;
                    row.Cells[i].ErrorText = "此為必填欄位";
                }
                dataGridViewX1.UpdateCellErrorText(i, row.Index);
            }

            //foreach (int i in new int[] { 3, 21, 22, 23, 24 })
            foreach (int i in new int[] { 4, 22, 23, 24, 25 })
            {
                row.Cells[i].ErrorText = "";
                int x = 0;
                if ("" + row.Cells[i].Value != "" && !int.TryParse("" + row.Cells[i].Value, out x))
                {
                    validatePass &= false;
                    row.Cells[i].ErrorText = "必須輸入整數";
                }
                dataGridViewX1.UpdateCellErrorText(i, row.Index);
            }

            // foreach (int i in new int[] { 4 })
            foreach (int i in new int[] { 5 })
            {
                if (row.Cells[i].Value != null)
                {

                    row.Cells[i].ErrorText = "";
                    decimal x = 0;
                    if ("" + row.Cells[i].Value != "" && !decimal.TryParse("" + row.Cells[i].Value, out x))
                    {
                        validatePass &= false;
                        row.Cells[i].ErrorText = "必須輸入數字";
                    }
                    dataGridViewX1.UpdateCellErrorText(i, row.Index);
                }
            }
            // foreach (int i in new int[] { 8, 9, 10, 11, 12, 16, 17, 18 }) // 2023/5/18,因前面多一個功能往後移
            foreach (int i in new int[] { 9, 10, 11, 12, 13, 17, 18, 19 })
            {
                row.Cells[i].ErrorText = "";
                double x = 0;
                if ("" + row.Cells[i].Value != "" && !double.TryParse("" + row.Cells[i].Value, out x))
                {
                    validatePass &= false;
                    row.Cells[i].ErrorText = "必須輸入數字";
                }
                dataGridViewX1.UpdateCellErrorText(i, row.Index);
            }

            // 2023/5/17 CT,因為補修要預計挖洞設定，定義驗證規則與這有差，將這註解不使用
            //#region 檢查補修相關資料
            //row.Cells[20].ErrorText = "";
            //row.Cells[21].ErrorText = "";
            //row.Cells[22].ErrorText = "";

            //if ("" + row.Cells[20].Value != "" && "" + row.Cells[20].Value == "True")
            //{

            //    if ("" + row.Cells[21].Value == "")
            //    {
            //        validatePass &= false;
            //        row.Cells[21].ErrorText = "必須輸入整數";
            //    }

            //    if ("" + row.Cells[22].Value == "")
            //    {
            //        validatePass &= false;
            //        row.Cells[22].ErrorText = "必須輸入整數";
            //    }
            //}

            //if ("" + row.Cells[21].Value != "" || "" + row.Cells[22].Value != "")
            //{
            //    if ("" + row.Cells[20].Value != "" && "" + row.Cells[20].Value == "True")
            //    {

            //    }
            //    else
            //    {
            //        validatePass &= false;
            //        row.Cells[20].ErrorText = "是否補修成績必須打勾";
            //    }

            //}
            //#endregion

            #region 檢查重修相關欄位
            if ("" + row.Cells[Column8.Index].Value != "")
            {

                if ("" + row.Cells[colRetakeSchoolYear.Index].Value == "")
                {
                    validatePass &= false;
                    row.Cells[colRetakeSchoolYear.Index].ErrorText = "必須輸入整數。";
                }

                if ("" + row.Cells[colRetakeSemester.Index].Value == "")
                {
                    validatePass &= false;
                    row.Cells[colRetakeSemester.Index].ErrorText = "必須輸入整數。";
                }
            }
            #endregion

            foreach (DataGridViewCell cell in row.Cells)
            {
                validatePass &= (cell.ErrorText == "");
            }

            return validatePass;
        }

        private bool ValidateSchoolYearSemester()
        {
            errorSchoolYear.Clear();
            errorSemester.Clear();
            bool validatePass = true;
            #region 檢查輸入欄位值
            int s = 0;
            if (!int.TryParse(comboBoxEx1.Text, out s))
            {
                validatePass &= false;
                errorSchoolYear.Icon = Properties.Resources.error;
                errorSchoolYear.SetError(comboBoxEx1, "必須輸入數字");
            }
            if (!int.TryParse(comboBoxEx2.Text, out s))
            {
                validatePass &= false;
                errorSemester.Icon = Properties.Resources.error;
                errorSemester.SetError(comboBoxEx2, "必須輸入數字");
            }
            #endregion
            #region 檢查空值
            if (comboBoxEx1.Text == "")
            {
                validatePass &= false;
                errorSchoolYear.Icon = Properties.Resources.error;
                errorSchoolYear.SetError(comboBoxEx1, "必須填寫");
            }
            if (comboBoxEx2.Text == "")
            {
                validatePass &= false;
                errorSemester.Icon = Properties.Resources.error;
                errorSemester.SetError(comboBoxEx2, "必須填寫");
            }
            #endregion
            return validatePass;
        }

        private void dataGridViewX1_CellValidated(object sender, DataGridViewCellEventArgs e)
        {
            ValidateRow(e.RowIndex);

            try
            {
                DataGridViewRow row = dataGridViewX1.Rows[e.RowIndex];
                row.Cells[SubjectColumn].ToolTipText = GetSubjectScorePlace(row);
            }
            catch { }
        }
        #endregion

        private string GetNumber(int p)
        {
            string levelNumber;
            switch (p)
            {
                #region 對應levelNumber
                case 1:
                    levelNumber = "I";
                    break;
                case 2:
                    levelNumber = "II";
                    break;
                case 3:
                    levelNumber = "III";
                    break;
                case 4:
                    levelNumber = "IV";
                    break;
                case 5:
                    levelNumber = "V";
                    break;
                case 6:
                    levelNumber = "VI";
                    break;
                case 7:
                    levelNumber = "VII";
                    break;
                case 8:
                    levelNumber = "VIII";
                    break;
                case 9:
                    levelNumber = "IX";
                    break;
                case 10:
                    levelNumber = "X";
                    break;
                default:
                    levelNumber = "" + (p);
                    break;
                    #endregion
            }
            return levelNumber;
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            dataGridViewX1.EndEdit();
            if (!ValidateAll())
            {
                MsgBox.Show("資料有誤，請修正。");
                return;
            }

            if (ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(_StudentID) == null)
            {
                MsgBox.Show("學生尚未設定計算標準");
                return;
            }
            XmlElement newEntryScoreElement = ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(_StudentID).CalculateSemesterEntryScore(CreateSubjectScoreElement());
            foreach (XmlNode score in newEntryScoreElement.SelectNodes("Entry"))
            {
                XmlElement element = (XmlElement)score;
                _beforeXml.AddElement("EntryCollection", element);

                #region 依分項填入格子
                switch (element.GetAttribute("分項"))
                {
                    case "學業":
                        textBoxX1.Text = element.GetAttribute("成績");
                        break;
                    case "體育":
                        textBoxX3.Text = element.GetAttribute("成績");
                        break;
                    case "國防通識":
                        textBoxX4.Text = element.GetAttribute("成績");
                        break;
                    case "健康與護理":
                        textBoxX5.Text = element.GetAttribute("成績");
                        break;
                    case "實習科目":
                        textBoxX6.Text = element.GetAttribute("成績");
                        break;
                    case "專業科目":
                        textBoxX7.Text = element.GetAttribute("成績");
                        break;
                    case "學業(原始)":
                        textBoxX13.Text = element.GetAttribute("成績");
                        break;
                    case "體育(原始)":
                        textBoxX12.Text = element.GetAttribute("成績");
                        break;
                    case "國防通識(原始)":
                        textBoxX10.Text = element.GetAttribute("成績");
                        break;
                    case "健康與護理(原始)":
                        textBoxX11.Text = element.GetAttribute("成績");
                        break;
                    case "實習科目(原始)":
                        textBoxX9.Text = element.GetAttribute("成績");
                        break;
                    case "專業科目(原始)":
                        textBoxX8.Text = element.GetAttribute("成績");
                        break;
                    case "德行":
                        textBoxX2.Text = element.GetAttribute("成績");
                        break;
                    default:
                        //throw new Exception("拎唄謀洗鰓機雷分項： " + element.GetAttribute("分項"));
                        break;
                }
                #endregion
            }
            //newEntryScoreElement.OwnerDocument.AppendChild(newEntryScoreElement);
            //newEntryScoreElement.OwnerDocument.Save("D:/1234.xml");
        }

        private void entryScoreChanged(object sender, EventArgs e)
        {
            TextBoxX control = (TextBoxX)sender;
            ResetErrorProvider(control);
            if (control == textBoxX2)
            {
                if (!_scoreUpdating && _entryScoreBase.ContainsKey(control) && control.Text != _entryScoreBase[control])
                {
                    SetErrorProvider(control, "由\"" + _entryScoreBase[control] + "\" 修改為 \"" + control.Text + "\"");
                    buttonItem4.Checked = true;
                }
                else
                {
                    buttonItem3.Checked = true;
                }
            }
            else if (!_scoreUpdating && _entryScoreBase.ContainsKey(control) && control.Text != _entryScoreBase[control])
            {
                SetErrorProvider(control, "由\"" + _entryScoreBase[control] + "\" 修改為 \"" + control.Text + "\"");
            }
        }
        private void SetErrorProvider(Control control, string p)
        {
            if (!_errorProviderDictionary.ContainsKey(control))
            {
                ErrorProvider ep = new ErrorProvider();
                ep.BlinkStyle = ErrorBlinkStyle.NeverBlink;
                ep.SetIconAlignment(control, ErrorIconAlignment.MiddleRight);
                ep.Icon = Properties.Resources.Info3D;
                ep.SetError(control, p);
                _errorProviderDictionary.Add(control, ep);
            }
        }

        private void ResetErrorProvider(Control control)
        {
            if (_errorProviderDictionary.ContainsKey(control))
            {
                _errorProviderDictionary[control].Clear();
                _errorProviderDictionary.Remove(control);
            }
        }

        private void dataGridViewX1_RowHeaderMouseClick(object sender, DataGridViewCellMouseEventArgs e)
        {
            dataGridViewX1.EndEdit();
        }

        private void textBoxX2_ButtonCustomClick(object sender, EventArgs e)
        {
            buttonX4.Popup(groupPanel2.PointToScreen(buttonX4.Location));
        }

        private void buttonItem4_CheckedChanged(object sender, EventArgs e)
        {
            //textBoxX2.BackColor = buttonItem4.Checked ? Color.PaleGoldenrod : textBoxX3.BackColor;
            textBoxX2.ButtonCustom2.Visible = buttonItem4.Checked;
            textBoxX2.ButtonCustom.Visible = !textBoxX2.ButtonCustom2.Visible;
        }

        private string GetSubjectScorePlace(DataGridViewRow row)
        {
            if (_subject_rating != null)
                return _subject_rating.GetTooltip(row);
            else
                return string.Empty;
        }

        class SubjectScoreToolTipProvider
        {
            private Dictionary<string, ScorePlace> _class_rating;
            private Dictionary<string, ScorePlace> _dept_rating;
            private Dictionary<string, ScorePlace> _year_rating;

            public SubjectScoreToolTipProvider(XmlElement classRating, XmlElement deptRating, XmlElement yearRating)
            {
                _class_rating = new Dictionary<string, ScorePlace>();
                _dept_rating = new Dictionary<string, ScorePlace>();
                _year_rating = new Dictionary<string, ScorePlace>();

                CreatePlaceObjects(classRating, _class_rating);
                CreatePlaceObjects(deptRating, _dept_rating);
                CreatePlaceObjects(yearRating, _year_rating);
            }

            private void CreatePlaceObjects(XmlElement ratingData, Dictionary<string, ScorePlace> target)
            {
                DSXmlHelper temp;
                foreach (XmlElement eachPlace in ratingData.SelectNodes("Rating/Item"))
                {
                    temp = new DSXmlHelper(eachPlace);
                    string key = temp.GetText("@科目") + temp.GetText("@科目級別");

                    if (target.ContainsKey(key)) continue;

                    target.Add(key, new ScorePlace(eachPlace, temp.GetText("../@範圍人數")));
                }
            }

            public string GetTooltip(DataGridViewRow row)
            {
                //第1欄是科目名稱，第2欄是級別，如果改了的話....。
                string key = row.Cells[SubjectColumn].Value + string.Empty + row.Cells[SubjectLevel].Value;
                ScorePlace temp;
                StringBuilder tooltip = new StringBuilder();

                if (_class_rating.TryGetValue(key, out temp))
                    tooltip.AppendFormat("班排名：{0}\n", temp.Place);

                if (_dept_rating.TryGetValue(key, out temp))
                    tooltip.AppendFormat("科排名：{0}\n", temp.Place);

                if (_year_rating.TryGetValue(key, out temp))
                    tooltip.AppendFormat("年排名：{0}\n", temp.Place);

                return tooltip.ToString();
            }

            class ScorePlace
            {
                public ScorePlace(XmlElement placeData, string ratingBase)
                {
                    DSXmlHelper hlpData = new DSXmlHelper(placeData);

                    _score = hlpData.GetText("@成績");
                    _rating_base = ratingBase;
                    _actual_base = hlpData.GetText("@成績人數");
                    _place = hlpData.GetText("@排名");
                }

                private string _score;
                public string Score
                {
                    get { return _score; }
                }

                private string _actual_base;
                public string ActualBase
                {
                    get { return _actual_base; }
                }

                private string _rating_base;
                public string RatingBase
                {
                    get { return _rating_base; }
                }

                private string _place;
                public string Place
                {
                    get { return _place; }
                }
            }
        }

        private void dataGridViewX1_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            // 選取科目按鈕
            try
            {
                if (e.ColumnIndex == 0 && e.RowIndex > -1)
                {
                    // 取得目前科目名稱與級別
                    List<string> subjLevelList = new List<string>();
                    foreach (DataGridViewRow row in dataGridViewX1.Rows)
                    {
                        if (row.Index == e.RowIndex) continue;

                        string key = row.Cells[3].Value + "_" + row.Cells[4].Value;
                        subjLevelList.Add(key);
                    }

                    frmPickSubject frm = new frmPickSubject();

                    if (dataGridViewX1.Rows[e.RowIndex].IsNewRow)
                        frm.SetIsNewRow(true);
                    else
                        frm.SetIsNewRow(false);

                    frm.SetHasSubjectNameAndLevel(subjLevelList);
                    frm.SetStudentID(_StudentID);

                    // 讀取成績年級
                    frm.SetGradeYear(cboAttendGradeYear.Text);

                    if (frm.ShowDialog() == DialogResult.OK)
                    {
                        // 取得選取的科目
                        List<SubjectInfo> ssList = frm.GetPickSubjects();

                        // 只有1筆覆蓋或新增在NewRow，多筆直接新增
                        if (ssList.Count > 0)
                        {
                            if (ssList.Count == 1)
                            {
                                SubjectInfo ss = ssList[0];

                                if (ss != null)
                                {
                                    // 領域
                                    dataGridViewX1.Rows[e.RowIndex].Cells[1].Value = ss.Domain;
                                    // 分項
                                    dataGridViewX1.Rows[e.RowIndex].Cells[2].Value = ss.Entry;
                                    // 科目
                                    dataGridViewX1.Rows[e.RowIndex].Cells[3].Value = ss.SubjectName;
                                    // 級別
                                    dataGridViewX1.Rows[e.RowIndex].Cells[4].Value = ss.SubjectLevel;
                                    // 學分
                                    dataGridViewX1.Rows[e.RowIndex].Cells[5].Value = ss.Credit;
                                    // 校部定
                                    dataGridViewX1.Rows[e.RowIndex].Cells[6].Value = ss.RequiredBy;
                                    // 必選修
                                    dataGridViewX1.Rows[e.RowIndex].Cells[7].Value = ss.Required;

                                    // 不計學分14

                                    if (ss.NotIncludedInCredit == "是")
                                        dataGridViewX1.Rows[e.RowIndex].Cells[14].Value = true;
                                    else
                                        dataGridViewX1.Rows[e.RowIndex].Cells[14].Value = false;

                                    // 不需評分15
                                    if (ss.NotIncludedInCalc == "是")
                                        dataGridViewX1.Rows[e.RowIndex].Cells[15].Value = true;
                                    else
                                        dataGridViewX1.Rows[e.RowIndex].Cells[15].Value = false;

                                    // 指定學年科目名稱
                                    dataGridViewX1.Rows[e.RowIndex].Cells[ColSpecifySubjectName.Index].Value = ss.SchoolYearSubjectName;

                                    // 課程代碼
                                    dataGridViewX1.Rows[e.RowIndex].Cells[colCourseCode.Index].Value = ss.CourseCode;

                                    // 報部科目名稱
                                    dataGridViewX1.Rows[e.RowIndex].Cells[colDSubjectName.Index].Value = ss.DeptSubjectName;
                                }

                                // 是否重讀
                                dataGridViewX1.Rows[e.RowIndex].Cells[colReread.Index].Value = ss.ReRead;

                                // 修課及格標準
                                dataGridViewX1.Rows[e.RowIndex].Cells[colPassingStandard.Index].Value = ss.PassStandard;

                                // 修課補考標準
                                dataGridViewX1.Rows[e.RowIndex].Cells[colMakeupStandard.Index].Value = ss.MakeUpStandard;

                                // 2024/6/16校務討論，是否補修成績需要填是
                                dataGridViewX1.Rows[e.RowIndex].Cells[colIsMakeupScore.Index].Value = true;

                            }
                            else
                            {
                                // 多筆往後加
                                foreach (SubjectInfo ss in ssList)
                                {
                                    if (ss != null)
                                    {
                                        int rowIdx = dataGridViewX1.Rows.Add();
                                        // 領域
                                        dataGridViewX1.Rows[rowIdx].Cells[1].Value = ss.Domain;
                                        // 分項
                                        dataGridViewX1.Rows[rowIdx].Cells[2].Value = ss.Entry;
                                        // 科目
                                        dataGridViewX1.Rows[rowIdx].Cells[3].Value = ss.SubjectName;
                                        // 級別
                                        dataGridViewX1.Rows[rowIdx].Cells[4].Value = ss.SubjectLevel;
                                        // 學分
                                        dataGridViewX1.Rows[rowIdx].Cells[5].Value = ss.Credit;
                                        // 校部定
                                        dataGridViewX1.Rows[rowIdx].Cells[6].Value = ss.RequiredBy;
                                        // 必選修
                                        dataGridViewX1.Rows[rowIdx].Cells[7].Value = ss.Required;

                                        // 不計學分14

                                        if (ss.NotIncludedInCredit == "是")
                                            dataGridViewX1.Rows[rowIdx].Cells[14].Value = true;
                                        else
                                            dataGridViewX1.Rows[rowIdx].Cells[14].Value = false;

                                        // 不需評分15
                                        if (ss.NotIncludedInCalc == "是")
                                            dataGridViewX1.Rows[rowIdx].Cells[15].Value = true;
                                        else
                                            dataGridViewX1.Rows[rowIdx].Cells[15].Value = false;

                                        // 指定學年科目名稱
                                        dataGridViewX1.Rows[rowIdx].Cells[ColSpecifySubjectName.Index].Value = ss.SchoolYearSubjectName;

                                        // 課程代碼
                                        dataGridViewX1.Rows[rowIdx].Cells[colCourseCode.Index].Value = ss.CourseCode;

                                        // 報部科目名稱
                                        dataGridViewX1.Rows[rowIdx].Cells[colDSubjectName.Index].Value = ss.DeptSubjectName;

                                        // 是否重讀
                                        dataGridViewX1.Rows[rowIdx].Cells[colReread.Index].Value = ss.ReRead;

                                        // 修課及格標準
                                        dataGridViewX1.Rows[rowIdx].Cells[colPassingStandard.Index].Value = ss.PassStandard;

                                        // 修課補考標準
                                        dataGridViewX1.Rows[rowIdx].Cells[colMakeupStandard.Index].Value = ss.MakeUpStandard;

                                        // 2024/6/16校務討論，是否補修成績需要填是
                                        dataGridViewX1.Rows[rowIdx].Cells[colIsMakeupScore.Index].Value = true;
                                    }
                                }
                            }
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

        }

        private void groupPanel1_Click(object sender, EventArgs e)
        {

        }

        private void dataGridViewX1_CellFormatting(object sender, DataGridViewCellFormattingEventArgs e)
        {

            if (dataGridViewX1.Columns[e.ColumnIndex] is DataGridViewButtonColumn && e.RowIndex >= 0)
            {
                DataGridViewButtonCell buttonCell = (DataGridViewButtonCell)dataGridViewX1.Rows[e.RowIndex].Cells[e.ColumnIndex];

                buttonCell.Style.BackColor = Color.SkyBlue;
                buttonCell.ToolTipText = "選取科目";
                buttonCell.Value = "...";
            }
        }

        private void btnCheckCourseCode_Click(object sender, EventArgs e)
        {
            btnCheckCourseCode.Enabled = false;

            // 檢查學生有課程規畫表才比對
            if (GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(_StudentID) != null)
            {
                GraduationPlanInfo gplan = GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(_StudentID);

                // 取得資料表內容
                foreach (DataGridViewRow row in dataGridViewX1.Rows)
                {
                    if (row.IsNewRow) continue;
                    // 使用科目名稱+科目級別 取得課規相同科目

                    string SubjName = "" + row.Cells[3].Value + row.Cells[4].Value;


                    // 有資料
                    // 科目3,級別4
                    GraduationPlanSubject subject = gplan.GetSubjectInfo("" + row.Cells[3].Value, "" + row.Cells[4].Value);

                    // 測試後subject name 如果沒有比對到，科目名稱會變成預設   
                    if (subject.SubjectName != "預設")
                    {
                        // 比對課程代碼相同，跳過不處理，課程代碼不同，以課規覆蓋後，抵免打勾。
                        if (row.Cells[colCourseCode.Index].Value != null)
                        {
                            if (row.Cells[colCourseCode.Index].Value + "" == subject.SubjectCode)
                            {
                                continue;
                            }
                            else
                            {
                                // 填入課程代碼
                                row.Cells[colCourseCode.Index].Value = subject.SubjectCode;
                                // 抵免打勾
                                row.Cells[colScoreN2.Index].Value = true;
                            }
                        }
                        else
                        {
                            // 完全沒有值
                            // 填入課程代碼
                            row.Cells[colCourseCode.Index].Value = subject.SubjectCode;
                            // 抵免打勾
                            row.Cells[colScoreN2.Index].Value = true;
                        }

                    }
                    else
                    {
                        // 沒有資料，課程代碼空白，不計學分14、不須評分15 打勾
                        row.Cells[colCourseCode.Index].Value = "";
                        row.Cells[14].Value = true;
                        row.Cells[15].Value = true;
                    }

                }
            }

            btnCheckCourseCode.Enabled = true;
        }

        private void dataGridViewX1_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void groupPanel2_Click(object sender, EventArgs e)
        {
            if (dataGridViewX1.SelectedRows.Count > 0)
            {
                // 有課程規劃
                if (GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(_StudentID) != null)
                {
                    GraduationPlanInfo gplan = GraduationPlan.GraduationPlan.Instance.GetStudentGraduationPlan(_StudentID);

                    foreach (DataGridViewRow row in dataGridViewX1.SelectedRows)
                    {
                        // 科目3,級別4
                        GraduationPlanSubject subject = gplan.GetSubjectInfo("" + row.Cells[3].Value, "" + row.Cells[4].Value);

                        // 填入課程代碼
                        row.Cells[colCourseCode.Index].Value = subject.SubjectCode;

                    }
                }
            }
        }

        class EntryScoreToolTipProvider
        {
            private System.Windows.Forms.ToolTip _obj_tooltip = new System.Windows.Forms.ToolTip();

            public EntryScoreToolTipProvider(string entryGroup, XmlElement semesterEntryScore)
            {
                XmlElement classRating, deptRating, yearRating;
                DSXmlHelper hlpEntryScore = new DSXmlHelper(semesterEntryScore);
                classRating = hlpEntryScore.GetElement("ClassRating");
                deptRating = hlpEntryScore.GetElement("DeptRating");
                yearRating = hlpEntryScore.GetElement("YearRating");

                StringBuilder tooltip = new StringBuilder();
                string path = string.Format("Rating/Item[@分項='{0}']", entryGroup);

                DSXmlHelper temp = new DSXmlHelper(classRating);
                if (temp.PathExist(path))
                    tooltip.AppendFormat("班排名：{0}\n", temp.GetText(path + "/@排名"));

                temp = new DSXmlHelper(deptRating);
                if (temp.PathExist(path))
                    tooltip.AppendFormat("科排名：{0}\n", temp.GetText(path + "/@排名"));

                temp = new DSXmlHelper(yearRating);
                if (temp.PathExist(path))
                    tooltip.AppendFormat("年排名：{0}\n", temp.GetText(path + "/@排名"));

                _score_tooltip = tooltip.ToString();
            }

            private string _score_tooltip;
            public string ScoreToolTip
            {
                get { return _score_tooltip; }
            }

            public void SetTooltip(Control ctl)
            {
                if (_obj_tooltip.CanExtend(ctl))
                    _obj_tooltip.SetToolTip(ctl, ScoreToolTip);
            }
        }
    }
}