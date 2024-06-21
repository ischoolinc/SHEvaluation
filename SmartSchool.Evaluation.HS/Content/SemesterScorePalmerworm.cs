using FISCA.Data;
using FISCA.DSAUtil;
using FISCA.UDT;
using SmartSchool.AccessControl;
using SmartSchool.ApplicationLog;
using SmartSchool.Common;
using SmartSchool.Evaluation.Content.ChangeSchoolYear;
using SmartSchool.Feature.Score;
using SmartSchool.StudentRelated;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Text;
using System.Windows.Forms;
using System.Xml;

namespace SmartSchool.Evaluation.Content
{
    [FeatureCode("Content0100")]
    public partial class SemesterScorePalmerworm : UserControl, SmartSchool.Customization.PlugIn.ExtendedContent.IContentItem, IPreference
    {

        BackgroundWorker _bkwEntryLoader = new BackgroundWorker();

        BackgroundWorker _bkwSubjectLoader = new BackgroundWorker();

        string _RunningEntryID;

        string _RunningSubjectID;

        string _CurrentID;

        DSResponse _EntryResponse;

        DSResponse _SubjectResponse;

        bool Reload = false;
        bool ArchiveReload = true;  //Cynthia
        private bool WaitingPicVisible { set { this.picWaiting.Visible = value; } }

        public static string FeatureCode = "";
        private FeatureAce _permission;

        public SemesterScorePalmerworm()
        {
            InitializeComponent();
            EventHub.Instance.ScoreChanged += new EventHandler<ScoreChangedEventArgs>(Instance_ScoreChanged);
            PreferenceUpdater.Instance.Items.Add(this);
            LoadPreference();
            lvColumnManager1.DenyColumns.AddRange(new ColumnHeader[] { colSchoolYear, colSemester, cboGradeYear });
            _bkwEntryLoader.DoWork += new DoWorkEventHandler(_bkwEntryLoader_DoWork);
            _bkwEntryLoader.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bkwEntryLoader_RunWorkerCompleted);
            _bkwSubjectLoader.DoWork += new DoWorkEventHandler(_bkwSubjectLoader_DoWork);
            _bkwSubjectLoader.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bkwSubjectLoader_RunWorkerCompleted);
            Application.Idle += new EventHandler(Application_Idle);

            //取得此 Class 定議的 FeatureCode。
            FeatureCodeAttribute code = Attribute.GetCustomAttribute(this.GetType(), typeof(FeatureCodeAttribute)) as FeatureCodeAttribute;
            FeatureCode = code.FeatureCode;

            _permission = CurrentUser.Acl[FeatureCode];

            btnAdd.Visible = _permission.Editable;
            btnDelete.Visible = _permission.Editable;
            btnModify.Visible = _permission.Editable;

            btnView.Visible = !_permission.Editable;

            //學期成績(封存)
            btnArchive.Visible = FISCA.Permission.UserAcl.Current[Permissions.學期成績封存].Editable;
            linkLabel1.Visible = FISCA.Permission.UserAcl.Current[Permissions.學期成績封存].Editable;

        }

        void Application_Idle(object sender, EventArgs e)
        {
            if (Reload)
            {
                LoadContent(_CurrentID);
                Reload = false;
            }
        }

        private void Instance_ScoreChanged(object sender, ScoreChangedEventArgs e)
        {
            foreach (string var in e.StudentIds)
            {
                if (var == _CurrentID)
                {
                    Reload = true;
                }
            }
        }

        private void FillScore()
        {
            if (_EntryResponse == null) return;
            if (_SubjectResponse == null) return;

            WaitingPicVisible = false;
            Dictionary<string, string[]> EntryScore = new Dictionary<string, string[]>();
            Dictionary<string, decimal> Credit = new Dictionary<string, decimal>();
            Dictionary<string, string> GradeYear = new Dictionary<string, string>();
            Dictionary<string, string> GradeYearError = new Dictionary<string, string>();
            foreach (XmlElement var in _SubjectResponse.GetContent().GetElements("SemesterSubjectScore"))
            {
                string schoolyear = var.SelectSingleNode("SchoolYear").InnerText;
                string semester = var.SelectSingleNode("Semester").InnerText;
                string gradeyear = var.SelectSingleNode("GradeYear").InnerText;
                decimal creditCount = 0;
                foreach (XmlElement cnode in var.SelectNodes("ScoreInfo/SemesterSubjectScoreInfo/Subject"))
                {
                    decimal credit = decimal.Parse(cnode.GetAttribute("開課學分數"));
                    bool getCredit = cnode.GetAttribute("是否取得學分") == "是";
                    bool notIncludedInCredit = cnode.GetAttribute("不計學分") == "是";

                    if (getCredit && !notIncludedInCredit)
                    {
                        creditCount += credit;
                    }
                }
                //學分數
                if (Credit.ContainsKey(schoolyear + "_" + semester))
                {
                    Credit[schoolyear + "_" + semester] = creditCount;
                }
                else
                    Credit.Add(schoolyear + "_" + semester, creditCount);
                //年級
                GradeYear.Add(schoolyear + "_" + semester, gradeyear);
            }
            foreach (XmlElement var in _EntryResponse.GetContent().GetElements("SemesterEntryScore"))
            {
                #region 填入分項成績資料
                string schoolyear = var.SelectSingleNode("SchoolYear").InnerText;
                string semester = var.SelectSingleNode("Semester").InnerText;
                string gradeyear = var.SelectSingleNode("GradeYear").InnerText;
                //有這學期的科目成績且成績年級居然不同
                if (GradeYear.ContainsKey(schoolyear + "_" + semester) && GradeYear[schoolyear + "_" + semester] != gradeyear)
                {
                    if (!GradeYearError.ContainsKey(schoolyear + "_" + semester))
                        GradeYearError.Add(schoolyear + "_" + semester, "成績年級錯誤：\n科目成績之年級為" + GradeYear[schoolyear + "_" + semester] + "年級\n分項成績之年級為" + gradeyear + "年級\n請編輯此學期成績，\n並重新指定成績年級為正確之年級後儲存。");
                }
                string[] subItems;
                if (EntryScore.ContainsKey(schoolyear + "_" + semester))
                {
                    subItems = EntryScore[schoolyear + "_" + semester];
                }
                else
                {
                    subItems = new string[11];
                    subItems[0] = schoolyear;
                    subItems[1] = semester;
                    subItems[2] = gradeyear;
                    EntryScore.Add(schoolyear + "_" + semester, subItems);
                }
                foreach (XmlNode score in var.SelectNodes("ScoreInfo/SemesterEntryScore/Entry"))
                {
                    XmlElement element = (XmlElement)score;
                    #region 依分項填入格子
                    switch (element.GetAttribute("分項").Trim())
                    {
                        case "學業":
                            subItems[3] = element.GetAttribute("成績");
                            break;
                        case "實習科目":
                            subItems[4] = element.GetAttribute("成績");
                            break;
                        case "專業科目":
                            subItems[5] = element.GetAttribute("成績");
                            break;
                        case "體育":
                            subItems[6] = element.GetAttribute("成績");
                            break;
                        case "國防通識":
                            subItems[7] = element.GetAttribute("成績");
                            break;
                        case "健康與護理":
                            subItems[8] = element.GetAttribute("成績");
                            break;
                        case "德行":
                            subItems[9] = element.GetAttribute("成績");
                            break;
                        default:
                            //throw new Exception("拎唄謀洗鰓機雷分項： " + element.GetAttribute("分項"));
                            break;
                    }
                    #endregion
                }
                #endregion
            }
            #region 填入總學分數
            foreach (string var in Credit.Keys)
            {
                string[] subItems;
                if (EntryScore.ContainsKey(var))
                {
                    subItems = EntryScore[var];
                }
                else
                {
                    subItems = new string[11];
                    subItems[0] = var.Split("_".ToCharArray())[0];
                    subItems[1] = var.Split("_".ToCharArray())[1];
                    subItems[2] = GradeYear[var];
                    EntryScore.Add(var, subItems);
                }
                subItems[10] = Credit[var].ToString();
            }
            #endregion
            foreach (string key in EntryScore.Keys)
            {
                ListViewItem item = new ListViewItem(EntryScore[key]);
                //如果這個學期成績有成績年級錯誤
                if (GradeYearError.ContainsKey(key))
                {
                    item.ToolTipText = GradeYearError[key];
                    item.ImageIndex = 0;
                }
                listView1.Items.Add(item);
            }
            lvColumnManager1.Reflash();
        }

        void _bkwSubjectLoader_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_RunningEntryID != _CurrentID)
            {
                _RunningEntryID = _CurrentID;
                _bkwSubjectLoader.RunWorkerAsync(_CurrentID);
                return;
            }
            _SubjectResponse = (DSResponse)e.Result;
            FillScore();
        }

        void _bkwSubjectLoader_DoWork(object sender, DoWorkEventArgs e)
        {
            e.Result = QueryScore.GetSemesterSubjectScore(e.Argument.ToString());
        }

        void _bkwEntryLoader_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_RunningEntryID != _CurrentID)
            {
                _RunningEntryID = _CurrentID;
                _bkwEntryLoader.RunWorkerAsync(_CurrentID);
                return;
            }
            _EntryResponse = (DSResponse)e.Result;
            FillScore();
        }

        void _bkwEntryLoader_DoWork(object sender, DoWorkEventArgs e)
        {
            e.Result = QueryScore.GetSemesterEntryScore(e.Argument.ToString());
        }

        public void LoadContent(string id)
        {
            btnModify.Enabled = btnDelete.Enabled = btnArchive.Enabled = btnChangeSchoolYear.Enabled = false;
            listView1.Items.Clear();
            _CurrentID = id;
            if (!_bkwEntryLoader.IsBusy)
            {
                _RunningEntryID = id;
                _EntryResponse = null;
                _bkwEntryLoader.RunWorkerAsync(id);
                WaitingPicVisible = true;
            }
            if (!_bkwSubjectLoader.IsBusy)
            {
                _RunningSubjectID = id;
                _SubjectResponse = null;
                _bkwSubjectLoader.RunWorkerAsync(id);
                WaitingPicVisible = true;
            }
        }
        public void Save()
        {
            //base.Save();
        }
        public void Undo()
        {
            LoadContent(_CurrentID);
        }

        public void UpdatePreference()
        {
            if (this.IsDisposed)
                return;
            #region 紀錄欄位顯示
            #region 取得PreferenceElement
            XmlElement PreferenceElement = CurrentUser.Instance.Preference["SemesterScorePalmerworm"];
            if (PreferenceElement == null)
            {
                PreferenceElement = new XmlDocument().CreateElement("SemesterScorePalmerworm");
            }
            #endregion
            //紀錄顯示欄位
            PreferenceElement.SetAttribute("col健康與護理Visible", listView1.Columns.Contains(col健康與護理).ToString());
            PreferenceElement.SetAttribute("col國防通識Visible", listView1.Columns.Contains(col國防通識).ToString());
            PreferenceElement.SetAttribute("col實得學分Visible", listView1.Columns.Contains(col實得學分).ToString());
            PreferenceElement.SetAttribute("col實習科目Visible", listView1.Columns.Contains(col實習科目).ToString());
            PreferenceElement.SetAttribute("col德行成績Visible", listView1.Columns.Contains(col德行成績).ToString());
            PreferenceElement.SetAttribute("col學業成績Visible", listView1.Columns.Contains(col學業成績).ToString());
            PreferenceElement.SetAttribute("col體育成績Visible", listView1.Columns.Contains(col體育成績).ToString());
            PreferenceElement.SetAttribute("col專業科目Visible", listView1.Columns.Contains(col專業科目).ToString());
            CurrentUser.Instance.Preference["SemesterScorePalmerworm"] = PreferenceElement;
            #endregion
        }

        private void LoadPreference()
        {
            XmlElement PreferenceData = CurrentUser.Instance.Preference["SemesterScorePalmerworm"];
            if (PreferenceData != null)
            {
                string[] attributeList = new string[] { "col健康與護理Visible", "col國防通識Visible", "col實得學分Visible", "col實習科目Visible", "col德行成績Visible", "col學業成績Visible", "col體育成績Visible", "col專業科目Visible" };
                ColumnHeader[] columnList = new ColumnHeader[] { col健康與護理, col國防通識, col實得學分, col實習科目, col德行成績, col學業成績, col體育成績, col專業科目 };
                for (int i = 0; i < attributeList.Length; i++)
                {
                    if (PreferenceData.HasAttribute(attributeList[i]))
                    {
                        bool show = bool.Parse(PreferenceData.GetAttribute(attributeList[i]));
                        if (!show)
                            lvColumnManager1.HideColumn(columnList[i]);
                    }
                }
            }
            else
                //預設顯示欄位
                lvColumnManager1.HideColumn(new ColumnHeader[] { col體育成績, col國防通識, col健康與護理, col實習科目, col專業科目 });
        }

        private void listView1_SelectedIndexChanged(object sender, EventArgs e)
        {
            btnView.Enabled = (listView1.SelectedIndices.Count == 1);
            btnModify.Enabled = (listView1.SelectedIndices.Count == 1 && _permission.Viewable);
            btnDelete.Enabled = (listView1.SelectedIndices.Count == 1 && _permission.Editable);
            btnChangeSchoolYear.Enabled = (listView1.SelectedIndices.Count == 1 && _permission.Editable);

            btnArchive.Enabled = (listView1.SelectedIndices.Count == 1 && FISCA.Permission.UserAcl.Current[Permissions.學期成績封存].Editable);
        }

        private void buttonX3_Click(object sender, EventArgs e)
        {
            if (MsgBox.Show("您確定要刪除此學期成績？", "", MessageBoxButtons.YesNo) == DialogResult.No) return;

            foreach (XmlElement var in QueryScore.GetSemesterEntryScoreBySemester(int.Parse(listView1.SelectedItems[0].SubItems[0].Text), int.Parse(listView1.SelectedItems[0].SubItems[1].Text), _CurrentID).GetContent().GetElements("SemesterEntryScore"))
            {
                RemoveScore.DeleteSemesterEntityScore(var.SelectSingleNode("@ID").InnerText);
            }
            foreach (XmlElement var in QueryScore.GetSemesterSubjectScoreBySemester(int.Parse(listView1.SelectedItems[0].SubItems[0].Text), int.Parse(listView1.SelectedItems[0].SubItems[1].Text), _CurrentID).GetContent().GetElements("SemesterSubjectScore"))
            {
                RemoveScore.DeleteSemesterSubjectScore(var.SelectSingleNode("@ID").InnerText);
            }

            // 刪除學期成績 log
            StringBuilder deleteDesc = new StringBuilder("");
            deleteDesc.AppendLine("學生姓名：" + Student.Instance.Items[_CurrentID].Name + " ");
            deleteDesc.AppendLine("刪除 " + listView1.SelectedItems[0].SubItems[0].Text + " 學年度 第 " + listView1.SelectedItems[0].SubItems[1].Text + " 學期 學期成績");
            CurrentUser.Instance.AppLog.Write(EntityType.Student, EntityAction.Delete, _CurrentID, deleteDesc.ToString(), Title, "");

            EventHub.Instance.InvokScoreChanged(_CurrentID);
        }

        private void buttonX1_Click(object sender, EventArgs e)
        {
            ScoreEditor.SemesterScoreEditor editor = new ScoreEditor.SemesterScoreEditor(_CurrentID);
            editor.FormClosed += new FormClosedEventHandler(editor_FormClosed);
            editor.ShowDialog();
        }

        void editor_FormClosed(object sender, FormClosedEventArgs e)
        {
            Form form = (Form)sender;
            form.FormClosed -= new FormClosedEventHandler(editor_FormClosed);
            //LoadById(_CurrentID);
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            ScoreEditor.SemesterScoreEditor editor = new ScoreEditor.SemesterScoreEditor(listView1.SelectedItems[0].SubItems[0].Text, listView1.SelectedItems[0].SubItems[1].Text, _CurrentID);
            editor.FormClosed += new FormClosedEventHandler(editor_FormClosed);
            editor.ShowDialog();
        }

        private void btnView_Click(object sender, EventArgs e)
        {
            ScoreEditor.SemesterScoreEditor editor = new ScoreEditor.SemesterScoreEditor(listView1.SelectedItems[0].SubItems[0].Text, listView1.SelectedItems[0].SubItems[1].Text, _CurrentID);
            editor.FormClosed += new FormClosedEventHandler(editor_FormClosed);
            editor.ShowDialog();
        }

        private void listView1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (listView1.SelectedIndices.Count == 1)
            {
                //MotherForm.SetWaitCursor();
                ScoreEditor.SemesterScoreEditor editor = new ScoreEditor.SemesterScoreEditor(listView1.SelectedItems[0].SubItems[0].Text, listView1.SelectedItems[0].SubItems[1].Text, _CurrentID);
                editor.FormClosed += new FormClosedEventHandler(editor_FormClosed);
                //MotherForm.ResetWaitCursor();
                editor.ShowDialog();
            }
        }

        #region IContentItem 成員

        public bool CancelButtonVisible
        {
            get { return false; }
        }

        public event EventHandler CancelButtonVisibleChanged;

        public Control DisplayControl
        {
            get { return this; }
        }

        public bool SaveButtonVisible
        {
            get { return false; }
        }

        public event EventHandler SaveButtonVisibleChanged;

        public string Title
        {
            get { return "學期成績"; }
        }

        #endregion

        #region ICloneable 成員

        public object Clone()
        {
            return new SemesterScorePalmerworm();
        }


        #endregion

        /// <summary>
        /// 建立UDT
        /// </summary>
        public static void CreateUDTTable()
        {
            FISCA.UDT.SchemaManager Manager = new SchemaManager(new DSConnection(FISCA.Authentication.DSAServices.DefaultDataSource));
            Manager.SyncSchema(new SemesterSubjectScoreArchive());
            Manager.SyncSchema(new SemesterEntryScoreArchive());
        }


        private void btnArchive_Click(object sender, EventArgs e)
        {
            CreateUDTTable();

            //寫入log
            StringBuilder sb_log = new StringBuilder();
            //封存 (複製sems_subj_score 和 sems_entry_score)

            if (MsgBox.Show("確定將" + listView1.SelectedItems[0].SubItems[0].Text + "-" + listView1.SelectedItems[0].SubItems[1].Text + "學期成績，複製到「學期成績(封存)」？", "", MessageBoxButtons.YesNo) == DialogResult.No) return;

            UpdateHelper updateHelper = new UpdateHelper();
            QueryHelper qh = new QueryHelper();
            int uid = 0;

            string searchSql = "SELECT id FROM sems_subj_score WHERE ref_student_id ={0} AND school_year={1} AND semester={2}";
            searchSql = string.Format(searchSql, _CurrentID, listView1.SelectedItems[0].SubItems[0].Text, listView1.SelectedItems[0].SubItems[1].Text);


            DataTable searchIDdt = qh.Select(searchSql);
            if (searchIDdt.Rows.Count > 0)
            {
                try
                {
                    foreach (XmlElement var in QueryScore.GetSemesterEntryScoreBySemester(int.Parse(listView1.SelectedItems[0].SubItems[0].Text), int.Parse(listView1.SelectedItems[0].SubItems[1].Text), _CurrentID).GetContent().GetElements("SemesterEntryScore"))
                    {
                        //複製學期分項成績(不複製德行成績)
                        string sql = @"INSERT INTO $semester_entry_score_archive (ref_student_id,school_year,semester, grade_year, score_info) 
                                                    SELECT ref_student_id,school_year,semester, grade_year, score_info 
                                                    FROM sems_entry_score 
                                                    WHERE entry_group=1  AND id={0}
                                                    RETURNING uid";

                        sql = string.Format(sql, var.SelectSingleNode("@ID").InnerText);
                        //updateHelper.Execute(sql);
                        //取得回傳的uid讓$semester_subject_score_archive 寫入ref_sems_entry_uid
                        DataTable dt = qh.Select(sql);
                        foreach (DataRow dr in dt.Rows)
                        {
                            uid = int.Parse(dr["uid"].ToString());
                        }
                    }
                }
                catch (Exception ex)
                {
                    MsgBox.Show("取得學生學期分項成績發生錯誤。" + ex.Message);
                }

                try
                {
                    foreach (XmlElement var in QueryScore.GetSemesterSubjectScoreBySemester(int.Parse(listView1.SelectedItems[0].SubItems[0].Text), int.Parse(listView1.SelectedItems[0].SubItems[1].Text), _CurrentID).GetContent().GetElements("SemesterSubjectScore"))
                    {
                        //複製學期科目成績
                        string sql = @"INSERT INTO $semester_subject_score_archive (ref_student_id,school_year,semester, grade_year, score_info,ref_sems_entry_uid) 
                                                    SELECT ref_student_id,school_year,semester, grade_year, score_info ,{0}
                                                    FROM sems_subj_score 
                                                    WHERE  id={1} ";

                        sql = string.Format(sql, uid, var.SelectSingleNode("@ID").InnerText);

                        updateHelper.Execute(sql);
                    }
                }
                catch (Exception ex)
                {
                    MsgBox.Show("取得學生學期科目成績發生錯誤。" + ex.Message);
                    return;
                }
                sb_log.AppendLine("複製學生「" + Student.Instance.Items[_CurrentID].Name + "(學號" + Student.Instance.Items[_CurrentID].StudentNumber + ")」" + listView1.SelectedItems[0].SubItems[0].Text + "學年度第" + listView1.SelectedItems[0].SubItems[1].Text + "學期的學期成績至「學期成績(封存)」。");
                FISCA.LogAgent.ApplicationLog.Log("學期成績", "封存", "學生", Student.Instance.Items[_CurrentID].ID, sb_log.ToString());
                MsgBox.Show("封存成功。");

                //todo 封存後，更新封存資料項目。
                EventHub.OnArchiveChanged(); //Cynthia
            }
            else
                MsgBox.Show(listView1.SelectedItems[0].SubItems[0].Text + "學年度第" + listView1.SelectedItems[0].SubItems[1].Text + "學期沒有科目成績。");
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            MsgBox.Show("\r\n" +
                "1. 需要有「學期成績(封存)」的編輯權限才能使用。\r\n\r\n" +
                "2. 此功能運用在「校內轉科」，用來保留學生當時的學期科目成績及課程代碼。\r\n\r\n" +
                "3. 封存的成績不可修改。\r\n\r\n" +
                "4. 德行成績已過時，故不在封存範圍。\r\n\r\n" +
                "5. 流程：\r\n" +
                "　　(1) 請先封存學生當前所有的學期成績至「學期成績(封存)」。\r\n" +
                "　　(2) 進行轉科異動作業。\r\n" +
                "　　(3) 將不可抵免的科目自學期成績中刪除。", "學期成績(封存)說明", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        private void btnChangeSchoolYear_Click(object sender, EventArgs e)
        {
            btnChangeSchoolYear.Enabled = false;

            frmChangeSemesterSchoolYearMain fm = new frmChangeSemesterSchoolYearMain();
            // 傳入學年度 0、學期 1、年級 2
            try
            {
                //   listView1.SelectedItems[0].SubItems[0].Text, listView1.SelectedItems[0].SubItems[1].Text
                fm.SetSourceInfo(listView1.SelectedItems[0].SubItems[0].Text, listView1.SelectedItems[0].SubItems[1].Text, listView1.SelectedItems[0].SubItems[2].Text);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            fm.ShowDialog();

            btnChangeSchoolYear.Enabled = true;
        }
    }
}
