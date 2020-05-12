using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using K12.Data;
using System.IO;
using FISCA.Data;
//using SmartSchool.Customization.Data;
using static K12.Data.StudentRecord;
using FISCA.Presentation.Controls;

namespace 班級定期評量成績單_固定排名
{
    public partial class ConfigForm : FISCA.Presentation.Controls.BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();
        private SmartSchool.Customization.Data.AccessHelper _smartAccessHelper = new SmartSchool.Customization.Data.AccessHelper();
        private Dictionary<string, List<string>> _ExamSubjects = new Dictionary<string, List<string>>();
        private Dictionary<string, List<string>> _ExamSubjectFull = new Dictionary<string, List<string>>();
        private List<TagConfigRecord> _TagConfigRecords = new List<TagConfigRecord>();
        private List<Configure> _Configures = new List<Configure>();
        private string _DefalutSchoolYear = "";
        private string _DefaultSemester = "";
        private List<string> _FixedRankSubjects = new List<string>();
       // private List<string> _SelectClassGrade;
        private QueryHelper _Qp = new QueryHelper();
        private List<string> _SelectedClasses; // 取得選擇學生

        public Dictionary<string, List<string>> FixedRankCumputeSubject = new Dictionary<string, List<string>>();
        /// <summary>
        /// 用來記錄此學年度 學期  此班級 結算的類別是否只有一種 只有一種( 可能情境: 1.自然組 、 2.社會組)
        /// </summary>
        private Dictionary<string, List<string>> FixedRankTags;


        public ConfigForm(List<String> selectClasses)
        {
            InitializeComponent();
            List<SmartSchool.Customization.Data.ClassRecord> _SelectedClasses = this._smartAccessHelper.ClassHelper.GetSelectedClass(); ; //取得所選班級
            List<ExamRecord> exams = new List<ExamRecord>();
            this._SelectedClasses = selectClasses;

            BackgroundWorker bkw = new BackgroundWorker();
            bkw.DoWork += delegate
            {
                bkw.ReportProgress(1);
                //預設學年度學期
                _DefalutSchoolYear = "" + K12.Data.School.DefaultSchoolYear;
                _DefaultSemester = "" + K12.Data.School.DefaultSemester;
                bkw.ReportProgress(10);
                //試別清單
                exams = K12.Data.Exam.SelectAll();
                bkw.ReportProgress(20);
                //學生類別清單
                _TagConfigRecords = K12.Data.TagConfig.SelectByCategory(TagCategory.Student);
                #region 整理所有試別對應科目
                var AEIncludeRecords = K12.Data.AEInclude.SelectAll();
                bkw.ReportProgress(30);
                var AssessmentSetupRecords = K12.Data.AssessmentSetup.SelectAll();
                bkw.ReportProgress(40);
                List<string> courseIDs = new List<string>();
                foreach (var scattentRecord in K12.Data.SCAttend.SelectByStudentIDs(K12.Data.Student.SelectByClassIDs(K12.Presentation.NLDPanels.Class.SelectedSource).Where(x => x.Status == StudentStatus.一般 || x.Status == StudentStatus.延修).Select(x => x.ID)))
                {
                    if (!courseIDs.Contains(scattentRecord.RefCourseID))
                        courseIDs.Add(scattentRecord.RefCourseID);
                }
                bkw.ReportProgress(60);
                foreach (var courseRecord in K12.Data.Course.SelectAll())
                {
                    foreach (var aeIncludeRecord in AEIncludeRecords)
                    {
                        if (aeIncludeRecord.RefAssessmentSetupID == courseRecord.RefAssessmentSetupID)
                        {
                            string key = courseRecord.SchoolYear + "^^" + courseRecord.Semester + "^^" + aeIncludeRecord.RefExamID;
                            if (!_ExamSubjectFull.ContainsKey(key))
                            {
                                _ExamSubjectFull.Add(key, new List<string>());
                            }
                            if (!_ExamSubjectFull[key].Contains(courseRecord.Subject))
                                _ExamSubjectFull[key].Add(courseRecord.Subject);
                            if (courseIDs.Contains(courseRecord.ID))
                            {
                                if (!_ExamSubjects.ContainsKey(key))
                                {
                                    _ExamSubjects.Add(key, new List<string>());
                                }
                                if (!_ExamSubjects[key].Contains(courseRecord.Subject))
                                    _ExamSubjects[key].Add(courseRecord.Subject);
                            }
                        }
                    }
                }
                bkw.ReportProgress(70);
                foreach (var list in _ExamSubjectFull.Values)
                {
                    #region 排序
                    list.Sort(new StringComparer("國文"
                                    , "英文"
                                    , "數學"
                                    , "理化"
                                    , "生物"
                                    , "社會"
                                    , "物理"
                                    , "化學"
                                    , "歷史"
                                    , "地理"
                                    , "公民"));
                    #endregion
                }
                #endregion
                bkw.ReportProgress(80);
                _Configures = _AccessHelper.Select<Configure>();
                bkw.ReportProgress(100);

            };
            bkw.WorkerReportsProgress = true;
            bkw.ProgressChanged += delegate (object sender, ProgressChangedEventArgs e)
            {
                circularProgress1.Value = e.ProgressPercentage;
            };
            bkw.RunWorkerCompleted += delegate
            {
                cboConfigure.Items.Clear();
                foreach (var item in _Configures)
                {
                    cboConfigure.Items.Add(item);
                }
                cboConfigure.Items.Add(new Configure() { Name = "新增" });
                int i;

                cboSchoolYear.Items.Clear();
                cboSchoolYear.Items.Add(_DefalutSchoolYear);
                cboSchoolYear.SelectedIndex = cboSchoolYear.Items.IndexOf(_DefaultSemester);

                cboSemester.Items.Clear();
                cboSemester.Items.Add(_DefaultSemester);
                cboSemester.SelectedIndex = cboSemester.Items.IndexOf(_DefaultSemester);

                cboExam.Items.Clear();
                cboRefExam.Items.Clear();
                cboExam.Items.AddRange(exams.ToArray());
                cboRefExam.Items.Add(new ExamRecord("", "", 0));
                cboRefExam.Items.AddRange(exams.ToArray());
                List<string> prefix = new List<string>();
                List<string> tag = new List<string>();
                foreach (var item in _TagConfigRecords)
                {
                    if (item.Prefix != "")
                    {
                        if (!prefix.Contains(item.Prefix))
                            prefix.Add(item.Prefix);
                    }
                    else
                    {
                        tag.Add(item.Name);
                    }
                }

                circularProgress1.Hide();
                if (_Configures.Count > 0)
                {
                    cboConfigure.SelectedIndex = 0;
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
                // 畫面初始化完後 依據條件(學年度 、學期 、試別、所選班級 等) select 回 固定排名之 科目
                GetFixRankSubjectsInclude(this.cboSchoolYear.Text, this.cboSemester.Text, ((ExamRecord)cboExam.SelectedItem).ID, this._SelectedClasses);
                this.pictureBox1.Visible = false;
            };
            bkw.RunWorkerAsync();


        }

        public Configure Configure { get; private set; }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            if (!CheckFixRankSubIsSameAndBack()) 
            {
                return;
            }
            SaveTemplate(null, null);
            Program.AvgRd = iptRd.Value;
            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void ExamChanged(object sender, EventArgs e)
        {
            this._FixedRankSubjects.Clear();

            if (!String.IsNullOrEmpty(this.cboSemester.Text) && !String.IsNullOrEmpty(this.cboSchoolYear.Text) && cboExam.SelectedItem != null)
            {
                #region 取得本次固定排名結算之科目

                GetFixRankSubjectsInclude(this.cboSchoolYear.Text, this.cboSemester.Text, ((ExamRecord)cboExam.SelectedItem).ID, this._SelectedClasses);
            
            }
            #endregion


            string key = cboSchoolYear.Text + "^^" + cboSemester.Text + "^^" +
                (cboExam.SelectedItem == null ? "" : ((ExamRecord)cboExam.SelectedItem).ID);

            listViewEx1.SuspendLayout();
            listViewEx1.Items.Clear();
            if (_ExamSubjectFull.ContainsKey(key))
            {
                foreach (var subject in _ExamSubjectFull[key])
                {
                    var i1 = listViewEx1.Items.Add(subject);

                    if (_ExamSubjects.ContainsKey(key) && !_ExamSubjects[key].Contains(subject))
                    {
                        i1.ForeColor = Color.DarkGray;
                    }
                }
            }
            listViewEx1.ResumeLayout(true);

        }

        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
            this.pictureBox1.Visible = true;
            if (cboConfigure.SelectedIndex == cboConfigure.Items.Count - 1)
            {
                //新增
                btnSaveConfig.Enabled = btnPrint.Enabled = false;
                NewConfigure dialog = new NewConfigure();
                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Configure = new Configure();
                    Configure.Name = dialog.ConfigName;
                    Configure.Template = dialog.Template;
                    Configure.SubjectLimit = dialog.SubjectLimit;
                    Configure.StudentLimit = dialog.StudentLimit;
                    Configure.SchoolYear = _DefalutSchoolYear;
                    Configure.Semester = _DefaultSemester;
                    if (cboExam.Items.Count > 0)
                        Configure.ExamRecord = (ExamRecord)cboExam.Items[0];
                    _Configures.Add(Configure);
                    cboConfigure.Items.Insert(cboConfigure.SelectedIndex, Configure);
                    cboConfigure.SelectedIndex = cboConfigure.SelectedIndex - 1;
                    Configure.Encode();
                    Configure.Save();
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            }
            else
            {
                if (cboConfigure.SelectedIndex >= 0)
                {
                    btnSaveConfig.Enabled = btnPrint.Enabled = true;
                    Configure = _Configures[cboConfigure.SelectedIndex];
                    if (Configure.Template == null)
                        Configure.Decode();
                    if (!cboSchoolYear.Items.Contains(Configure.SchoolYear))
                        cboSchoolYear.Items.Add(Configure.SchoolYear);
                    cboSchoolYear.Text = Configure.SchoolYear;
                    cboSemester.Text = Configure.Semester;
                    if (Configure.ExamRecord != null)
                    {
                        foreach (var item in cboExam.Items)
                        {
                            if (((ExamRecord)item).ID == Configure.ExamRecord.ID)
                            {
                                cboExam.SelectedIndex = cboExam.Items.IndexOf(item);
                                break;
                            }
                        }
                    }
                    cboRefExam.SelectedIndex = -1;
                    if (Configure.RefenceExamRecord != null)
                    {
                        foreach (var item in cboRefExam.Items)
                        {
                            if (((ExamRecord)item).ID == Configure.RefenceExamRecord.ID)
                            {
                                cboRefExam.SelectedIndex = cboRefExam.Items.IndexOf(item);
                                break;
                            }
                        }
                    }

                    foreach (ListViewItem item in listViewEx1.Items)
                    {
                        item.Checked = Configure.PrintSubjectList.Contains(item.Text);
                    }


                    if (Configure.AvgRd.HasValue)
                        iptRd.Value = Configure.AvgRd.Value;
                    else
                        iptRd.Value = 2;
                }
                else
                {
                    Configure = null;
                    cboSchoolYear.SelectedIndex = -1;
                    cboSemester.SelectedIndex = -1;
                    cboExam.SelectedIndex = -1;
                    cboRefExam.SelectedIndex = -1;

                    foreach (ListViewItem item in listViewEx1.Items)
                    {
                        item.Checked = false;
                    }

                }
            }
            this.pictureBox1.Visible = false;
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.Configure == null) return;
            #region 儲存檔案
            string inputReportName = "個人學期成績單樣板(" + this.Configure.Name + ").doc";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                //document.Save(path, Aspose.Words.SaveFormat.Doc);
                System.IO.FileStream stream = new FileStream(path, FileMode.Create, FileAccess.Write);
                this.Configure.Template.Save(stream, Aspose.Words.SaveFormat.Doc);
                //stream.Write(Properties.Resources.個人學期成績單樣板_高中_, 0, Properties.Resources.個人學期成績單樣板_高中_.Length);
                stream.Flush();
                stream.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Excel檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        //document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.高中班級考試成績單樣版, 0, Properties.Resources.高中班級考試成績單樣版.Length);
                        stream.Flush();
                        stream.Close();

                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
        }

        private void linkLabel2_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Configure == null) return;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Excel檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                try
                {
                    this.Configure.Template = new Aspose.Words.Document(dialog.FileName);
                    List<string> fields = new List<string>(this.Configure.Template.MailMerge.GetFieldNames());
                    this.Configure.SubjectLimit = 0;
                    while (fields.Contains("科目名稱" + (this.Configure.SubjectLimit + 1)))
                    {
                        this.Configure.SubjectLimit++;
                    }
                    this.Configure.StudentLimit = 0;
                    while (fields.Contains("姓名" + (this.Configure.StudentLimit + 1)))
                    {
                        this.Configure.StudentLimit++;
                    }
                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗");
                }
            }
        }

        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Configure == null) return;
            if (MessageBox.Show("樣板刪除後將無法回復，確定刪除樣板?", "刪除樣板", MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == System.Windows.Forms.DialogResult.OK)
            {
                _Configures.Remove(Configure);
                if (Configure.UID != "")
                {
                    Configure.Deleted = true;
                    Configure.Save();
                }
                var conf = Configure;
                cboConfigure.SelectedIndex = -1;
                cboConfigure.Items.Remove(conf);
            }
        }

        private void linkLabel3_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Configure == null) return;
            CloneConfigure dialog = new CloneConfigure() { ParentName = Configure.Name };
            if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                Configure conf = new Configure();
                conf.Name = dialog.NewConfigureName;
                conf.ExamRecord = Configure.ExamRecord;
                conf.PrintSubjectList.AddRange(Configure.PrintSubjectList);
                conf.RankFilterTagList.AddRange(Configure.RankFilterTagList);
                conf.RankFilterTagName = Configure.RankFilterTagName;
                conf.RefenceExamRecord = Configure.RefenceExamRecord;
                conf.SchoolYear = Configure.SchoolYear;
                conf.Semester = Configure.Semester;
                conf.SubjectLimit = Configure.SubjectLimit;
                conf.StudentLimit = Configure.StudentLimit;
                conf.TagRank1SubjectList.AddRange(Configure.TagRank1SubjectList);
                conf.TagRank1TagList.AddRange(Configure.TagRank1TagList);
                conf.TagRank1TagName = Configure.TagRank1TagName;
                conf.TagRank2SubjectList.AddRange(Configure.TagRank2SubjectList);
                conf.TagRank2TagList.AddRange(Configure.TagRank2TagList);
                conf.TagRank2TagName = Configure.TagRank2TagName;
                conf.Template = Configure.Template;
                conf.AvgRd = Configure.AvgRd;
                conf.Encode();
                conf.Save();
                _Configures.Add(conf);
                cboConfigure.Items.Insert(cboConfigure.Items.Count - 1, conf);
                cboConfigure.SelectedIndex = cboConfigure.Items.Count - 2;
            }
        }

        /// <summary>
        /// 收集 configure 會儲存之資料
        /// </summary>
        private void CollectConfigure()
        {
            if (Configure == null) return;
            Configure.SchoolYear = cboSchoolYear.Text;
            Configure.Semester = cboSemester.Text;
            Configure.ExamRecord = ((ExamRecord)cboExam.SelectedItem);
            Configure.RefenceExamRecord = ((ExamRecord)cboRefExam.SelectedItem);
            if (Configure.RefenceExamRecord != null && Configure.RefenceExamRecord.Name == "")
                Configure.RefenceExamRecord = null;

        }

        private void SaveTemplate(object sender, EventArgs e)
        {
            this.CollectConfigure();
            foreach (ListViewItem item in listViewEx1.Items)
            {
                if (item.Checked)
                {
                    if (!Configure.PrintSubjectList.Contains(item.Text))
                        Configure.PrintSubjectList.Add(item.Text);
                }
                else
                {
                    if (Configure.PrintSubjectList.Contains(item.Text))
                        Configure.PrintSubjectList.Remove(item.Text);
                }
            }
            //Configure.TagRank1TagName = cboTagRank1.Text;

            // 處理 tag 
            Configure.TagRank1TagList.Clear();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {

                }
                else
                {

                }
            }


            Configure.AvgRd = iptRd.Value;

            Configure.TagRank2TagList.Clear();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {

                }
                else
                {

                }
            }

            Configure.RankFilterTagList.Clear();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {

                }
                else
                {

                }
            }

            Configure.Encode();
            Configure.Save();
        }

        private void linkLabel5_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this.CollectConfigure();
            Utility.CreateFieldTemplate();
        }

        private void cboSchoolYear_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void linkLabFixRankSubjInclude_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            FixedRankInclued fixedRankInclued = new FixedRankInclued(FixedRankCumputeSubject,this._SelectedClasses);
            fixedRankInclued.ShowDialog();
            if (fixedRankInclued.BringSelectedSubj) //如果有要帶入
            {
                foreach (ListViewItem item in listViewEx1.Items)
                {
                    if (fixedRankInclued.SelectSubjucts.Contains(item.Text))
                    {
                        item.Checked = true;
                    }
                    else
                    {
                        item.Checked = false;
                    }
                }


            }
        }

        private void labelX1_Click(object sender, EventArgs e)
        {

        }


        /// <summary>
        /// 取得最新結算之固定排名 第一層 key值 【"年、科、班排名" 、"類別1排名" 、"類別2排名"】
        /// </summary>
        /// <param name="schoolYear">學年度</param>
        /// <param name="semester">學期</param>
        /// <param name="refExamID">試別</param>
        /// <param name="grades">年級s</param>
        private void GetFixRankSubjectsInclude(string schoolYear, string semester, string refExamID, List<string> selectClasseIDs)
        {
            QueryHelper queryHelper = new QueryHelper();
            this.FixedRankCumputeSubject.Clear();
            string sql = @"
SELECT 
	item_name
	, rank_type
	, count (*)  
FROM 
	rank_matrix  
	INNER JOIN  
	( SELECT * FROM rank_detail  WHERE ref_student_id  IN ( SELECT id FROM student WHERE ref_class_id  IN ({3}) AND status = 1) ) AS rank_detail 
	ON rank_matrix.id =rank_detail.ref_matrix_id 
WHERE 
	ref_exam_id = {2}  
    AND item_type ='定期評量/科目成績'   
    AND school_year = {0} 
    AND semester = {1}  
    AND is_alive = true 
GROUP BY 
	item_name 
	, rank_type
";
            sql = string.Format(sql, schoolYear, semester, refExamID, string.Join(",", selectClasseIDs));

            DataTable dt = queryHelper.Select(sql);

            foreach (DataRow dr in dt.Rows)
            {
                string rankType = "" + dr["rank_type"]; 
                 
                string itemName = "" + dr["item_name"];

                if(rankType== "年排名"|| rankType =="科排名" ||rankType =="班排名") 
                {
                    rankType = "年、科、班排名";
                }

                if (!this.FixedRankCumputeSubject.ContainsKey(rankType))
                {
                    this.FixedRankCumputeSubject.Add(rankType, new List<string>());
                }
                if (!this.FixedRankCumputeSubject[rankType].Contains(itemName)) 
                {
                    this.FixedRankCumputeSubject[rankType].Add(itemName);
                }
            }
        }

        
        /// <summary>
        /// 確認是所勾選科目否一致
        /// </summary>
        private Boolean CheckFixRankSubIsSameAndBack() 
        {
            Boolean normalWarmHasShowed = false; // 紀錄 是否跳出ㄊㄧ
            Boolean Tag1WarnHasShowed = false;
            Boolean Tag2WarnHasShowed = false;

            Boolean result = true;

            int selectSubjCount = this.listViewEx1.CheckedItems.Count;

            foreach (ListViewItem subj in this.listViewEx1.CheckedItems)
            {
                // 班科年
                if (!normalWarmHasShowed &&(!this.FixedRankCumputeSubject["年、科、班排名"].Contains(subj.Text) || selectSubjCount!= this.FixedRankCumputeSubject["年、科、班排名"].Count ))
                {
                    normalWarmHasShowed = true;
                    if (DialogResult.No == MsgBox.Show("勾選之列印科目與固定排名計算(班、科、年)之科目不一致。 \n可能導致相關變數有誤，確定列印 ?", MessageBoxButtons.YesNo))
                    {
                        result = false;
                    }
                  
                }
                // 類別1
                if (Tag1WarnHasShowed &&(!this.FixedRankCumputeSubject["類別1排名"].Contains(subj.Text) || selectSubjCount!= this.FixedRankCumputeSubject["類別1排名"].Count ) )
                {
                    Tag1WarnHasShowed = true;

                    if (DialogResult.No == MsgBox.Show("勾選之列印科目與固定排名計算(類別1排名)之科目不一致。 \n可能導致相關變數有誤 ，確定列印 ?", MessageBoxButtons.YesNo))
                    {
                        result = false;
                    }
                }
                // 類別2
                if (Tag2WarnHasShowed && (!this.FixedRankCumputeSubject["類別2排名"].Contains(subj.Text) || selectSubjCount != this.FixedRankCumputeSubject["類別2排名"].Count)) 
                {
                    Tag2WarnHasShowed = true;
                    if (DialogResult.No == MsgBox.Show("勾選之列印科目與固定排名計算(類別2排名)之科目不一致。 \n可能導致相關變數有誤 ，確定列印 ?", MessageBoxButtons.YesNo))
                    {
                        result = false;
                    }
                }
            }

            return result;
        }

  
    }
}
