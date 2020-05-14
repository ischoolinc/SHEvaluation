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
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using Aspose.Words;

namespace ClassSemesterScoreReportFixed_SH
{
    public partial class ConfigForm : FISCA.Presentation.Controls.BaseForm
    {
        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();
        private Dictionary<string, List<string>> _ExamSubjects = new Dictionary<string, List<string>>();
        private Dictionary<string, List<string>> _ExamSubjectFull = new Dictionary<string, List<string>>();
        private List<TagConfigRecord> _TagConfigRecords = new List<TagConfigRecord>();
        private List<Configure> _Configures = new List<Configure>();
        private string _DefalutSchoolYear = "";
        private string _DefaultSemester = "";

        // 紀錄樣板設定
        List<DAO.UDT_ScoreConfig> _UDTConfigList;

        public ConfigForm()
        {
            InitializeComponent();
            List<ExamRecord> exams = new List<ExamRecord>();
            BackgroundWorker bkw = new BackgroundWorker();
            bkw.DoWork += delegate
            {
                bkw.ReportProgress(1);
                //預設學年度學期
                _DefalutSchoolYear = "" + K12.Data.School.DefaultSchoolYear;
                _DefaultSemester = "" + K12.Data.School.DefaultSemester;
                bkw.ReportProgress(10);

                // 檢查預設樣板是否存在
                _UDTConfigList = DAO.UDTTransfer.GetDefaultConfigNameListByTableName(Global._UDTTableName);
                 // 沒有設定檔，建立預設設定檔
                if (_UDTConfigList.Count < 2)
                {
                    bkw.ReportProgress(20);
                    foreach (string name in Global.DefaultConfigNameList())
                    {
                        Configure cn = new Configure();
                        cn.Name = name;
                        cn.SchoolYear = K12.Data.School.DefaultSchoolYear;
                        cn.Semester = K12.Data.School.DefaultSemester;
                        DAO.UDT_ScoreConfig conf = new DAO.UDT_ScoreConfig();
                        conf.Name = name;
                        conf.UDTTableName = Global._UDTTableName;
                        conf.ProjectName = Global._ProjectName;
                        conf.Type = Global._DefaultConfTypeName;
                        _UDTConfigList.Add(conf);

                        // 設預設樣板
                        switch (name)
                        {
                            case "班級學期成績單13科":
                                cn.Template = new Document(new MemoryStream(Properties.Resources.高中班級學期成績單樣版));

                                break;

                            case "班級學期成績單24科":
                                cn.Template = new Document(new MemoryStream(Properties.Resources.班級_學期成績單24科50學生));
                                break;

                        }

                        if (cn.Template == null)
                            cn.Template = new Document(new MemoryStream(Properties.Resources.高中班級學期成績單樣版));

                        try
                        {
                            List<string> fields = new List<string>(cn.Template.MailMerge.GetFieldNames());
                            cn.SubjectLimit = 0;
                            while (fields.Contains("科目名稱" + (cn.SubjectLimit + 1)))
                            {
                                cn.SubjectLimit++;
                            }
                            cn.StudentLimit = 0;
                            while (fields.Contains("姓名" + (cn.StudentLimit + 1)))
                            {
                                cn.StudentLimit++;
                            }
                        }
                        catch (Exception ex)
                        { }

                        cn.Encode();
                        cn.Save();
                    }
                    if (_UDTConfigList.Count > 0)
                        DAO.UDTTransfer.InsertConfigData(_UDTConfigList);
                }
                bkw.ReportProgress(40);
                 _Configures = _AccessHelper.Select<Configure>();
                //學生類別清單
                _TagConfigRecords = K12.Data.TagConfig.SelectByCategory(TagCategory.Student);
                #region 整理所有試別對應科目
                
                bkw.ReportProgress(60);
                // 取得班級學生(一般生)
                List<string> ClassStudentIDList = DAO.QueryData.GetClassStudentIDByClassID(K12.Presentation.NLDPanels.Class.SelectedSource);
                AccessHelper accessHelper = new AccessHelper();
                List<SmartSchool.Customization.Data.StudentRecord> StudentRecList = accessHelper.StudentHelper.GetStudents(ClassStudentIDList);
                bkw.ReportProgress(70);
                // 放入學生學期成績
                accessHelper.StudentHelper.FillSemesterSubjectScore(true, StudentRecList);                
                bkw.ReportProgress(80);
                // 取得科目名稱
                foreach (SmartSchool.Customization.Data.StudentRecord stud in StudentRecList)
                {
                    foreach (SemesterSubjectScoreInfo ss in stud.SemesterSubjectScoreList)
                    {
                        string key = ss.SchoolYear + "^^" + ss.Semester;
                        if (!_ExamSubjectFull.ContainsKey(key))
                            _ExamSubjectFull.Add(key, new List<string>());
                        if (!_ExamSubjectFull[key].Contains(ss.Subject))
                            _ExamSubjectFull[key].Add(ss.Subject);
                    }                
                }
               
                bkw.ReportProgress(90);
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
               
                bkw.ReportProgress(100);

            };
            bkw.WorkerReportsProgress = true;
            bkw.ProgressChanged += delegate(object sender, ProgressChangedEventArgs e)
            {
                circularProgress1.Value = e.ProgressPercentage;
            };
            bkw.RunWorkerCompleted += delegate
            {
                if (Configure == null)
                    Configure = new Configure();

                cboConfigure.Items.Clear();
                foreach (var item in _Configures)
                {
                    cboConfigure.Items.Add(item);
                }
                cboConfigure.Items.Add(new Configure() { Name = "新增" });
                int i;
                if (int.TryParse(_DefalutSchoolYear, out i))
                {
                    for (int j = 0; j < 5; j++)
                    {
                        cboSchoolYear.Items.Add("" + (i - j));
                    }
                }
                cboSemester.Items.Add("1");
                cboSemester.Items.Add("2");
         
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
                    string userSelectConfigName = "";
                    // 檢查畫面上是否有使用者選的
                    foreach (DAO.UDT_ScoreConfig conf in _UDTConfigList)
                        if (conf.Type == Global._UserConfTypeName)
                        {
                            userSelectConfigName = conf.Name;
                            break;
                        } 
                 
                        cboConfigure.Text = userSelectConfigName;
                }
                else
                {
                    cboConfigure.SelectedIndex = -1;
                }
            };
            bkw.RunWorkerAsync();
        }

        public Configure Configure { get; private set; }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            // 檢查設定學年度學期與目前系統預設是否相同
            string s1 = cboSchoolYear.Text + cboSemester.Text;
            string s2 = School.DefaultSchoolYear + School.DefaultSemester;

            if (s1 != s2)
                if (FISCA.Presentation.Controls.MsgBox.Show("畫面上學年度學期與系統學年度學期不相同，請問是否繼續?", "學年度學期不同", MessageBoxButtons.YesNo, MessageBoxIcon.Warning, MessageBoxDefaultButton.Button1) == System.Windows.Forms.DialogResult.No)
                    return;

            SaveTemplate(null, null);

            this.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

        private void ExamChanged(object sender, EventArgs e)
        {
            string key = cboSchoolYear.Text + "^^" + cboSemester.Text;// +"^^" +
                //(cboExam.SelectedItem == null ? "" : ((ExamRecord)cboExam.SelectedItem).ID);
            lvwSubjectName.SuspendLayout();
          
            lvwSubjectName.Items.Clear();
          
            if (_ExamSubjectFull.ContainsKey(key))
            {
                foreach (var subject in _ExamSubjectFull[key])
                {
                    var i1 = lvwSubjectName.Items.Add(subject);
          
                    if (Configure != null && Configure.PrintSubjectList.Contains(subject))
                        i1.Checked = true;
             
                    if (_ExamSubjects.ContainsKey(key) && !_ExamSubjects[key].Contains(subject))
                    {
                        i1.ForeColor = Color.DarkGray;
             
                    }
                }
            }
            lvwSubjectName.ResumeLayout(true);
          
        }

        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
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
                 
                
                    foreach (ListViewItem item in lvwSubjectName.Items)
                    {
                        item.Checked = Configure.PrintSubjectList.Contains(item.Text);
                    }
          
                }
                else
                {
                    Configure = null;
                    cboSchoolYear.SelectedIndex = -1;
                    cboSemester.SelectedIndex = -1;
                
          
                    foreach (ListViewItem item in lvwSubjectName.Items)
                    {
                        item.Checked = false;
                    }
          
                }
            }
        }

       



        private void linkLabel4_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Configure == null) return;

            // 檢查是否是預設設定檔名稱，如果是無法刪除
            if (Global.DefaultConfigNameList().Contains(Configure.Name))
            {
                FISCA.Presentation.Controls.MsgBox.Show("系統預設設定檔案無法刪除");
                return;
            }

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
                conf.Encode();
                conf.Save();
                _Configures.Add(conf);
                cboConfigure.Items.Insert(cboConfigure.Items.Count - 1, conf);
                cboConfigure.SelectedIndex = cboConfigure.Items.Count - 2;
            }
        }

        private void SaveTemplate(object sender, EventArgs e)
        {
            if (Configure == null) return;
            Configure.SchoolYear = cboSchoolYear.Text;
            Configure.Semester = cboSemester.Text;
            Configure.SelSetConfigName = cboConfigure.Text;
            if (Configure.RefenceExamRecord != null && Configure.RefenceExamRecord.Name == "")
                Configure.RefenceExamRecord = null;
            foreach (ListViewItem item in lvwSubjectName.Items)
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
          
            Configure.TagRank1TagList.Clear();           
            Configure.RankFilterTagList.Clear();          

            Configure.Encode();
            Configure.Save();

            #region 樣板設定檔記錄用

            // 記錄使用這選的專案            
            List<DAO.UDT_ScoreConfig> uList = new List<DAO.UDT_ScoreConfig>();
            foreach (DAO.UDT_ScoreConfig conf in _UDTConfigList)
                if (conf.Type == Global._UserConfTypeName)
                {
                    conf.Name = cboConfigure.Text;
                    uList.Add(conf);
                    break;
                }

            if (uList.Count > 0)
            {
                DAO.UDTTransfer.UpdateConfigData(uList);
            }
            else
            {
                // 新增
                List<DAO.UDT_ScoreConfig> iList = new List<DAO.UDT_ScoreConfig>();
                DAO.UDT_ScoreConfig conf = new DAO.UDT_ScoreConfig();
                conf.Name = cboConfigure.Text;
                conf.ProjectName = Global._ProjectName;
                conf.Type = Global._UserConfTypeName;
                conf.UDTTableName = Global._UDTTableName;
                iList.Add(conf);
                DAO.UDTTransfer.InsertConfigData(iList);
            } 
            #endregion
        }

        private void chkAllSubj_CheckedChanged(object sender, EventArgs e)
        {
            foreach (ListViewItem lvi in lvwSubjectName.Items)
                lvi.Checked = chkAllSubj.Checked;
        }
        
        
        private void lnkDownloadTemplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (this.Configure == null) return;
            lnkDownloadTemplate.Enabled = false;
            #region 儲存檔案
            string inputReportName = "個人學期成績單樣板(" + this.Configure.Name + ").doc";
            string reportName =  Utility.ParseFileName(inputReportName);

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
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        //document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.高中班級學期成績單樣版, 0, Properties.Resources.高中班級學期成績單樣版.Length);
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
            lnkDownloadTemplate.Enabled = true;
        }

        private void lnkUploadTenplate_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (Configure == null) return;
            lnkUploadTenplate.Enabled = false;
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Title = "上傳樣板";
            dialog.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
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
            lnkUploadTenplate.Enabled = true;
        }

        private void lnkViewMapping_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkViewMapping.Enabled = false;
            Program.CreateFieldTemplate();
            lnkViewMapping.Enabled = true;
        }

        private void ConfigForm_Load(object sender, EventArgs e)
        {

        }

        private void tableLayoutPanel1_Paint(object sender, PaintEventArgs e)
        {

        }

        private void btnCancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {

        }
    }
}
