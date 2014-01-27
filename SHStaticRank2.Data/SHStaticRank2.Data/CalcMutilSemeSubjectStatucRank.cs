using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using K12.Data;
using FISCA.Data;
using Aspose.Words;
using System.IO;
using DevComponents.DotNetBar.Controls;

namespace SHStaticRank2.Data
{
    public partial class CalcMutilSemeSubjectStatucRank : BaseForm
    {
      
        private List<TagConfigRecord> _TagConfigRecords = new List<TagConfigRecord>();

        private string _NewCboConfigName = "";
        

        private FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();
        
        BackgroundWorker _bgWorker1 = new BackgroundWorker();
        BackgroundWorker _bgWorker2 = new BackgroundWorker();
        BackgroundWorker _bgWorker3 = new BackgroundWorker();
        BackgroundWorker _bgWorker4 = new BackgroundWorker();

        List<string> SubjectNameList = new List<string>();
        
        private string _ReportName = "高中多學期成績單測試版";
        
        private List<Configure> _Configures = new List<SHStaticRank2.Data.Configure>();

        // 用來紀錄那個ListView有沒有勾選"部訂必修專業科目"或"部訂必修實習科目"
        // Key: ListView's name, Value: ListViewItem's index
        private Dictionary<string, int> _SpecialListViewItem = new Dictionary<string,int>();

        public CalcMutilSemeSubjectStatucRank()
        {
            _TagConfigRecords = K12.Data.TagConfig.SelectByCategory(TagCategory.Student);           

            _bgWorker1.DoWork += new DoWorkEventHandler(_bgWorker1_DoWork);
            _bgWorker1.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker1_RunWorkerCompleted);
            _bgWorker1.ProgressChanged += new ProgressChangedEventHandler(_bgWorker1_ProgressChanged);
            _bgWorker1.WorkerReportsProgress = true;

            _bgWorker2.DoWork += new DoWorkEventHandler(_bgWorker2_DoWork);
            _bgWorker2.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker2_RunWorkerCompleted);
            _bgWorker2.ProgressChanged += new ProgressChangedEventHandler(_bgWorker2_ProgressChanged);
            _bgWorker2.WorkerReportsProgress = true;

            _bgWorker3.DoWork += new DoWorkEventHandler(_bgWorker3_DoWork);
            _bgWorker3.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker3_RunWorkerCompleted);
            _bgWorker3.ProgressChanged += new ProgressChangedEventHandler(_bgWorker3_ProgressChanged);
            _bgWorker3.WorkerReportsProgress = true;

            _bgWorker4.DoWork += new DoWorkEventHandler(_bgWorker4_DoWork);
            _bgWorker4.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker4_RunWorkerCompleted);
            _bgWorker4.ProgressChanged += new ProgressChangedEventHandler(_bgWorker4_ProgressChanged);
            _bgWorker4.WorkerReportsProgress = true;

            InitializeComponent();
            buttonX1.Enabled = false;
            chkGrade1.Enabled = chkGrade2.Enabled = chkGrade3.Enabled = chkGrade4.Enabled = false;          

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
            // 不排名學生類別
            cboRankRilter.Items.Clear();
            // 類別排名1
            cboTagRank1.Items.Clear();
            // 類別排名2
            cboTagRank2.Items.Clear();

            cboRankRilter.Items.Add("");
            cboTagRank1.Items.Add("");
            cboTagRank2.Items.Add("");
            foreach (var s in prefix)
            {
                cboRankRilter.Items.Add("[" + s + "]");
                cboTagRank1.Items.Add("[" + s + "]");
                cboTagRank2.Items.Add("[" + s + "]");
            }
            foreach (var s in tag)
            {
                cboRankRilter.Items.Add(s);
                cboTagRank1.Items.Add(s);
                cboTagRank2.Items.Add(s);
            }

            // 處理ListView中item選取的事件
            lvwSubjectPri.ItemChecked += ListViewItemChecked;
            lvwSubjectPri.ItemCheck += ListViewItemCheck;
            lvwSubjectOrd1.ItemChecked += ListViewItemChecked;
            lvwSubjectOrd1.ItemCheck += ListViewItemCheck;
            lvwSubjectOrd2.ItemChecked += ListViewItemChecked;
            lvwSubjectOrd2.ItemCheck += ListViewItemCheck;
            // 用來紀錄那個ListView有沒有勾選"部訂必修專業科目"或"部訂必修實習科目"
            _SpecialListViewItem.Add("lvwSubjectPri", int.MinValue);
            _SpecialListViewItem.Add("lvwSubjectOrd1", int.MinValue);
            _SpecialListViewItem.Add("lvwSubjectOrd2", int.MinValue);
        }

        void _bgWorker4_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
           // circularProgress1.Value = e.ProgressPercentage;
        }

        void _bgWorker4_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            RunWorkerCompleted();  
        }

        void _bgWorker4_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker4.ReportProgress(1);
            P1();
            _bgWorker4.ReportProgress(70);
            P2();
            _bgWorker4.ReportProgress(100);
        }

        void _bgWorker3_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //circularProgress1.Value = e.ProgressPercentage;
        }

        void _bgWorker3_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            RunWorkerCompleted();  
        }

        void _bgWorker3_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker3.ReportProgress(1);
            P1();
            _bgWorker3.ReportProgress(70);
            P2();
            _bgWorker3.ReportProgress(100);
        }

        void _bgWorker2_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            //circularProgress1.Value = e.ProgressPercentage;
        }

        void _bgWorker2_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            RunWorkerCompleted();  
        }

        void _bgWorker2_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker2.ReportProgress(1);
            P1();
            _bgWorker2.ReportProgress(70);
            P2();
            _bgWorker2.ReportProgress(100);
        }

        void _bgWorker1_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
           // circularProgress1.Value = e.ProgressPercentage;
        }

        void _bgWorker1_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            RunWorkerCompleted();  
        }

        void _bgWorker1_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker1.ReportProgress(1);
            P1();
            _bgWorker1.ReportProgress(70);
            P2();
            _bgWorker1.ReportProgress(100);
            
        }    

        private void RunWorkerCompleted()
        {
            // 新增"部訂必修專業科目"及"部訂必修實習科目"選項
            ListViewItem lvItem = new ListViewItem();
            lvItem.Text = "部訂必修專業科目";
            lvwSubjectPri.Items.Add(lvItem.Clone() as ListViewItem);
            lvwSubjectOrd1.Items.Add(lvItem.Clone() as ListViewItem);
            lvwSubjectOrd2.Items.Add(lvItem.Clone() as ListViewItem);

            lvItem = new ListViewItem();
            lvItem.Text = "部訂必修實習科目";
            lvwSubjectPri.Items.Add(lvItem.Clone() as ListViewItem);
            lvwSubjectOrd1.Items.Add(lvItem.Clone() as ListViewItem);
            lvwSubjectOrd2.Items.Add(lvItem.Clone() as ListViewItem);
            
            foreach (string str in SubjectNameList)
            {
                ListViewItem lv1 = new ListViewItem();
                ListViewItem lv2 = new ListViewItem();
                ListViewItem lv3 = new ListViewItem();
                lv1.Text = lv2.Text = lv3.Text = str;
                //lv1.Checked = lv2.Checked = lv3.Checked = true;

                lvwSubjectPri.Items.Add(lv1);
                lvwSubjectOrd1.Items.Add(lv2);
                lvwSubjectOrd2.Items.Add(lv3);
            }
            circularProgress1.Visible = false;
            buttonX1.Enabled = true;
            btnSaveConfig.Enabled = true;

            cboConfigure.Items.Clear();
            foreach (var item in _Configures)
            {
                cboConfigure.Items.Add(item);
            }


            cboConfigure.Items.Add(new Configure() { Name = "新增" });

            chkGrade1.Enabled = chkGrade2.Enabled = chkGrade3.Enabled = chkGrade4.Enabled = true;


            //if (!string.IsNullOrWhiteSpace(_NewCboConfigName))
            //{
            int idx = cboConfigure.FindString(_NewCboConfigName);
            if (idx > -1)
            {
                cboConfigure.Text = _NewCboConfigName;                
                //cboConfigure.SelectedIndex = idx;
            }
            //}

            readData();
            cboConfigure.Enabled = true;
        }



        private void P1()
        {
            // 畫面上選的年級
            List<string> grList = new List<string>();
            if (chkGrade1.Checked)
                grList.Add("1");

            if (chkGrade2.Checked)
                grList.Add("2");

            if (chkGrade3.Checked)
                grList.Add("3");

            if (chkGrade4.Checked)
                grList.Add("4");

            // 狀態一般學生編號
            List<string> studIDList = new List<string>();
            QueryHelper qh = new QueryHelper();

            string strSQ = "select student.id from student inner join class on student.ref_class_id=class.id where student.status=1 and class.grade_year in(" + string.Join(",", grList.ToArray()) + ");";
            DataTable dt = qh.Select(strSQ);
            foreach (DataRow dr in dt.Rows)
                studIDList.Add(dr[0].ToString());

            // 學期科目名稱列表          
            SubjectNameList.Clear();
            List<K12.Data.SemesterScoreRecord> scoreList = K12.Data.SemesterScore.SelectByStudentIDs(studIDList);
            foreach (K12.Data.SemesterScoreRecord sRec in scoreList)
            {
                foreach (string str in sRec.Subjects.Keys)
                {
                    if (!SubjectNameList.Contains(str))
                        SubjectNameList.Add(str);
                }
            }
        }

        private void P2()
        {
            // 科目排序
            SubjectNameList.Sort(new StringComparer("國文"
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

            _Configures = _AccessHelper.Select<Configure>();

        }

      

        private void CalcSemeSubjectStatucRank_Load(object sender, EventArgs e)
        {
            chkGrade1.Checked = true;
            //_bgWorker1.RunWorkerAsync();
            _Configures = _AccessHelper.Select<Configure>();
            cboConfigure.Items.Clear();
            foreach (var item in _Configures)
            {
                cboConfigure.Items.Add(item);
            }
            cboConfigure.Items.Add(new Configure() { Name = "新增" });


            if (cboConfigure.Items.Count > 1)
                cboConfigure.SelectedIndex = 0;
            else
                _bgWorker1.RunWorkerAsync();
        }

          

        public Configure Setting { get; private set; }
        private void buttonX1_Click(object sender, EventArgs e)
        {
            SaveTenplate(null, null);
            if (lvwSubjectPri.CheckedItems.Count == 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇列印科目..");
                return;
            }        

            // 檢查年級學期
            int chkGrSemesterSelect = 0;
            foreach (Control co in gpGradeSemester.Controls)
            {
                CheckBoxX cb = co as CheckBoxX;
                if (cb != null && cb.Checked)
                    chkGrSemesterSelect++;                    
            }

            if (chkGrSemesterSelect == 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇成績年級學期..");
                return;
            }

            // 檢查採用成績
            int chkUseScoreSelect = 0;
            foreach (Control co in gpUseScore.Controls)
            {
                CheckBoxX cb = co as CheckBoxX;
                if (cb != null && cb.Checked)
                    chkUseScoreSelect++;
            }
            if (chkUseScoreSelect == 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請選擇採計成績..");
                return;
            }

            Setting = Configure;

            //Setting = new Configure();


            string LogGString = "";
            if (chkGrade1.Checked)
                LogGString = chkGrade1.Text;

            if (chkGrade2.Checked)
                LogGString = chkGrade2.Text;

            if (chkGrade3.Checked)
                LogGString = chkGrade3.Text;

            if (chkGrade4.Checked)
                LogGString = chkGrade4.Text;

            FISCA.LogAgent.ApplicationLog.Log("成績", "計算", "計算"+LogGString+"多學期成績單。");

            Setting.CalcGradeYear1 = chkGrade1.Checked;
            Setting.CalcGradeYear2 = chkGrade2.Checked;
            Setting.CalcGradeYear3 = chkGrade3.Checked;
            Setting.CalcGradeYear4 = chkGrade4.Checked;
            Setting.DoNotSaveIt = chkDoNotSaveIt.Checked;
            Setting.NotRankTag = cboRankRilter.Text;
            Setting.use原始成績 = chk原始成績.Checked;
            Setting.use補考成績 = chk補考成績.Checked;
            Setting.use重修成績 = chk重修成績.Checked;
            Setting.use手動調整成績 = chk手動調整成績.Checked;
            Setting.use學年調整成績 = chk學年調整成績.Checked;
            Setting.Rank1Tag = cboTagRank1.Text;
            Setting.Rank2Tag = cboTagRank2.Text;
            Setting.計算學業成績排名 = chk計算學業成績排名.Checked;

            if (cbg11.Checked)
                Setting.useGradeSemesterList.Add("11");

            if (cbg12.Checked)
                Setting.useGradeSemesterList.Add("12");

            if (cbg21.Checked)
                Setting.useGradeSemesterList.Add("21");

            if (cbg22.Checked)
                Setting.useGradeSemesterList.Add("22");

            if (cbg31.Checked)
                Setting.useGradeSemesterList.Add("31");

            if (cbg32.Checked)
                Setting.useGradeSemesterList.Add("32");

            if (cbg41.Checked)
                Setting.useGradeSemesterList.Add("41");

            if (cbg42.Checked)
                Setting.useGradeSemesterList.Add("42");


            // 列印科目
            foreach (ListViewItem lvi in lvwSubjectPri.CheckedItems)
                Setting.useSubjectPrintList.Add(lvi.Text);

            // 排名科目1
            foreach (ListViewItem lvi in lvwSubjectOrd1.CheckedItems)
                Setting.useSubjecOrder1List.Add(lvi.Text);

            // 排名科目2
            foreach (ListViewItem lvi in lvwSubjectOrd2.CheckedItems)
                Setting.useSubjecOrder2List.Add(lvi.Text);

            DialogResult = System.Windows.Forms.DialogResult.OK;
            this.Close();
        }

      

        private void chkGrade1_CheckedChanged(object sender, EventArgs e)
        {
            lvwSubjectPri.Items.Clear();
            lvwSubjectOrd1.Items.Clear();
            lvwSubjectOrd2.Items.Clear();
            btnSaveConfig.Enabled= buttonX1.Enabled = false;
            chkGrade1.Enabled = chkGrade2.Enabled = chkGrade3.Enabled = chkGrade4.Enabled = false;
            cboConfigure.Enabled = false;
                circularProgress1.Visible = true;

                if (chkGrade1.Checked && _bgWorker1.IsBusy==false)
                    _bgWorker1.RunWorkerAsync();

                if (chkGrade2.Checked && _bgWorker2.IsBusy == false)
                    _bgWorker2.RunWorkerAsync();

                if (chkGrade3.Checked && _bgWorker3.IsBusy == false)
                    _bgWorker3.RunWorkerAsync();

                if (chkGrade4.Checked && _bgWorker4.IsBusy == false)
                    _bgWorker4.RunWorkerAsync();
            
        }

        private void lnkDownload_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            DownloadDefaultTemplate();
        }
        /// <summary>
        /// 下載樣板
        /// </summary>
        private void DownloadDefaultTemplate()
        {
            if (this.Configure == null) return;
            #region 儲存檔案
            string inputReportName = "多學期成績單樣板(" + this.Configure.Name + ").doc";
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
                sd.Filter = "Word檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        //document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);
                        System.IO.FileStream stream = new FileStream(sd.FileName, FileMode.Create, FileAccess.Write);
                        stream.Write(Properties.Resources.高中多學期學生成績證明書, 0, Properties.Resources.高中多學期學生成績證明書.Length);
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

        private void lnkUpload_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            UploadUserDefTemplate();
        }

        /// <summary>
        /// 上傳使用者範本
        /// </summary>
        private void UploadUserDefTemplate()
        {
            if (Configure == null) return;
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
                }
                catch
                {
                    MessageBox.Show("樣板開啟失敗");
                }
            }
        }

      

        private void lblMappingTemp_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
                Document doc = new Document(new MemoryStream(Properties.Resources.合併欄位總表));

            SaveFileDialog saveDialog = new SaveFileDialog();
            saveDialog.Filter = "Word (*.doc)|*.doc";
            saveDialog.FileName = "合併欄位總表";
            if (saveDialog.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    doc.Save(saveDialog.FileName); 
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("儲存失敗。" + ex.Message);
                    return;
                }

                try
                {
                    System.Diagnostics.Process.Start(saveDialog.FileName);
                }
                catch (Exception ex)
                {
                    FISCA.Presentation.Controls.MsgBox.Show("開啟失敗。" + ex.Message);
                    return;
                }
            }
        }

        public Configure Configure { get; private set; }

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
                conf.PrintSubjectList.AddRange(Configure.PrintSubjectList);
                conf.RankFilterTagList.AddRange(Configure.RankFilterTagList);
                conf.RankFilterTagName = Configure.RankFilterTagName;
                conf.SubjectLimit = Configure.SubjectLimit;
                conf.TagRank1SubjectList.AddRange(Configure.TagRank1SubjectList);
                conf.TagRank1TagList.AddRange(Configure.TagRank1TagList);
                conf.TagRank1TagName = Configure.TagRank1TagName;
                conf.TagRank2SubjectList.AddRange(Configure.TagRank2SubjectList);
                conf.TagRank2TagList.AddRange(Configure.TagRank2TagList);
                conf.TagRank2TagName = Configure.TagRank2TagName;
                conf.RankFilterGradeSemeter = Configure.RankFilterGradeSemeter;
                conf.RankFilterUseScoreName = Configure.RankFilterUseScoreName;
                conf.SortGradeYear = Configure.SortGradeYear;
                conf.WithCalSemesterScoreRank = Configure.WithCalSemesterScoreRank;
                conf.RankFilterUseScoreList.AddRange(Configure.RankFilterUseScoreList);
                conf.RankFilterGradeSemeterList.AddRange(Configure.RankFilterGradeSemeterList);
                conf.WithCalSemesterScoreRank = Configure.WithCalSemesterScoreRank;
                conf.Template = Configure.Template;
                conf.Encode();
                conf.Save();
                _Configures.Add(conf);
                cboConfigure.Items.Insert(cboConfigure.Items.Count - 1, conf);
                cboConfigure.SelectedIndex = cboConfigure.Items.Count - 2;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void SaveTenplate(object sender, EventArgs e)
        {
            if (Configure == null) return;
            btnSaveConfig.Enabled = false;
            foreach (ListViewItem item in lvwSubjectPri.Items)
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
            Configure.TagRank1TagName = cboTagRank1.Text;
            Configure.TagRank1TagList.Clear();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {
                    if (cboTagRank1.Text == "[" + item.Prefix + "]")
                        Configure.TagRank1TagList.Add(item.ID);
                }
                else
                {
                    if (cboTagRank1.Text == item.Name)
                        Configure.TagRank1TagList.Add(item.ID);
                }
            }
            foreach (ListViewItem item in lvwSubjectOrd1.Items)
            {
                if (item.Checked)
                {
                    if (!Configure.TagRank1SubjectList.Contains(item.Text))
                        Configure.TagRank1SubjectList.Add(item.Text);
                }
                else
                {
                    if (Configure.TagRank1SubjectList.Contains(item.Text))
                        Configure.TagRank1SubjectList.Remove(item.Text);
                }
            }

            Configure.RankFilterGradeSemeterList.Clear();
            // 勾選年級學期
            foreach (Control cr in gpGradeSemester.Controls)
            {
                CheckBoxX item = cr as CheckBoxX;
                if (item != null)
                    if (item.Checked)
                        if(!Configure.RankFilterGradeSemeterList.Contains(item.Text))
                            Configure.RankFilterGradeSemeterList.Add(item.Text);
            }
            Configure.RankFilterUseScoreList.Clear();
            // 採計成績
            foreach (Control cr in gpUseScore.Controls)
            {
                CheckBoxX item = cr as CheckBoxX;
                if (item != null)
                    if (item.Checked)
                        if(!Configure.RankFilterUseScoreList.Contains(item.Text))
                            Configure.RankFilterUseScoreList.Add(item.Text);
            }

            // 排名對象
            if (chkGrade1.Checked)
                Configure.SortGradeYear = chkGrade1.Text;

            if (chkGrade2.Checked)
                Configure.SortGradeYear = chkGrade2.Text;

            if (chkGrade3.Checked)
                Configure.SortGradeYear = chkGrade3.Text;

            if (chkGrade4.Checked)
                Configure.SortGradeYear = chkGrade4.Text;


            Configure.WithCalSemesterScoreRank = chk計算學業成績排名.Checked;

            Configure.TagRank2TagName = cboTagRank2.Text;
            Configure.TagRank2TagList.Clear();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {
                    if (cboTagRank2.Text == "[" + item.Prefix + "]")
                        Configure.TagRank2TagList.Add(item.ID);
                }
                else
                {
                    if (cboTagRank2.Text == item.Name)
                        Configure.TagRank2TagList.Add(item.ID);
                }
            }
            foreach (ListViewItem item in lvwSubjectOrd2.Items)
            {
                if (item.Checked)
                {
                    if (!Configure.TagRank2SubjectList.Contains(item.Text))
                        Configure.TagRank2SubjectList.Add(item.Text);
                }
                else
                {
                    if (Configure.TagRank2SubjectList.Contains(item.Text))
                        Configure.TagRank2SubjectList.Remove(item.Text);
                }
            }

            Configure.RankFilterTagName = cboRankRilter.Text;
            Configure.RankFilterTagList.Clear();
            foreach (var item in _TagConfigRecords)
            {
                if (item.Prefix != "")
                {
                    if (cboRankRilter.Text == "[" + item.Prefix + "]")
                        Configure.RankFilterTagList.Add(item.ID);
                }
                else
                {
                    if (cboRankRilter.Text == item.Name)
                        Configure.RankFilterTagList.Add(item.ID);
                }
            }

            Configure.Encode();
            Configure.Save();
            btnSaveConfig.Enabled = true;
        
        }


        private void btnSaveConfig_Click(object sender, EventArgs e)
        {
            SaveTenplate(null, null);
        
        }

        private void readData()
        {
            if (_NewCboConfigName != cboConfigure.Text)
                chkGrade1.Checked = chkGrade2.Checked = chkGrade3.Checked = chkGrade4.Checked = false;

            // 初始化, 並宣告temp用來暫存使用者有沒有勾選"部訂必修專業科目"或"部訂必修實習科目"
            Dictionary<string, int> tmpSpecialListViewItem = new Dictionary<string, int>();
            foreach (string key in _SpecialListViewItem.Keys.ToList<string>())
            {
                _SpecialListViewItem[key] = int.MinValue;
                tmpSpecialListViewItem.Add(key, int.MinValue);
            }

            if (cboConfigure.SelectedIndex == cboConfigure.Items.Count - 1)
            {
                //新增
                btnSaveConfig.Enabled = buttonX1.Enabled = false;
                NewConfigure dialog = new NewConfigure();

                List<string> nameList = new List<string>();
                foreach (Configure cf in _Configures)
                    nameList.Add(cf.Name);

                dialog.SetNameList(nameList);

                if (dialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    Configure = new Configure();
                    Configure.Name = dialog.ConfigName;
                    Configure.Template = dialog.Template;
                    Configure.SubjectLimit = dialog.SubjectLimit;
                    Configure.SortGradeYear = "三年級";
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
                    btnSaveConfig.Enabled = buttonX1.Enabled = true;
                    Configure = _Configures[cboConfigure.SelectedIndex];
                    if (Configure.Template == null)
                        Configure.Decode();
                    cboRankRilter.Text = Configure.RankFilterTagName;
                    foreach (ListViewItem item in lvwSubjectPri.Items)
                    {
                        // 這兩個特殊選項, 需要最後勾選, 不然會造成其他正常選項無法勾選, 以及太早上色
                        if (item.Text == "部訂必修專業科目" || item.Text == "部訂必修實習科目")
                        {
                            if (Configure.PrintSubjectList.Contains(item.Text))
                            {
                                // 先用temp存起來, 不然其他正常選項無法勾選
                                tmpSpecialListViewItem["lvwSubjectPri"] = item.Index;
                            }
                        }
                        else
                            item.Checked = Configure.PrintSubjectList.Contains(item.Text);
                    }
                    cboTagRank1.Text = Configure.TagRank1TagName;
                    foreach (ListViewItem item in lvwSubjectOrd1.Items)
                    {
                        // 這兩個特殊選項, 需要最後勾選, 不然會造成其他正常選項無法勾選, 以及太早上色
                        if (item.Text == "部訂必修專業科目" || item.Text == "部訂必修實習科目")
                        {
                            if (Configure.TagRank1SubjectList.Contains(item.Text))
                            {
                                // 先用temp存起來, 不然其他正常選項無法勾選
                                tmpSpecialListViewItem["lvwSubjectOrd1"] = item.Index;
                            }
                        }
                        else
                            item.Checked = Configure.TagRank1SubjectList.Contains(item.Text);
                    }
                    cboTagRank2.Text = Configure.TagRank2TagName;
                    foreach (ListViewItem item in lvwSubjectOrd2.Items)
                    {
                        // 這兩個特殊選項, 需要最後勾選, 不然會造成其他正常選項無法勾選, 以及太早上色
                        if (item.Text == "部訂必修專業科目" || item.Text == "部訂必修實習科目")
                        {
                            if (Configure.TagRank2SubjectList.Contains(item.Text))
                            {
                                // 先用temp存起來, 不然其他正常選項無法勾選
                                tmpSpecialListViewItem["lvwSubjectOrd2"] = item.Index;
                            }
                        }
                        else
                            item.Checked = Configure.TagRank2SubjectList.Contains(item.Text);
                    }

                    // 勾選特殊選項, 讓其他選項無法勾選, 以及上色
                    if (tmpSpecialListViewItem["lvwSubjectPri"] >= 0)
                    {
                        _SpecialListViewItem["lvwSubjectPri"] = tmpSpecialListViewItem["lvwSubjectPri"];
                        lvwSubjectPri.Items[_SpecialListViewItem["lvwSubjectPri"]].Checked = true;
                    }
                    if (tmpSpecialListViewItem["lvwSubjectOrd1"] >= 0)
                    {
                        _SpecialListViewItem["lvwSubjectOrd1"] = tmpSpecialListViewItem["lvwSubjectOrd1"];
                        lvwSubjectOrd1.Items[_SpecialListViewItem["lvwSubjectOrd1"]].Checked = true;
                    }
                    if (tmpSpecialListViewItem["lvwSubjectOrd2"] >= 0)
                    {
                        _SpecialListViewItem["lvwSubjectOrd2"] = tmpSpecialListViewItem["lvwSubjectOrd2"];
                        lvwSubjectOrd2.Items[_SpecialListViewItem["lvwSubjectOrd2"]].Checked = true;
                    }

                    chk計算學業成績排名.Checked = Configure.WithCalSemesterScoreRank;

                    if (chkGrade1.Checked == false && chkGrade2.Checked == false && chkGrade3.Checked == false && chkGrade4.Checked == false)
                    {
                        if (chkGrade1.Text == Configure.SortGradeYear)
                            chkGrade1.Checked = true;

                        if (chkGrade2.Text == Configure.SortGradeYear)
                            chkGrade2.Checked = true;

                        if (chkGrade3.Text == Configure.SortGradeYear)
                            chkGrade3.Checked = true;

                        if (chkGrade4.Text == Configure.SortGradeYear)
                            chkGrade4.Checked = true;

                    }
                    // 成績年級學期
                    if (Configure.RankFilterGradeSemeterList != null)
                        foreach (Control cr in gpGradeSemester.Controls)
                        {
                            CheckBoxX cb = cr as CheckBoxX;
                            if (cb != null)
                                cb.Checked = Configure.RankFilterGradeSemeterList.Contains(cb.Text);

                        }

                    // 採計成績
                    if (Configure.RankFilterUseScoreList != null)
                        foreach (Control cr in gpUseScore.Controls)
                        {
                            CheckBoxX cb = cr as CheckBoxX;
                            if (cb != null)
                                cb.Checked = Configure.RankFilterUseScoreList.Contains(cb.Text);

                        }
                }
                else
                {
                    Configure = null;
                    chkGrade1.Checked = true;
                    cboRankRilter.SelectedIndex = -1;
                    cboTagRank1.SelectedIndex = -1;
                    cboTagRank2.SelectedIndex = -1;
                    foreach (ListViewItem item in lvwSubjectPri.Items)
                    {
                        item.Checked = false;
                    }
                    foreach (ListViewItem item in lvwSubjectOrd1.Items)
                    {
                        item.Checked = false;
                    }
                    foreach (ListViewItem item in lvwSubjectOrd2.Items)
                    {
                        item.Checked = false;
                    }

                    foreach (Control cr in gpGradeSemester.Controls)
                    {
                        CheckBoxX cb = cr as CheckBoxX;
                        if (cb != null)
                            cb.Checked = false;

                    }

                    // 採計成績                    
                    foreach (Control cr in gpUseScore.Controls)
                    {
                        CheckBoxX cb = cr as CheckBoxX;
                        if (cb != null)
                            cb.Checked = false;

                    }
                }
            }
            _NewCboConfigName = cboConfigure.Text;
        }

        private void cboConfigure_SelectedIndexChanged(object sender, EventArgs e)
        {
            readData();
        }

        private void btnCopy1_Click(object sender, EventArgs e)
        {
            List<string> selNameList = new List<string>();
            int itemIndex = int.MinValue;

            foreach (ListViewItem lvi in lvwSubjectPri.Items)
                if (lvi.Checked)
                    selNameList.Add(lvi.Text);

            foreach (ListViewItem lvi in lvwSubjectOrd1.Items)
            {
                if(selNameList.Contains(lvi.Text))
                {
                    // 這兩個特殊選項, 需要最後勾選, 不然會造成其他正常選項無法勾選, 以及太早上色
                    if (lvi.Text == "部訂必修專業科目" || lvi.Text == "部訂必修實習科目")
                        itemIndex = lvi.Index;
                    else
                        lvi.Checked=true;
                }
                else
                    lvi.Checked = false;
            }

            // 勾選特殊選項, 讓其他選項無法勾選, 以及上色
            if (itemIndex >= 0)
            {
                _SpecialListViewItem["lvwSubjectOrd1"] = itemIndex;
                lvwSubjectOrd1.Items[itemIndex].Checked = true;
            }
            else
            {
                _SpecialListViewItem["lvwSubjectOrd1"] = int.MinValue;
            }
        }

        private void buttonX2_Click(object sender, EventArgs e)
        {
            List<string> selNameList = new List<string>();
            int itemIndex = int.MinValue;
            foreach (ListViewItem lvi in lvwSubjectPri.Items)
                if (lvi.Checked)
                    selNameList.Add(lvi.Text);

            foreach (ListViewItem lvi in lvwSubjectOrd2.Items)
                if (selNameList.Contains(lvi.Text))
                {
                    // 這兩個特殊選項, 需要最後勾選, 不然會造成其他正常選項無法勾選, 以及太早上色
                    if (lvi.Text == "部訂必修專業科目" || lvi.Text == "部訂必修實習科目")
                        itemIndex = lvi.Index;
                    else
                        lvi.Checked=true;
                }
                else
                    lvi.Checked = false;

            // 勾選特殊選項, 讓其他選項無法勾選, 以及上色
            if (itemIndex >= 0)
            {
                _SpecialListViewItem["lvwSubjectOrd2"] = itemIndex;
                lvwSubjectOrd2.Items[itemIndex].Checked = true;
            }
            else
            {
                _SpecialListViewItem["lvwSubjectOrd2"] = int.MinValue;
            }         
        }

        #region ListView的事件
        /// <summary>
        /// 假如有勾選"部訂必修專業科目"或"部訂必修實習科目", 不可改變勾選的結果
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ListViewItemCheck(object sender, ItemCheckEventArgs e)
        {
            string listViewName = (sender as ListView).Name;
            if ((_SpecialListViewItem[listViewName] >= 0) &&
                (_SpecialListViewItem[listViewName] != e.Index))
            {
                e.NewValue = e.CurrentValue;
            }
        }
        /// <summary>
        /// 處理勾選"部訂必修專業科目"或"部訂必修實習科目"
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        void ListViewItemChecked(object sender, ItemCheckedEventArgs e)
        {
            ListView listView = sender as ListView;
            string listViewName = listView.Name;

            if (e.Item.Text == "部訂必修專業科目" || e.Item.Text == "部訂必修實習科目")
            {
                if (e.Item.Checked == true)
                {
                    if (_SpecialListViewItem[listViewName] < 0)
                    {
                        _SpecialListViewItem[listViewName] = e.Item.Index;
                    }
                }
                else
                {
                    if (_SpecialListViewItem[listViewName] == e.Item.Index)
                    {
                        _SpecialListViewItem[listViewName] = int.MinValue;
                    }
                }
            }

            // 幫選項上顏色
            foreach (ListViewItem item in listView.Items)
            {
                if (_SpecialListViewItem[listViewName] < 0)
                {
                    item.ForeColor = Color.Black;
                }
                else
                {
                    if (_SpecialListViewItem[listViewName] == item.Index)
                        item.ForeColor = Color.Black;
                    else
                        item.ForeColor = Color.Gray;
                }
            }
        }
        #endregion ListView的事件

    }
}
