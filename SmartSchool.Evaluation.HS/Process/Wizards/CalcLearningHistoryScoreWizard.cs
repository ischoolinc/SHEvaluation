using DevComponents.DotNetBar;
using DevComponents.DotNetBar.Rendering;
using FISCA.Data;
using FISCA.DSAUtil;
using SmartSchool.Customization.Data;
using SmartSchool.Evaluation.Process.Wizards.LearningHistory;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;



namespace SmartSchool.Evaluation.Process.Wizards
{
    public partial class CalcLearningHistoryScoreWizard  : SmartSchool.Common.BaseForm
    {
        private ErrorViewer _ErrorViewer = new ErrorViewer();
        private WarnViewer _WarnViewer;

        private StudentLearningHistoryProcessor _processor;

        private BackgroundWorker runningBackgroundWorker = new BackgroundWorker();

        private SelectType _Type;

        private List<string> warningList = new List<string>();

        private List<StudentRecord> resultList; // 儲存最後結果的List

        private bool hasWarning = false;
        private bool hasError = false;


        public CalcLearningHistoryScoreWizard(SelectType type)
        {
            _Type = type;
            InitializeComponent();
            _processor = new StudentLearningHistoryProcessor();

            #region 設定Wizard會跟著Style跑
            //this.wizard1.FooterStyle.ApplyStyle(( GlobalManager.Renderer as Office2007Renderer ).ColorTable.GetClass(ElementStyleClassKeys.RibbonFileMenuBottomContainerKey));
            this.wizard1.HeaderStyle.ApplyStyle((GlobalManager.Renderer as Office2007Renderer).ColorTable.GetClass(ElementStyleClassKeys.RibbonFileMenuBottomContainerKey));
            this.wizard1.FooterStyle.BackColorGradientAngle = -90;
            this.wizard1.FooterStyle.BackColorGradientType = eGradientType.Linear;
            this.wizard1.FooterStyle.BackColor = (GlobalManager.Renderer as Office2007Renderer).ColorTable.RibbonBar.Default.TopBackground.Start;
            this.wizard1.FooterStyle.BackColor2 = (GlobalManager.Renderer as Office2007Renderer).ColorTable.RibbonBar.Default.TopBackground.End;
            this.wizard1.BackColor = (GlobalManager.Renderer as Office2007Renderer).ColorTable.RibbonBar.Default.TopBackground.Start;
            this.wizard1.BackgroundImage = null;
            for (int i = 0; i < 5; i++)
            {
                (this.wizard1.Controls[1].Controls[i] as ButtonX).ColorTable = eButtonColor.OrangeWithBackground;
            }
            (this.wizard1.Controls[0].Controls[1] as System.Windows.Forms.Label).ForeColor = (GlobalManager.Renderer as Office2007Renderer).ColorTable.RibbonBar.MouseOver.TitleText;
            (this.wizard1.Controls[0].Controls[2] as System.Windows.Forms.Label).ForeColor = (GlobalManager.Renderer as Office2007Renderer).ColorTable.RibbonBar.Default.TitleText;
            #endregion
            this.wizard1.BackButtonWidth = 0;
            switch (_Type)
            {
                default:
                case SelectType.Student:
                    this.numericUpDown1.Value = SmartSchool.Customization.Data.SystemInformation.SchoolYear;
                    this.numericUpDown2.Value = SmartSchool.Customization.Data.SystemInformation.Semester;
                    labelX1.Text = ""; //選擇學年度學期
                    labelX2.Text = "學年度";
                    labelX3.Text = "學期";
                    labelX2.Top += 10;
                    labelX3.Top += 10;
                    numericUpDown1.Top += 10;
                    numericUpDown2.Top += 10;
                    this.numericUpDown1.Enabled = false;
                    this.numericUpDown2.Enabled = false;
                    break;
                case SelectType.GradeYearStudent:
                    this.Text = "產生" + SmartSchool.Customization.Data.SystemInformation.SchoolYear + "學年度第" + SmartSchool.Customization.Data.SystemInformation.Semester + "學期歷程成績";
                    this.numericUpDown1.Minimum = 1;
                    this.numericUpDown1.Maximum = 5;
                    this.numericUpDown1.Value = 1;
                    wizardPage1.PageTitle = labelX1.Text = "選擇年級";
                    labelX2.Text = "年級";
                    labelX3.Visible = false;
                    numericUpDown2.Visible = false;
                    break;
            }
        }

        // 產生歷程成績
        private void wizardPage2_AfterPageDisplayed(object sender, WizardPageChangeEventArgs e)
        {
            AccessHelper helper = new AccessHelper();
            List<StudentRecord> selectedStudents;
            int schooyYear;
            int semester;
            switch (_Type)
            {
                default:
                case SelectType.Student:
                    selectedStudents = helper.StudentHelper.GetSelectedStudent();
                    schooyYear = (int)numericUpDown1.Value;
                    semester = (int)numericUpDown2.Value;
                    break;
                case SelectType.GradeYearStudent:
                    selectedStudents = new List<StudentRecord>();
                    foreach (ClassRecord classrecord in helper.ClassHelper.GetAllClass())
                    {
                        int tryParseGradeYear;
                        if (int.TryParse(classrecord.GradeYear, out tryParseGradeYear) && tryParseGradeYear == (int)numericUpDown1.Value)
                            selectedStudents.AddRange(classrecord.Students);
                    }
                    schooyYear = SmartSchool.Customization.Data.SystemInformation.SchoolYear;
                    semester = SmartSchool.Customization.Data.SystemInformation.Semester;
                    break;
            }
            linkLabel1.Visible = false;
            linkLabel2.Visible = false;
            labelX4.Text = "產生學期歷程成績中...";
            runningBackgroundWorker = new BackgroundWorker();
            runningBackgroundWorker.WorkerSupportsCancellation = true;
            runningBackgroundWorker.WorkerReportsProgress = true;
            runningBackgroundWorker.ProgressChanged += RunningBackgroundWorker_ProgressChanged;
            runningBackgroundWorker.DoWork += RunningBackgroundWorker_DoWork;
            runningBackgroundWorker.RunWorkerCompleted += RunningBackgroundWorker_RunWorkerCompleted;

            runningBackgroundWorker.RunWorkerAsync(new object[] { schooyYear, semester, helper, selectedStudents });
        }

        private void RunningBackgroundWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (!((BackgroundWorker)sender).CancellationPending)
            {
                if (e.Result == null)
                {
                    hasError = true; //有錯誤，使用者將無法繼續計算
                    _WarnViewer = new WarnViewer(hasError);
                }
                else
                {
                    hasError = false;
                    _WarnViewer = new WarnViewer(hasError);
                }

                foreach (string s in warningList)
                {
                 //   _WarnViewer.SetMessage("課程:" + s, "科目名稱空白, 故此課程不列入計算");
                }

                if (hasError && !hasWarning) // 有錯誤 沒警告
                {
                    linkLabel1.Visible = true;
                    labelX4.Text = "計算失敗，請檢查錯誤訊息。";
                }
                else if (hasError && hasWarning) // 有錯誤 有警告
                {
                    linkLabel1.Visible = true;
                    linkLabel2.Visible = true; // 不再主動跳出視窗，警告訊息放在這裡供使用者點選打開檢查
                    labelX4.Text = "計算失敗，請先處理錯誤訊息，再檢視警告訊息";

                }
                else if (!hasError && hasWarning) // 沒錯誤 有警告
                {
                    linkLabel2.Visible = true; // 不再主動跳出視窗，警告訊息放在這裡供使用者點選打開檢查
                    labelX4.Text = "計算成功，請檢查警告訊息後，上傳成績";
                    resultList = (List<StudentRecord>)e.Result;
                }
                else // 沒錯誤 沒警告
                {
                    this.progressBarX2.Value = 100;
                    wizard1.SelectedPage = wizardPage4;
                    labelX5.Text = "學生學期歷程成績產生完成";
                    wizardPage4.FinishButtonEnabled = eWizardButtonState.True;
                    FISCA.Features.Invoke("StudentLearningHistoryDetailContent");
                }

            }
        }

        private void RunningBackgroundWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            warningList.Clear();// 將警告名單清空
            BackgroundWorker bkw = ((BackgroundWorker)sender);
            int schoolyear = (int)((object[])e.Argument)[0];
            int semester = (int)((object[])e.Argument)[1];
            AccessHelper helper = (AccessHelper)((object[])e.Argument)[2];
            List<StudentRecord> selectedStudents = (List<StudentRecord>)((object[])e.Argument)[3];            
           
            _processor.ProcessLearningHistory(helper,selectedStudents, schoolyear, semester, bkw);
            e.Result = selectedStudents;
        }

        private void RunningBackgroundWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            if (!((BackgroundWorker)sender).CancellationPending)
            {
                if (e.UserState != null)
                {
                    Dictionary<StudentRecord, List<string>> errormessages = (Dictionary<StudentRecord, List<string>>)e.UserState;
                    if (errormessages.Count > 0)
                    {
                        foreach (StudentRecord stu in errormessages.Keys)
                        {
                            _ErrorViewer.SetMessage(stu, errormessages[stu]);
                        }                        
                    }
                }
                this.progressBarX1.Value = e.ProgressPercentage;
            }
        }

        private void wizardPage1_FinishButtonClick(object sender, CancelEventArgs e)
        {
            this.Close();
            if (runningBackgroundWorker.IsBusy)
                runningBackgroundWorker.CancelAsync();
            this._ErrorViewer.Clear();
            this._ErrorViewer.Hide();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            this._ErrorViewer.ShowDialog();
        }

        private void wizard1_FinishButtonClick(object sender, CancelEventArgs e)
        {
            this.Close();
            if (runningBackgroundWorker.IsBusy)
                runningBackgroundWorker.CancelAsync();
            this._ErrorViewer.Clear();
            this._ErrorViewer.Hide();
        }
    }
}
