using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using FISCA.Presentation.Controls;
using StudentCourseScoreCheck100_SH.DAO;
using Aspose.Cells;

namespace StudentCourseScoreCheck100_SH
{
    public partial class StudentCourseScoreForm : BaseForm
    {

        BackgroundWorker _bgWorker;
        int _SchoolYear = 0;
        int _Semester = 0;
        List<string> _SelGradeYear = new List<string>();
        List<StudentCourseScoreBase> _StudentCourseScoreBaseList;
        // 資料存取用
        DataTable _dt = new DataTable();

        public StudentCourseScoreForm()
        {
            InitializeComponent();
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += new DoWorkEventHandler(_bgWorker_DoWork);
            _bgWorker.ProgressChanged += new ProgressChangedEventHandler(_bgWorker_ProgressChanged);
            _bgWorker.WorkerReportsProgress = true;
            _bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(_bgWorker_RunWorkerCompleted);

        }

        void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("課程成績評量檢查中..", e.ProgressPercentage);
        }

        void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (_dt.Rows.Count > 0)
            {
                Workbook wb = new Workbook();
                string name = _SchoolYear + "學年度第" + _Semester + "學期" + string.Join(",", _SelGradeYear.ToArray()) + "年級,課程成績評量檢查";
                wb.Worksheets[0].PageSetup.SetHeader(0, name);
                Utility.CompletedXls(name, _dt, wb);
            }
            else
                FISCA.Presentation.Controls.MsgBox.Show("沒有資料!");
            btnPrint.Enabled = true;

            FISCA.Presentation.MotherForm.SetStatusBarMessage("課程成績評量檢查完成..", 100);
        }

        void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(1);
            // 讀取資料
            _StudentCourseScoreBaseList = QueryData.GetStudentClassBase(_SchoolYear, _Semester, _SelGradeYear);
            _bgWorker.ReportProgress(30);
            _dt.Clear();
            _dt.Columns.Clear();
            // 建立DataTable 欄位名稱           
            _dt.Columns.Add("年級");
            _dt.Columns.Add("學號");
            _dt.Columns.Add("班級");
            _dt.Columns.Add("座號");
            _dt.Columns.Add("姓名");
            _dt.Columns.Add("課程名稱");
            _dt.Columns.Add("授課教師");
            _dt.Columns.Add("課程成績");
            // 動態欄位            
            QueryData._StudentExamName.Sort();
            foreach (string name in QueryData._StudentExamName)
                _dt.Columns.Add(name);

            _bgWorker.ReportProgress(50);

            // 取得缺免內容設定
            Dictionary<string, string> examUseTextReportValueDict = QueryData.GetExamUseTextReportValue();

            // 填入資料
            foreach (StudentCourseScoreBase scsb in _StudentCourseScoreBaseList)
            {
                foreach (KeyValuePair<string, string> courseName in scsb.CourseNameDict)
                {
                    DataRow row = _dt.NewRow();
                    row["年級"] = scsb.GradeYear;
                    row["學號"] = scsb.StudentNumber;
                    row["班級"] = scsb.ClassName;
                    row["座號"] = scsb.SeatNo;
                    row["姓名"] = scsb.Name;
                    row["課程名稱"] = courseName.Value;
                    if (scsb.CourseTeacherDict.ContainsKey(courseName.Key))
                        row["授課教師"] = scsb.CourseTeacherDict[courseName.Key];

                    if (scsb.CousreScoreDict.ContainsKey(courseName.Key))
                        row["課程成績"] = scsb.CousreScoreDict[courseName.Key];

                    if (scsb.ExamScoreDict.ContainsKey(courseName.Key))
                    {
                        foreach (KeyValuePair<string, decimal> exScore in scsb.ExamScoreDict[courseName.Key])
                        {
                            // 判斷是否缺免
                            if (exScore.Value == -1 || exScore.Value == -2)
                            {
                                if (scsb.ExamScoreTextDict[courseName.Key].ContainsKey(exScore.Key))
                                {
                                    string eValue = scsb.ExamScoreTextDict[courseName.Key][exScore.Key];

                                    // 缺考原因
                                    if (examUseTextReportValueDict.ContainsKey(eValue))
                                        row[exScore.Key] = examUseTextReportValueDict[eValue];
                                    else
                                        row[exScore.Key] = eValue;
                                }
                                else
                                    row[exScore.Key] = exScore.Value;
                            }
                            else
                            {
                                row[exScore.Key] = exScore.Value;
                            }
                        }

                    }
                    _dt.Rows.Add(row);
                }
            }
            _bgWorker.ReportProgress(90);
        }

        private void StudentCourseScoreForm_Load(object sender, EventArgs e)
        {
            lvData.Items.Clear();
            // 預設學年度學期
            iptSchoolYear.Value = int.Parse(K12.Data.School.DefaultSchoolYear);
            iptSemester.Value = int.Parse(K12.Data.School.DefaultSemester);
            List<string> grList = DAO.QueryData.GetClassGradeYearList();

            foreach (string str in grList)
            {
                ListViewItem lvi = new ListViewItem();
                lvi.Text = str + "年級";
                lvi.Checked = true;
                lvi.Tag = str;
                lvData.Items.Add(lvi);
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void btnPrint_Click(object sender, EventArgs e)
        {
            _SchoolYear = iptSchoolYear.Value;
            _Semester = iptSemester.Value;
            _SelGradeYear.Clear();
            foreach (ListViewItem lvi in lvData.CheckedItems)
                _SelGradeYear.Add(lvi.Tag.ToString());

            if (_SelGradeYear.Count == 0)
            {
                FISCA.Presentation.Controls.MsgBox.Show("請勾選年級!");
                return;
            }
            btnPrint.Enabled = false;
            _bgWorker.RunWorkerAsync();
        }
    }
}
