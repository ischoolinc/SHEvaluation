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

namespace SmartSchool.Evaluation.Content.ChangeSchoolYear
{
    public partial class frmChangeSemesterSchoolYearMain : BaseForm
    {
        // 來源訊息
        private string SourceMessage = "";

        // 學生資訊
        private StudentInfo studentInfo = null;

        private string SourceSchoolYear = "", Semester = "", GradeYear = "";

        public frmChangeSemesterSchoolYearMain()
        {
            InitializeComponent();
            studentInfo = new StudentInfo();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.No;
        }

        public void SetStudentInfo(StudentInfo studentInfo)
        {
            this.studentInfo = studentInfo;
        }

        private void btnChange_Click(object sender, EventArgs e)
        {
            btnChange.Enabled = false;

            try
            {
                // 先取得這位學生有成績學年度學期
                Dictionary<string, string> hasSemsScoreSchoolYearSemesterDict = SemesterScoreDataTransfer.GetStudentHasSemsScoreSchoolYearSemesterByID(studentInfo.StudentID);

                string ChangGradeYear = GradeYear;

                string sKey = cboSchoolYear.Text + "_" + Semester;
                if (hasSemsScoreSchoolYearSemesterDict.ContainsKey(sKey))
                {
                    ChangGradeYear = hasSemsScoreSchoolYearSemesterDict[sKey];
                }

                // 開啟確認視窗
                frmChangeSemesterSchoolYearMsg fmm = new frmChangeSemesterSchoolYearMsg();
                // 傳入來成績資料來源
                fmm.SetSourceMessage(SourceSchoolYear, Semester, GradeYear);
                // 傳入調整後成績學年度
                fmm.SetChangeMessage(cboSchoolYear.Text, Semester, ChangGradeYear);

                // 檢查來源與調整學年度是否相同，相同不處理
                if (SourceSchoolYear == cboSchoolYear.Text)
                {
                    MsgBox.Show("選擇成績學年度與欲調整學年度相同，無法進行調整。");
                    btnChange.Enabled = true;
                    return;
                }

                // 傳入學生
                fmm.SetStudentInfo(studentInfo);

                if (fmm.ShowDialog() == DialogResult.Yes)
                {
                    this.DialogResult = DialogResult.Yes;
                }
                else
                {
                    this.DialogResult = DialogResult.No;
                }
            }
            catch (Exception ex)
            {
                MsgBox.Show("調整學年度過程發生錯誤：" + ex.Message);
            }

            btnChange.Enabled = true;
        }

        private void frmChangeSemesterSchoolYearMain_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MinimumSize = this.Size;
            // 傳入訊息
            labelSourceMsg.Text = SourceMessage;

            try
            {
                // 學年度選項
                int defSchoolYear = int.Parse(K12.Data.School.DefaultSchoolYear);
                for (int i = defSchoolYear + 3; i >= defSchoolYear - 3; i--)
                    cboSchoolYear.Items.Add(i + "");

            }
            catch (Exception ex)
            { Console.WriteLine(ex.Message); }


            // 填入預設學年度            
            cboSchoolYear.Text = K12.Data.School.DefaultSchoolYear;
        }

        // 設定來源學年度、學期、年級
        public void SetSourceInfo(string SchoolYear, string Semester, string GradeYear)
        {
            SourceMessage = string.Format("{0}學年度第{1}學期成績年級{2}年級", SchoolYear, Semester, GradeYear);
            SourceSchoolYear = SchoolYear;
            this.Semester = Semester;
            this.GradeYear = GradeYear;
        }
    }
}
