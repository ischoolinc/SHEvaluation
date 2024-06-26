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
using Aspose.Cells;

namespace SmartSchool.Evaluation.Content.ChangeSchoolYear
{
    public partial class frmChangeSemesterSchoolYearHasData : BaseForm
    {

        // 有重複清單
        private List<SubjectScoreInfo> HasDataList = new List<SubjectScoreInfo>();
        private string Message = "";
        private StudentInfo studentInfo = null;


        public frmChangeSemesterSchoolYearHasData()
        {
            InitializeComponent();
            studentInfo = new StudentInfo();
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void lnkHasDataList_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            lnkHasDataList.Enabled = false;

            if (HasDataList.Count > 0)
            {
                try
                {
                    Workbook wb = new Workbook();                    
                    wb.Open(new System.IO.MemoryStream(Properties.Resources.查看重複清單樣板));

                    int rowIdx = 1;
                    foreach(SubjectScoreInfo ss in HasDataList)
                    {
                        // 學號 0
                        wb.Worksheets[0].Cells[rowIdx, 0].PutValue(studentInfo.StudentNumber);

                        // 班級 1
                        wb.Worksheets[0].Cells[rowIdx, 1].PutValue(studentInfo.ClassName);

                        // 座號 2
                        wb.Worksheets[0].Cells[rowIdx, 2].PutValue(studentInfo.SeatNo);

                        // 姓名 3
                        wb.Worksheets[0].Cells[rowIdx, 3].PutValue(studentInfo.StudentName);

                        // 學年度 4
                        wb.Worksheets[0].Cells[rowIdx, 4].PutValue(ss.SchoolYear);

                        // 學期 5
                        wb.Worksheets[0].Cells[rowIdx, 5].PutValue(ss.Semester);

                        // 成績年級 6
                        wb.Worksheets[0].Cells[rowIdx, 6].PutValue(ss.GradeYear);

                        // 科目名稱 7
                        wb.Worksheets[0].Cells[rowIdx, 7].PutValue(ss.SubjectName);

                        // 科目級別 8
                        wb.Worksheets[0].Cells[rowIdx, 8].PutValue(ss.SubjectLevel);

                        // 學分數 9
                        wb.Worksheets[0].Cells[rowIdx, 9].PutValue(ss.Credit);

                        // 校部定 10
                        wb.Worksheets[0].Cells[rowIdx, 10].PutValue(ss.RequiredBy);

                        // 必選修 11
                        wb.Worksheets[0].Cells[rowIdx, 11].PutValue(ss.Required);

                        rowIdx++;

                    }

                    Utility.ExprotXls("重複科目名稱級別", wb);

                }
                catch (Exception ex)
                {
                    MsgBox.Show(ex.Message);

                }
            }
            lnkHasDataList.Enabled = true;
        }

        private void frmChangeSemesterSchoolYearHasData_Load(object sender, EventArgs e)
        {
            this.MaximumSize = this.MaximumSize = this.Size;
            lblMsg.Text = Message;
        }

        public void SetMessage(string SchoolYear, string Semester, string GradeYear)
        {
            Message = string.Format(@"將{0}學年度第{1}學期成績年級{2}年級 已有重複科目名稱+級別，請確認後再進行作業。", SchoolYear, Semester, GradeYear);
        }

        public void SetHasDataList(List<SubjectScoreInfo> HasDataList)
        {
            this.HasDataList = HasDataList;
        }

        public void SetStudentInfo(StudentInfo studentInfo)
        {
            this.studentInfo = studentInfo;
        }
    }
}
