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
using System.IO;
using FISCA.Presentation;
using StudentDuplicateSubjectCheck.DAO;

namespace StudentDuplicateSubjectCheck
{
    public partial class hasSpecifySubjectNameForm : BaseForm
    {

        // 加入課程代處理
        List<string> AddTempList;

        // 匯出清單
        BackgroundWorker _bgExport;

        Dictionary<string, CourseSpecifySubjectNameInfo> dataDict;

        public hasSpecifySubjectNameForm()
        {
            InitializeComponent();
            AddTempList = new List<string>();
            dataDict = new Dictionary<string, CourseSpecifySubjectNameInfo>();
            _bgExport = new BackgroundWorker();
            _bgExport.DoWork += _bgExport_DoWork;
            _bgExport.RunWorkerCompleted += _bgExport_RunWorkerCompleted;
            _bgExport.ProgressChanged += _bgExport_ProgressChanged;
            _bgExport.WorkerReportsProgress = true;
        }

        private void _bgExport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("課程清單產生中...", e.ProgressPercentage);
        }

        private void _bgExport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                try
                {
                    Workbook wb = e.Result as Workbook;
                    Utility.ExportXls("課程已有指定學年科目名稱清單", wb);
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            btnExportList.Enabled = true;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        private void _bgExport_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                Workbook wb = new Workbook(new System.IO.MemoryStream(Properties.Resources.課程已有指定學年科目名稱樣板));
                int rowIdx = 1;
                foreach (CourseSpecifySubjectNameInfo ci in dataDict.Values)
                {
                    // 課程系統編號
                    wb.Worksheets[0].Cells[rowIdx, 0].PutValue(ci.CourseID);
                    // 學年度
                    wb.Worksheets[0].Cells[rowIdx, 1].PutValue(ci.SchoolYear);
                    // 學期
                    wb.Worksheets[0].Cells[rowIdx, 2].PutValue(ci.Semester);
                    // 課程名稱
                    wb.Worksheets[0].Cells[rowIdx, 3].PutValue(ci.CourseName);
                    // 指定學年科目名稱
                    wb.Worksheets[0].Cells[rowIdx, 4].PutValue(ci.SpecifySubjectName);

                    rowIdx++;
                }

                wb.Worksheets[0].AutoFitColumns();
                
                e.Result = wb;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Cancel;
        }

        private void btnWrite_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Yes;
        }

        private void btnAddTemp_Click(object sender, EventArgs e)
        {
            AddTempList.Clear();
            if (dataDict.Count > 0)
            {
                foreach (string id in dataDict.Keys)
                {
                    if (!K12.Presentation.NLDPanels.Course.TempSource.Contains(id))
                        AddTempList.Add(id);
                }

                K12.Presentation.NLDPanels.Course.AddToTemp(AddTempList);
                MsgBox.Show("已於[課程待處理]加入" + dataDict.Count + "筆課程。");
            }
        }

        private void btnExportList_Click(object sender, EventArgs e)
        {
            btnExportList.Enabled = false;
            _bgExport.RunWorkerAsync();
        }

        public void SetData(Dictionary<string, CourseSpecifySubjectNameInfo> data)
        {
            dataDict = data;
            dgData.Rows.Clear();
            lblMsg.Text = "共 " + data.Count + " 筆";

            foreach (CourseSpecifySubjectNameInfo ci in data.Values)
            {
                int rowIdx = dgData.Rows.Add();
                string cid = ci.CourseID;
                dgData.Rows[rowIdx].Tag = cid;
                dgData.Rows[rowIdx].Cells[colSchoolYear.Index].Value = ci.SchoolYear;
                dgData.Rows[rowIdx].Cells[colSemester.Index].Value = ci.Semester;
                dgData.Rows[rowIdx].Cells[colCourseName.Index].Value = ci.CourseName;
                dgData.Rows[rowIdx].Cells[colSpecifySubjectName.Index].Value = ci.SpecifySubjectName;
            }
        }

    }
}
