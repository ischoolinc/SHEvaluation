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

namespace StudentDuplicateSubjectCheck
{
    public partial class HasPassScoreForm : BaseForm
    {
        List<string> studentIDList = new List<string>();
        BackgroundWorker _bgExport = new BackgroundWorker();
        List<DataRow> _dataRowList = new List<DataRow>();


        public HasPassScoreForm()
        {
            InitializeComponent();
            _bgExport.DoWork += _bgExport_DoWork;
            _bgExport.RunWorkerCompleted += _bgExport_RunWorkerCompleted;
            _bgExport.ProgressChanged += _bgExport_ProgressChanged;
            _bgExport.WorkerReportsProgress = true;
        }

        private void _bgExport_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            FISCA.Presentation.MotherForm.SetStatusBarMessage("清單產生中...", e.ProgressPercentage);
        }

        private void _bgExport_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (e.Error == null)
            {
                Workbook wb = e.Result as Workbook;
                wb.FileName = "已有及格補考標準清單";

                string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
                if (!Directory.Exists(path))
                    Directory.CreateDirectory(path);
                path = Path.Combine(path, "已有及格補考標準清單" + ".xlsx");

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
                    wb.Save(path, SaveFormat.Xlsx);
                    System.Diagnostics.Process.Start(path);
                }
                catch
                {
                    System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = "已有及格補考標準清單.xlsx";
                    sd.Filter = "Excel檔案 (*.xlsx)|*.xlsx|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        try
                        {
                            wb.Save(path, SaveFormat.Xlsx);
                            System.Diagnostics.Process.Start(path);

                        }
                        catch
                        {
                            FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                            return;
                        }
                    }
                }
            }
            btnExportList.Enabled = true;
            FISCA.Presentation.MotherForm.SetStatusBarMessage("");
        }

        private void _bgExport_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                Workbook wb = new Workbook(new System.IO.MemoryStream(Properties.Resources.已有及格補考標準清單));
                int rowIdx = 1;
                foreach (DataRow dr in _dataRowList)
                {
                    wb.Worksheets[0].Cells[rowIdx, 0].PutValue(dr["school_year"].ToString());
                    wb.Worksheets[0].Cells[rowIdx, 1].PutValue(dr["semester"].ToString());
                    wb.Worksheets[0].Cells[rowIdx, 2].PutValue(dr["course_name"].ToString());
                    wb.Worksheets[0].Cells[rowIdx, 3].PutValue(dr["student_name"].ToString());
                    wb.Worksheets[0].Cells[rowIdx, 4].PutValue(dr["student_number"].ToString());
                    wb.Worksheets[0].Cells[rowIdx, 5].PutValue(dr["passing_standard"].ToString());
                    wb.Worksheets[0].Cells[rowIdx, 6].PutValue(dr["makeup_standard"].ToString());
                    wb.Worksheets[0].Cells[rowIdx, 7].PutValue(dr["remark"].ToString());
                    rowIdx++;
                }

                e.Result = wb;
            }
            catch (Exception ex)
            {
                throw;
            }
        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Yes;
        }

        private void btnAddTemp_Click(object sender, EventArgs e)
        {
            List<string> addIDList = new List<string>();
            if (studentIDList.Count > 0)
            {
                foreach (string id in studentIDList)
                {
                    if (!K12.Presentation.NLDPanels.Student.TempSource.Contains(id))
                        addIDList.Add(id);
                }

                K12.Presentation.NLDPanels.Student.AddToTemp(addIDList);
                MsgBox.Show("已於[學生待處理]加入" + studentIDList.Count + "名學生");
            }
        }

        public void SetDataRows(List<DataRow> dataRows)
        {
            _dataRowList = dataRows;
            dgData.Rows.Clear();
            lblMsg.Text = "共 " + dataRows.Count + " 筆";
            studentIDList.Clear();
            foreach (DataRow dr in dataRows)
            {
                int rowIdx = dgData.Rows.Add();
                string sid = dr["student_id"].ToString();
                dgData.Rows[rowIdx].Tag = sid;
                if (!studentIDList.Contains(sid))
                    studentIDList.Add(sid);

                dgData.Rows[rowIdx].Cells[colSchoolYear.Index].Value = dr["school_year"].ToString();
                dgData.Rows[rowIdx].Cells[colSemester.Index].Value = dr["semester"].ToString();
                dgData.Rows[rowIdx].Cells[colCourseName.Index].Value = dr["course_name"].ToString();
                dgData.Rows[rowIdx].Cells[colStudentName.Index].Value = dr["student_name"].ToString();
                dgData.Rows[rowIdx].Cells[colStudentNumber.Index].Value = dr["student_number"].ToString();
                dgData.Rows[rowIdx].Cells[colPassingStandard.Index].Value = dr["passing_standard"].ToString();
                dgData.Rows[rowIdx].Cells[colMakeupStandard.Index].Value = dr["makeup_standard"].ToString();
                dgData.Rows[rowIdx].Cells[colRemark.Index].Value = dr["remark"].ToString();
            }
        }

        private void btnExportList_Click(object sender, EventArgs e)
        {
            btnExportList.Enabled = false;
            _bgExport.RunWorkerAsync();
        }

        private void HasPassScoreForm_Load(object sender, EventArgs e)
        {

        }

        private void lblMsg_Click(object sender, EventArgs e)
        {

        }

        private void btnWrite_Click(object sender, EventArgs e)
        {
            DialogResult = DialogResult.Yes;
        }
    }
}
