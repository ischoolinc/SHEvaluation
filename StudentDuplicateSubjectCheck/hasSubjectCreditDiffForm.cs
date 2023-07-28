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
    public partial class hasSubjectCreditDiffForm : BaseForm
    {
        Dictionary<string, List<GPlanSubjectCreditDifInfo>> dataDict;

        public hasSubjectCreditDiffForm()
        {
            InitializeComponent();
            dataDict = new Dictionary<string, List<GPlanSubjectCreditDifInfo>>();

        }

        private void btnExit_Click(object sender, EventArgs e)
        {
            this.DialogResult = DialogResult.Cancel;
        }

        private void btnExportList_Click(object sender, EventArgs e)
        {
            btnExportList.Enabled = false;
            try
            {
                // Excel 樣板
                Workbook wb = new Workbook(new System.IO.MemoryStream(Properties.Resources.課程規劃科目同年級學分數不同));
                int rowIdx = 1;

                // 讀取 DataGrid 畫面資料
                foreach (DataGridViewRow drv in dgData.Rows)
                {
                    if (drv.IsNewRow)
                        continue;

                    // 課程規劃名稱 0
                    wb.Worksheets[0].Cells[rowIdx, 0].PutValue(drv.Cells[colGplanName.Index].Value + "");

                    // 年級 1
                    wb.Worksheets[0].Cells[rowIdx, 1].PutValue(drv.Cells[colGradeYear.Index].Value + "");
                    // 科目名稱 2
                    wb.Worksheets[0].Cells[rowIdx, 2].PutValue(drv.Cells[colSubjectName.Index].Value + "");

                    // 上學期學分數 3
                    wb.Worksheets[0].Cells[rowIdx, 3].PutValue(drv.Cells[colCreditA.Index].Value + "");

                    // 下學期學分數 4
                    wb.Worksheets[0].Cells[rowIdx, 4].PutValue(drv.Cells[colCreditB.Index].Value + "");

                    // 指定學年科目名稱
                    wb.Worksheets[0].Cells[rowIdx, 5].PutValue(drv.Cells[colSpecifySubjectName.Index].Value + "");

                    rowIdx++;
                }
                wb.Worksheets[0].AutoFitColumns();

                Utility.ExportXls("課程規劃科目同年級學分數不同", wb);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            btnExportList.Enabled = true;
        }

        public void SetData(Dictionary<string, List<GPlanSubjectCreditDifInfo>> data)
        {
            dataDict = data;
            dgData.Rows.Clear();

            List<string> tmpStr = new List<string>();
            foreach (string key in data.Keys)
            {
                int rowIdx = dgData.Rows.Add();
                tmpStr.Clear();
                foreach (GPlanSubjectCreditDifInfo gi in data[key])
                {
                    dgData.Rows[rowIdx].Tag = gi;
                    dgData.Rows[rowIdx].Cells[colGplanName.Index].Value = gi.GPlanName;
                    dgData.Rows[rowIdx].Cells[colGradeYear.Index].Value = gi.GradeYear;
                    dgData.Rows[rowIdx].Cells[colSubjectName.Index].Value = gi.SubjectName;
                    if (gi.Semester == "1")
                        dgData.Rows[rowIdx].Cells[colCreditA.Index].Value = gi.Credit;
                    else
                        dgData.Rows[rowIdx].Cells[colCreditB.Index].Value = gi.Credit;

                    if (!tmpStr.Contains(gi.SpecifySubjectName))
                        tmpStr.Add(gi.SpecifySubjectName);
                }
                dgData.Rows[rowIdx].Cells[colSpecifySubjectName.Index].Value = string.Join(",", tmpStr.ToArray());
            }

            lblMsg.Text = "共 " + dgData.Rows.Count + " 筆";

        }
    }
}
