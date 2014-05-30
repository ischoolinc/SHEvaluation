using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Data;

namespace SemesterScoreReportNewEpost
{
    class Utility
    {
        public static void CompletedXlsCsv(string inputReportName, DataTable dt)
        {

            #region 儲存檔案
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".txt");

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

            //Workbook workbook = inputXls;
            StreamWriter sw = new StreamWriter(path, false, System.Text.Encoding.Unicode);

            // StringBuilder sb = new StringBuilder();

            //sb.Append(Encoding.UTF8.GetString(Encoding.UTF8.GetPreamble()));
            //    sb.Append(Encoding.Unicode.GetString(Encoding.Unicode.GetPreamble()));
            //if (workbook == null || workbook.Worksheets == null || workbook.Worksheets.Count == 0)
            //{
            //    MessageBox.Show("沒有資料！");
            //    return;
            //}
            //Cells cells = workbook.Worksheets[0].Cells;
            DataTable dataTable = dt;//workbook.Worksheets[0].Cells.ExportDataTable(cells.MinRow, cells.MinColumn, cells.Rows.Count, cells.Columns.Count);
            //  int row_count = 0;


            //dataTable.Rows.Cast<DataRow>().ToList().ForEach((x) =>
            //{
            //    row_count++;
            //    sb.Append(string.Join(",", x.ItemArray));
            //    if (!(row_count == dataTable.Rows.Count))
            //        sb.AppendLine();
            //});

            List<string> strList = new List<string>();
            foreach (DataColumn dc in dt.Columns)
                strList.Add(dc.ColumnName);

            sw.WriteLine(string.Join(",", strList.ToArray()));
            // sb.AppendLine(string.Join(",",strList.ToArray()));

            foreach (DataRow dr in dt.Rows)
            {
                List<string> subList = new List<string>();
                for (int col = 0; col < dt.Columns.Count; col++)
                {
                    subList.Add(dr[col].ToString());
                }
                sw.WriteLine(string.Join(",", subList.ToArray()));
                //sb.AppendLine(string.Join(",",subList.ToArray()));
            }

            //string out_put_string = sb.ToString();

            sw.Close();
            try
            {
                //FileStream fs = new FileStream(path, FileMode.Create);
                //fs.Write(Encoding.Unicode.GetBytes(out_put_string), 0, Encoding.Unicode.GetByteCount(out_put_string));
                //fs.Close();
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                try
                {
                    System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                    sd.Title = "另存新檔";
                    sd.FileName = reportName + ".txt";
                    sd.Filter = "txt檔案 (*.txt)|*.txt|所有檔案 (*.*)|*.*";
                    if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                    {
                        System.Diagnostics.Process.Start(sd.FileName);
                    }
                }
                catch
                {
                    FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                    return;
                }
            }
            #endregion
        
        }
    }
}
