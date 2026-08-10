using System;
using System.Collections.Generic;
using System.Text;
using FISCA.Authentication;
using FISCA.Presentation.Controls;
using SHImportExportScoreRank.DAO;
using SmartSchool.API.PlugIn;

namespace SHImportExportScoreRank.ImportExport
{
    public class ExportSchoolYearEntryRank
        : SmartSchool.API.PlugIn.Export.Exporter
    {
        private readonly List<string> _exportFields = new List<string>
        {
            "學號",
            "班級",
            "座號",
            "科別",
            "姓名",
            "學年度",
            "年級",
            "學年學業排名"
        };

        public ExportSchoolYearEntryRank()
        {
            this.Image = null;
            this.Text = "匯出學年分項排名";
        }

        public override void InitializeExport(
            SmartSchool.API.PlugIn.Export.ExportWizard wizard)
        {
            wizard.ExportableFields.AddRange(_exportFields);
            wizard.ExportPackage += ExportPackageHandler;
        }

        private void ExportPackageHandler(
            object sender,
            SmartSchool.API.PlugIn.Export.ExportPackageEventArgs e)
        {
            if (e.List == null || e.List.Count == 0)
                return;

            List<SchoolYearEntryRankRecord> records;
            try
            {
                RankOverrideDataAccess dataAccess = new RankOverrideDataAccess();
                records = dataAccess.GetSchoolYearEntryRanksByStudentIDs(e.List);
            }
            catch (Exception ex)
            {
                FISCA.LogAgent.ApplicationLog.Log(
                    "學年分項排名",
                    "匯出失敗",
                    BuildErrorLog(
                        "批次查詢 rank_override 失敗",
                        string.Join(",", e.List),
                        ex));

                MsgBox.Show("匯出學年分項排名失敗：" + ex.Message);
                return;
            }

            List<string> errors = new List<string>();
            int exportedCount = 0;

            foreach (SchoolYearEntryRankRecord record in records)
            {
                try
                {
                    RowData row = new RowData();
                    row.ID = record.StudentID ?? string.Empty;

                    foreach (string field in e.ExportFields)
                    {
                        if (!_exportFields.Contains(field))
                            continue;

                        row.Add(field, GetFieldValue(record, field));
                    }

                    e.Items.Add(row);
                    exportedCount++;
                }
                catch (Exception ex)
                {
                    errors.Add(string.Format(
                        "學生系統編號：{0}，rank_override.id：{1}，錯誤：{2}",
                        record.StudentID,
                        record.ID,
                        ex.Message));
                }
            }

            try
            {
                FISCA.LogAgent.ApplicationLog.Log(
                    "學年分項排名",
                    "匯出",
                    string.Format(
                        "匯出學年分項排名，共 {0} 筆，選取學生 {1} 人。執行帳號：{2}",
                        exportedCount,
                        e.List.Count,
                        DSAServices.UserAccount));
            }
            catch
            {
            }

            if (errors.Count > 0)
            {
                StringBuilder errorLog = new StringBuilder();
                errorLog.AppendLine(
                    "匯出學年分項排名部分資料轉換失敗，共 " +
                    errors.Count +
                    " 筆：");
                foreach (string error in errors)
                    errorLog.AppendLine(error);

                try
                {
                    FISCA.LogAgent.ApplicationLog.Log(
                        "學年分項排名",
                        "匯出轉換失敗",
                        errorLog.ToString());
                }
                catch
                {
                }

                MsgBox.Show(
                    "匯出學年分項排名完成，但有 " +
                    errors.Count +
                    " 筆資料轉換失敗，詳見系統日誌。");
            }
        }

        private static string GetFieldValue(
            SchoolYearEntryRankRecord record,
            string field)
        {
            switch (field)
            {
                case "學號":
                    return record.StudentNumber ?? string.Empty;
                case "班級":
                    return record.RankName ?? string.Empty;
                case "座號":
                    return record.SeatNo ?? string.Empty;
                case "科別":
                    return record.DepartmentName ?? string.Empty;
                case "姓名":
                    return record.StudentName ?? string.Empty;
                case "學年度":
                    return record.SchoolYear.HasValue
                        ? record.SchoolYear.Value.ToString()
                        : string.Empty;
                case "年級":
                    return record.GradeYear.HasValue
                        ? record.GradeYear.Value.ToString()
                        : string.Empty;
                case "學年學業排名":
                    return record.Rank.HasValue
                        ? record.Rank.Value.ToString()
                        : string.Empty;
                default:
                    return string.Empty;
            }
        }

        private static string BuildErrorLog(
            string title,
            string studentIds,
            Exception ex)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine(title);
            sb.AppendLine("目前使用者：" + DSAServices.UserAccount);
            sb.AppendLine("學生系統編號：" + studentIds);
            sb.AppendLine(ex.Message);
            sb.AppendLine(ex.StackTrace);
            return sb.ToString();
        }
    }
}
