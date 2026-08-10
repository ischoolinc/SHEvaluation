using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Campus.DocumentValidator;
using Campus.Import2014;
using FISCA.Authentication;
using SHImportExportScoreRank.DAO;

namespace SHImportExportScoreRank.ImportExport
{
    public class ImportSchoolYearEntryRank : ImportWizard
    {
        private ImportOption _option;

        public ImportSchoolYearEntryRank()
        {
            this.IsSplit = false;
        }

        public override ImportAction GetSupportActions()
        {
            return ImportAction.InsertOrUpdate;
        }

        public override string GetValidateRule()
        {
            const string resourceName =
                "SHImportExportScoreRank.ImportExport.ImportSchoolYearEntryRank.xml";

            Assembly assembly = Assembly.GetExecutingAssembly();
            using (Stream stream = assembly.GetManifestResourceStream(resourceName))
            {
                if (stream == null)
                {
                    throw new InvalidOperationException(
                        "找不到匯入驗證規則資源：" + resourceName);
                }

                using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                {
                    return reader.ReadToEnd();
                }
            }
        }

        public override void Prepare(ImportOption option)
        {
            _option = option;
        }

        public override string Import(List<IRowStream> Rows)
        {
            this.ImportProgress = 5;

            RankOverrideDataAccess dataAccess = new RankOverrideDataAccess();
            List<string> studentNumbers = new List<string>();

            foreach (IRowStream row in Rows)
            {
                if (row.Contains("學號"))
                {
                    string studentNumber = (row.GetValue("學號") ?? string.Empty).Trim();
                    if (!string.IsNullOrEmpty(studentNumber))
                        studentNumbers.Add(studentNumber);
                }
            }

            Dictionary<string, List<StudentLookupRecord>> studentDict =
                dataAccess.GetStudentsByStudentNumbers(studentNumbers);

            List<SchoolYearEntryRankImportRecord> records =
                new List<SchoolYearEntryRankImportRecord>();

            foreach (IRowStream row in Rows)
            {
                SchoolYearEntryRankImportRecord record;
                if (!TryBuildRecord(row, studentDict, out record))
                    continue;

                records.Add(record);
            }

            if (records.Count == 0)
            {
                this.ImportProgress = 100;
                return BuildSummary(0, 0);
            }

            List<string> duplicateMessages =
                dataAccess.FindDuplicateNaturalKeys(records);

            if (duplicateMessages.Count > 0)
            {
                StringBuilder error = new StringBuilder();
                error.AppendLine("資料庫中已存在重複的學年分項排名自然鍵，已中止匯入：");
                foreach (string message in duplicateMessages)
                    error.AppendLine(message);

                FISCA.LogAgent.ApplicationLog.Log(
                    "匯入學年分項排名",
                    "匯入失敗",
                    error.ToString());

                throw new Exception(error.ToString());
            }

            UpsertResult upsertResult;
            try
            {
                upsertResult = dataAccess.UpsertSchoolYearEntryRanks(records);
            }
            catch (Exception ex)
            {
                FISCA.LogAgent.ApplicationLog.Log(
                    "匯入學年分項排名",
                    "匯入失敗",
                    ex.Message + Environment.NewLine + ex.StackTrace);

                throw;
            }

            this.ImportProgress = 100;

            string summary = BuildSummary(
                upsertResult.InsertedCount,
                upsertResult.UpdatedCount);

            FISCA.LogAgent.ApplicationLog.Log(
                "匯入學年分項排名",
                "匯入",
                string.Format(
                    "匯入檔資料筆數：{0}{1}新增筆數：{2}{1}更新筆數：{3}{1}執行帳號：{4}{1}執行時間：{5:yyyy/MM/dd HH:mm:ss}",
                    Rows.Count,
                    Environment.NewLine,
                    upsertResult.InsertedCount,
                    upsertResult.UpdatedCount,
                    DSAServices.UserAccount,
                    DateTime.Now));

            try
            {
                FISCA.Features.Invoke("SchoolYearEntryRankDetailContent");
            }
            catch
            {
            }

            return summary;
        }

        private static bool TryBuildRecord(
            IRowStream row,
            Dictionary<string, List<StudentLookupRecord>> studentDict,
            out SchoolYearEntryRankImportRecord record)
        {
            record = null;

            if (!row.Contains("學號") ||
                !row.Contains("姓名") ||
                !row.Contains("學年度") ||
                !row.Contains("年級") ||
                !row.Contains("學年學業排名"))
            {
                return false;
            }

            string studentNumber = (row.GetValue("學號") ?? string.Empty).Trim();
            string rankName = row.Contains("班級")
                ? (row.GetValue("班級") ?? string.Empty).Trim()
                : string.Empty;
            string excelName = (row.GetValue("姓名") ?? string.Empty).Trim();
            string schoolYearText = (row.GetValue("學年度") ?? string.Empty).Trim();
            string gradeYearText = (row.GetValue("年級") ?? string.Empty).Trim();
            string rankText = (row.GetValue("學年學業排名") ?? string.Empty).Trim();

            if (string.IsNullOrEmpty(studentNumber) ||
                string.IsNullOrEmpty(excelName) ||
                string.IsNullOrEmpty(schoolYearText) ||
                string.IsNullOrEmpty(gradeYearText) ||
                string.IsNullOrEmpty(rankText))
            {
                return false;
            }

            int schoolYear;
            int gradeYear;
            int rank;
            if (!int.TryParse(schoolYearText, out schoolYear) ||
                !int.TryParse(gradeYearText, out gradeYear) ||
                !int.TryParse(rankText, out rank))
            {
                return false;
            }

            if (schoolYear < 1 || schoolYear > 3000 ||
                gradeYear < 1 || gradeYear > 12 ||
                rank < 1)
            {
                return false;
            }

            List<StudentLookupRecord> matches;
            if (!studentDict.TryGetValue(studentNumber, out matches) ||
                matches == null ||
                matches.Count != 1)
            {
                return false;
            }

            string systemName = (matches[0].StudentName ?? string.Empty).Trim();
            if (!string.Equals(systemName, excelName, StringComparison.Ordinal))
                return false;

            record = new SchoolYearEntryRankImportRecord
            {
                RefStudentID = matches[0].StudentID,
                SchoolYear = schoolYear,
                GradeYear = gradeYear,
                RankName = rankName,
                Rank = rank,
                StudentNumber = studentNumber
            };

            return true;
        }

        private static string BuildSummary(int insertedCount, int updatedCount)
        {
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("匯入學年分項排名完成。");
            sb.AppendLine("新增筆數：" + insertedCount);
            sb.AppendLine("更新筆數：" + updatedCount);
            return sb.ToString();
        }
    }
}
