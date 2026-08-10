using System;
using System.Collections.Generic;
using Campus.DocumentValidator;
using FISCA.Data;
using SHImportExportScoreRank.DAO;

namespace SHImportExportScoreRank.ValidationRule.RowValidator
{
    /// <summary>
    /// 檢查學號是否存在、是否只對應一位學生，以及姓名是否與系統一致。
    /// </summary>
    public class CheckSchoolYearEntryRankStudentValidator : IRowVaildator
    {
        private readonly Dictionary<string, List<StudentLookupRecord>> _studentsByNumber;

        public CheckSchoolYearEntryRankStudentValidator()
        {
            _studentsByNumber = LoadAllStudentsByNumber();
        }

        public string Correct(IRowStream Value)
        {
            return string.Empty;
        }

        public string ToString(string template)
        {
            return template;
        }

        public bool Validate(IRowStream Value)
        {
            if (!Value.Contains("學號") || !Value.Contains("姓名"))
                return false;

            string studentNumber = (Value.GetValue("學號") ?? string.Empty).Trim();
            string excelName = (Value.GetValue("姓名") ?? string.Empty).Trim();

            if (string.IsNullOrEmpty(studentNumber) || string.IsNullOrEmpty(excelName))
                return false;

            List<StudentLookupRecord> matches;
            if (!_studentsByNumber.TryGetValue(studentNumber, out matches) || matches == null)
                return false;

            if (matches.Count != 1)
                return false;

            string systemName = (matches[0].StudentName ?? string.Empty).Trim();
            return string.Equals(systemName, excelName, StringComparison.Ordinal);
        }

        private static Dictionary<string, List<StudentLookupRecord>> LoadAllStudentsByNumber()
        {
            Dictionary<string, List<StudentLookupRecord>> result =
                new Dictionary<string, List<StudentLookupRecord>>(StringComparer.Ordinal);

            QueryHelper queryHelper = new QueryHelper();
            System.Data.DataTable table = queryHelper.Select(
                "SELECT id, student_number, name FROM student WHERE student_number IS NOT NULL AND student_number <> '';");

            foreach (System.Data.DataRow row in table.Rows)
            {
                string studentNumber = (Convert.ToString(row["student_number"]) ?? string.Empty).Trim();
                if (string.IsNullOrEmpty(studentNumber))
                    continue;

                StudentLookupRecord record = new StudentLookupRecord
                {
                    StudentID = Convert.ToInt64(row["id"]),
                    StudentNumber = studentNumber,
                    StudentName = Convert.ToString(row["name"]) ?? string.Empty
                };

                if (!result.ContainsKey(studentNumber))
                    result[studentNumber] = new List<StudentLookupRecord>();

                result[studentNumber].Add(record);
            }

            return result;
        }
    }
}
