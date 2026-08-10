using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using FISCA.Data;
using Newtonsoft.Json;
using SHImportExportScoreRank.Constants;

namespace SHImportExportScoreRank.DAO
{
    public class RankOverrideDataAccess
    {
        public Dictionary<string, List<StudentLookupRecord>> GetStudentsByStudentNumbers(
            IEnumerable<string> studentNumbers)
        {
            Dictionary<string, List<StudentLookupRecord>> result =
                new Dictionary<string, List<StudentLookupRecord>>(StringComparer.Ordinal);

            List<string> numbers = studentNumbers
                .Where(x => !string.IsNullOrWhiteSpace(x))
                .Select(x => x.Trim())
                .Distinct(StringComparer.Ordinal)
                .ToList();

            if (numbers.Count == 0)
                return result;

            StringBuilder inClause = new StringBuilder();
            for (int i = 0; i < numbers.Count; i++)
            {
                if (i > 0)
                    inClause.Append(",");

                inClause.Append("'");
                inClause.Append(SqlLiteralHelper.Escape(numbers[i]));
                inClause.Append("'");
            }

            string sql = string.Format(
                "SELECT id, student_number, name FROM student WHERE student_number IN ({0});",
                inClause);

            QueryHelper queryHelper = new QueryHelper();
            DataTable table = queryHelper.Select(sql);

            foreach (DataRow row in table.Rows)
            {
                StudentLookupRecord record = new StudentLookupRecord
                {
                    StudentID = Convert.ToInt64(row["id"]),
                    StudentNumber = Convert.ToString(row["student_number"]) ?? string.Empty,
                    StudentName = Convert.ToString(row["name"]) ?? string.Empty
                };

                string key = record.StudentNumber.Trim();
                if (!result.ContainsKey(key))
                    result[key] = new List<StudentLookupRecord>();

                result[key].Add(record);
            }

            return result;
        }

        public List<string> FindDuplicateNaturalKeys(
            List<SchoolYearEntryRankImportRecord> records)
        {
            List<string> messages = new List<string>();

            if (records == null || records.Count == 0)
                return messages;

            List<long> studentIds = records
                .Select(x => x.RefStudentID)
                .Distinct()
                .ToList();

            if (studentIds.Count == 0)
                return messages;

            string studentIdClause = string.Join(",", studentIds);
            string sql = string.Format(@"
SELECT
    ro.ref_student_id,
    s.student_number,
    ro.school_year,
    ro.grade_year,
    COUNT(*) AS cnt
FROM rank_override ro
LEFT JOIN student s ON s.id = ro.ref_student_id
WHERE ro.semester = {0}
  AND ro.item_type = '{1}'
  AND ro.rank_type = '{2}'
  AND ro.ref_student_id IN ({3})
GROUP BY
    ro.ref_student_id,
    s.student_number,
    ro.school_year,
    ro.grade_year
HAVING COUNT(*) > 1;",
                RankOverrideConstants.SchoolYearSemester,
                SqlLiteralHelper.Escape(RankOverrideConstants.ItemType),
                SqlLiteralHelper.Escape(RankOverrideConstants.RankType),
                studentIdClause);

            QueryHelper queryHelper = new QueryHelper();
            DataTable table = queryHelper.Select(sql);

            HashSet<string> importKeys = new HashSet<string>(
                records.Select(x => BuildNaturalKey(
                    x.RefStudentID, x.SchoolYear, x.GradeYear)));

            foreach (DataRow row in table.Rows)
            {
                if (row["grade_year"] == null || row["grade_year"] == DBNull.Value)
                    continue;

                long studentId = Convert.ToInt64(row["ref_student_id"]);
                int schoolYear = Convert.ToInt32(row["school_year"]);
                int gradeYear = Convert.ToInt32(row["grade_year"]);
                string key = BuildNaturalKey(studentId, schoolYear, gradeYear);

                if (!importKeys.Contains(key))
                    continue;

                string studentNumber = Convert.ToString(row["student_number"]) ?? string.Empty;
                messages.Add(string.Format(
                    "學生系統編號：{0}，學號：{1}，學年度：{2}，年級：{3}（資料庫已有 {4} 筆重複自然鍵）",
                    studentId,
                    studentNumber,
                    schoolYear,
                    gradeYear,
                    row["cnt"]));
            }

            return messages;
        }

        private static string BuildNaturalKey(
            long studentID,
            int schoolYear,
            int gradeYear)
        {
            return studentID + "_" + schoolYear + "_" + gradeYear;
        }

        public List<SchoolYearEntryRankRecord> GetSchoolYearEntryRanksByStudentID(
            string studentID)
        {
            List<SchoolYearEntryRankRecord> result =
                new List<SchoolYearEntryRankRecord>();

            long id;
            if (string.IsNullOrWhiteSpace(studentID) ||
                !long.TryParse(studentID.Trim(), out id) ||
                id <= 0)
            {
                return result;
            }

            string sql = string.Format(@"
SELECT
    ro.id,
    ro.ref_student_id,
    ro.school_year,
    ro.grade_year,
    ro.rank_name,
    ro.rank
FROM rank_override ro
WHERE ro.ref_student_id = {0}
  AND ro.semester = {1}
  AND ro.item_type = '{2}'
  AND ro.item_name = '{3}'
  AND ro.rank_type = '{4}'
  AND ro.ref_exam_id IS NULL
ORDER BY
    ro.school_year DESC,
    ro.grade_year DESC,
    ro.id DESC;",
                id,
                RankOverrideConstants.SchoolYearSemester,
                SqlLiteralHelper.Escape(RankOverrideConstants.ItemType),
                SqlLiteralHelper.Escape(RankOverrideConstants.ItemName),
                SqlLiteralHelper.Escape(RankOverrideConstants.RankType));

            QueryHelper queryHelper = new QueryHelper();
            DataTable table = queryHelper.Select(sql);

            foreach (DataRow row in table.Rows)
                result.Add(MapRankRecord(row, false));

            return result;
        }

        public List<SchoolYearEntryRankRecord> GetSchoolYearEntryRanksByStudentIDs(
            IEnumerable<string> studentIDs)
        {
            List<SchoolYearEntryRankRecord> result =
                new List<SchoolYearEntryRankRecord>();

            List<long> ids = new List<long>();
            if (studentIDs != null)
            {
                foreach (string studentID in studentIDs)
                {
                    long id;
                    if (!string.IsNullOrWhiteSpace(studentID) &&
                        long.TryParse(studentID.Trim(), out id) &&
                        id > 0 &&
                        !ids.Contains(id))
                    {
                        ids.Add(id);
                    }
                }
            }

            if (ids.Count == 0)
                return result;

            string studentIdClause = string.Join(",", ids);
            string sql = string.Format(@"
SELECT
    ro.id,
    ro.ref_student_id,
    ro.school_year,
    ro.grade_year,
    ro.rank_name,
    ro.rank,
    s.student_number,
    s.seat_no,
    s.name AS student_name,
    CASE
        WHEN s.ref_dept_id IS NOT NULL THEN student_dept.name
        ELSE class_dept.name
    END AS department_name
FROM rank_override ro
LEFT JOIN student s ON s.id = ro.ref_student_id
LEFT JOIN class ON class.id = s.ref_class_id
LEFT JOIN dept class_dept ON class_dept.id = class.ref_dept_id
LEFT JOIN dept student_dept ON student_dept.id = s.ref_dept_id
WHERE ro.ref_student_id IN ({0})
ORDER BY
    ro.ref_student_id,
    ro.school_year,
    ro.grade_year,
    ro.id;",
                studentIdClause);

            QueryHelper queryHelper = new QueryHelper();
            DataTable table = queryHelper.Select(sql);

            foreach (DataRow row in table.Rows)
                result.Add(MapRankRecord(row, true));

            return result;
        }

        private static SchoolYearEntryRankRecord MapRankRecord(
            DataRow row,
            bool includeStudentFields)
        {
            SchoolYearEntryRankRecord record = new SchoolYearEntryRankRecord
            {
                ID = Convert.ToInt64(row["id"]),
                StudentID = Convert.ToString(row["ref_student_id"]) ?? string.Empty,
                SchoolYear = ToNullableInt(row["school_year"]),
                GradeYear = ToNullableInt(row["grade_year"]),
                RankName = NullToEmpty(row["rank_name"]),
                Rank = ToNullableInt(row["rank"])
            };

            if (includeStudentFields)
            {
                record.StudentNumber = NullToEmpty(row["student_number"]);
                record.SeatNo = NullToEmpty(row["seat_no"]);
                record.StudentName = NullToEmpty(row["student_name"]);
                record.DepartmentName = NullToEmpty(row["department_name"]);
            }

            return record;
        }

        private static int? ToNullableInt(object value)
        {
            if (value == null || value == DBNull.Value)
                return null;

            return Convert.ToInt32(value);
        }

        private static string NullToEmpty(object value)
        {
            if (value == null || value == DBNull.Value)
                return string.Empty;

            return Convert.ToString(value) ?? string.Empty;
        }

        public UpsertResult UpsertSchoolYearEntryRanks(
            List<SchoolYearEntryRankImportRecord> records)
        {
            UpsertResult result = new UpsertResult();

            if (records == null || records.Count == 0)
                return result;

            List<object> jsonRows = records.Select(x => new
            {
                ref_student_id = x.RefStudentID,
                school_year = x.SchoolYear,
                grade_year = x.GradeYear,
                rank_name = x.RankName,
                rank = x.Rank
            }).Cast<object>().ToList();

            string json = JsonConvert.SerializeObject(jsonRows);
            string escapedJson = SqlLiteralHelper.Escape(json);

            string sql = string.Format(@"
WITH row AS (
    SELECT '{0}'::JSONB AS input
), input_rows AS (
    SELECT
        (obj->>'ref_student_id')::BIGINT AS ref_student_id,
        (obj->>'school_year')::INT AS school_year,
        (obj->>'grade_year')::INT AS grade_year,
        obj->>'rank_name' AS rank_name,
        (obj->>'rank')::INT AS rank
    FROM jsonb_array_elements((SELECT input FROM row)) AS obj
), matched AS (
    SELECT
        ir.ref_student_id,
        ir.school_year,
        ir.grade_year,
        ir.rank_name,
        ir.rank,
        ro.id AS current_id
    FROM input_rows ir
    LEFT JOIN rank_override ro
        ON ro.ref_student_id = ir.ref_student_id
        AND ro.school_year = ir.school_year
        AND ro.semester = {1}
        AND ro.grade_year = ir.grade_year
        AND ro.item_type = '{2}'
        AND ro.rank_type = '{3}'
), data_update AS (
    UPDATE rank_override
    SET rank = matched.rank,
        extension = jsonb_build_array(
            jsonb_build_object(
                'extension_name', '排名資料',
                'create_time', now(),
                '建立方式', '匯入',
                '成績類型', '學年',
                '成績項目', '分項'
            )
        )
    FROM matched
    WHERE rank_override.id = matched.current_id
      AND matched.current_id IS NOT NULL
    RETURNING rank_override.id, 'updated'::TEXT AS action
), data_insert AS (
    INSERT INTO rank_override (
        ref_student_id,
        school_year,
        semester,
        grade_year,
        item_type,
        ref_exam_id,
        item_name,
        rank_type,
        rank_name,
        rank,
        extension
    )
    SELECT
        matched.ref_student_id,
        matched.school_year,
        {1},
        matched.grade_year,
        '{2}',
        -1,
        '{4}',
        '{3}',
        matched.rank_name,
        matched.rank,
        jsonb_build_array(
            jsonb_build_object(
                'extension_name', '排名資料',
                'create_time', now(),
                '建立方式', '匯入',
                '成績類型', '學年',
                '成績項目', '分項'
            )
        )
    FROM matched
    WHERE matched.current_id IS NULL
    RETURNING id, 'inserted'::TEXT AS action
)
SELECT
    COUNT(CASE WHEN action = 'inserted' THEN 1 END) AS inserted_count,
    COUNT(CASE WHEN action = 'updated' THEN 1 END) AS updated_count
FROM (
    SELECT action FROM data_update
    UNION ALL
    SELECT action FROM data_insert
) AS upsert_result;",
                escapedJson,
                RankOverrideConstants.SchoolYearSemester,
                SqlLiteralHelper.Escape(RankOverrideConstants.ItemType),
                SqlLiteralHelper.Escape(RankOverrideConstants.RankType),
                SqlLiteralHelper.Escape(RankOverrideConstants.ItemName));

            QueryHelper queryHelper = new QueryHelper();
            DataTable table = queryHelper.Select(sql);

            if (table.Rows.Count > 0)
            {
                result.InsertedCount = Convert.ToInt32(table.Rows[0]["inserted_count"]);
                result.UpdatedCount = Convert.ToInt32(table.Rows[0]["updated_count"]);
            }

            return result;
        }
    }

    public class UpsertResult
    {
        public int InsertedCount { get; set; }
        public int UpdatedCount { get; set; }
    }
}
