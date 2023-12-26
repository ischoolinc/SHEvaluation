using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data;
using FISCA.Data;
using System.Windows.Forms;

namespace RegularAssessmentTranscriptFixedRank.DAO
{
    // 學生人數統計
    public class StudentCountInfo
    {
        List<string> _StudentSourceIDs; // 統計來源學生
        Dictionary<string, int> _ClassStudentCountDict; // 班級人數統計結果
        Dictionary<string, int> _DeptStudentCountDict; // 科別人數統計結果
        Dictionary<string, int> _GradeStudentCountDict; // 年級人數統計結果
        Dictionary<string, int> _TagStudentCountDict; // 類別人數統計結果
        Dictionary<string, int> _FullNameTagStudentCountDict; // 類別完整名稱人數統計結果
        Dictionary<string, List<string>> _StudentTagPrefixList; // 來源學生前置類別清單
        Dictionary<string, List<string>> _StudentFullNameTagList; // 來源學生完整類別清單

        public StudentCountInfo()
        {
            _StudentSourceIDs = new List<string>();
            _ClassStudentCountDict = new Dictionary<string, int>();
            _DeptStudentCountDict = new Dictionary<string, int>();
            _GradeStudentCountDict = new Dictionary<string, int>();
            _TagStudentCountDict = new Dictionary<string, int>();
            _StudentTagPrefixList = new Dictionary<string, List<string>>();
            _StudentFullNameTagList = new Dictionary<string, List<string>>();
            _FullNameTagStudentCountDict = new Dictionary<string, int>();
        }

        // 設定統計來源學生
        public void SetStudentSourceIDs(List<string> StudentIDs)

        {
            _StudentSourceIDs = StudentIDs;
        }

        // 計算學生人數
        public void CalcStudents()
        {

            QueryHelper qh = new QueryHelper();
            string strSQL = "";
            // 有人再計算
            if (_StudentSourceIDs.Count > 0)
            {
                // 計算班級人數
                _ClassStudentCountDict.Clear();
                try
                {
                    strSQL = string.Format(@"
                    SELECT
                        class_name,
                        count(student.id) AS 人數
                    FROM
                        student
                        INNER JOIN class ON student.ref_class_id = class.id
                    WHERE
                        student.status IN (1, 2)
                        AND class.id IN(
                            SELECT
                                DISTINCT ref_class_id
                            FROM
                                student
                            WHERE
                                student.id IN({0})
                        )
                    GROUP BY
                        class_name
                ", string.Join(",", _StudentSourceIDs.ToArray()));

                    DataTable dtClass = new DataTable();
                    dtClass = qh.Select(strSQL);
                    foreach (DataRow dr in dtClass.Rows)
                    {
                        int count = 0;
                        string class_name = dr["class_name"] + "";
                        if (int.TryParse(dr["人數"] + "", out count))
                        {
                            if (!_ClassStudentCountDict.ContainsKey(class_name))
                                _ClassStudentCountDict.Add(class_name, count);
                        }
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("計算班級人數:" + ex.Message);
                }

                // 計算年級科別人數
                _DeptStudentCountDict.Clear();
                try
                {
                    strSQL = string.Format(@"                  
                    WITH dept_data AS(
                        SELECT
                            id AS dept_id,
                            name AS dept_name
                        FROM
                            dept
                    ),
                    student_count AS(
                        SELECT
                            class.grade_year AS 年級,
                            COALESCE(student.ref_dept_id, class.ref_dept_id) AS dept_id,
                            count(student.id) AS 人數
                        FROM
                            student
                            INNER JOIN class ON student.ref_class_id = class.id
                        WHERE
                            student.status IN (1, 2)
                            AND class.grade_year IN(
                                SELECT
                                    DISTINCT class.grade_year
                                FROM
                                    student
                                    INNER JOIN class ON student.ref_class_id = class.id
                                WHERE
                                    student.id IN({0})
                            )
                        GROUP BY
                            class.grade_year,
                            dept_id
                    )
                    SELECT
                        student_count.年級 AS 年級,
                        dept_data.dept_name AS 科別名稱,
                        student_count.人數 AS 人數
                    FROM
                        dept_data
                        INNER JOIN student_count ON student_count.dept_id = dept_data.dept_id
                    ORDER BY
                        年級,
                        科別名稱
                    ", string.Join(",", _StudentSourceIDs.ToArray()));

                    DataTable dtDept = new DataTable();
                    dtDept = qh.Select(strSQL);
                    foreach (DataRow dr in dtDept.Rows)
                    {
                        int count = 0;
                        string grade_year = dr["年級"] + "";
                        string dept_name = dr["科別名稱"] + "";
                        if (int.TryParse(dr["人數"] + "", out count))
                        {
                            string key = grade_year + "_" + dept_name;
                            if (!_DeptStudentCountDict.ContainsKey(key))
                                _DeptStudentCountDict.Add(key, count);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine("計算年級科別人數:" + ex.Message);
                }

                // 計算年級人數
                _GradeStudentCountDict.Clear();
                try
                {
                    strSQL = string.Format(@"
                    SELECT
                        class.grade_year AS grade_year,
                        count(student.id) AS 人數
                    FROM
                        student
                        INNER JOIN class ON student.ref_class_id = class.id
                    WHERE
                        student.status IN (1, 2)
                        AND class.grade_year IN(
                            SELECT
                                DISTINCT class.grade_year
                            FROM
                                student
                                INNER JOIN class ON student.ref_class_id = class.id
                            WHERE
                                student.id IN({0})
                        )
                    GROUP BY
                        class.grade_year
                    ORDER BY
                        class.grade_year
                    ", string.Join(",", _StudentSourceIDs.ToArray()));
                }
                catch (Exception ex)
                {
                    Console.WriteLine("計算年級人數:" + ex.Message);
                }

                DataTable dtGradeYear = new DataTable();
                dtGradeYear = qh.Select(strSQL);
                foreach (DataRow dr in dtGradeYear.Rows)
                {
                    int count = 0;
                    string grade_year = dr["grade_year"] + "";
                    if (int.TryParse(dr["人數"] + "", out count))
                    {
                        if (!_GradeStudentCountDict.ContainsKey(grade_year))
                            _GradeStudentCountDict.Add(grade_year, count);
                    }
                }

                // 計算年級類別人數
                _TagStudentCountDict.Clear();
                try
                {
                    strSQL = string.Format(@"
                    SELECT
                        class.grade_year AS grade_year,
                        tag.prefix AS tag_prefix,
                       -- tag.prefix || ':' || tag.name AS tag_full_name,
                        count(student.id) AS 人數
                    FROM
                        student
                        INNER JOIN class ON student.ref_class_id = class.id
                        INNER JOIN tag_student ON tag_student.ref_student_id = student.id
                        INNER JOIN tag ON tag.id = tag_student.ref_tag_id
                    WHERE
                        student.status IN (1, 2)
                        AND tag_student.ref_tag_id IN(
                            SELECT
                                DISTINCT ref_tag_id
                            FROM
                                student
                                INNER JOIN tag_student ON tag_student.ref_student_id = student.id
                            WHERE
                                student.id IN({0})
                        )
                    GROUP BY
                        class.grade_year,
                        tag_prefix    
                    ORDER BY
                        class.grade_year,
                        tag_prefix
                    ", string.Join(",", _StudentSourceIDs.ToArray()));

                    DataTable dtTag = new DataTable();
                    dtTag = qh.Select(strSQL);
                    foreach (DataRow dr in dtTag.Rows)
                    {
                        int count = 0;
                        string grade_year = dr["grade_year"] + "";
                        string tag_prefix = dr["tag_prefix"] + "";
                        string key = grade_year + "_" + tag_prefix;
                        if (int.TryParse(dr["人數"] + "", out count))
                        {
                            if (!_TagStudentCountDict.ContainsKey(key))
                                _TagStudentCountDict.Add(key, count);
                        }
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("計算年級類別人數:" + ex.Message);
                }

                _StudentTagPrefixList.Clear();
                // 取得來源學生類別前置
                try
                {
                    strSQL = string.Format(@"
                    SELECT
                        ref_student_id AS student_id,
                        tag.prefix AS tag_prefix
                    FROM
                        tag_student 
                        INNER JOIN tag ON tag.id = tag_student.ref_tag_id
                    WHERE
                        tag_student.ref_student_id IN({0})
                    ", string.Join(",", _StudentSourceIDs.ToArray()));


                    DataTable dtTag = new DataTable();
                    dtTag = qh.Select(strSQL);
                    foreach (DataRow dr in dtTag.Rows)
                    {
                        string sid = dr["student_id"] + "";
                        string prefix = dr["tag_prefix"] + "";
                        if (!_StudentTagPrefixList.ContainsKey(sid))
                            _StudentTagPrefixList.Add(sid, new List<string>());

                        if (!_StudentTagPrefixList[sid].Contains(prefix))
                            _StudentTagPrefixList[sid].Add(prefix);
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("取得來源學生類別前置:" + ex.Message);
                }


                // 計算年級完整類別人數
                _FullNameTagStudentCountDict.Clear();
                try
                {
                    strSQL = string.Format(@"
                    SELECT
                        class.grade_year AS grade_year,
                        -- tag.prefix AS tag_prefix,
                        tag.prefix || ':' || tag.name AS tag_full_name,
                        count(student.id) AS 人數
                    FROM
                        student
                        INNER JOIN class ON student.ref_class_id = class.id
                        INNER JOIN tag_student ON tag_student.ref_student_id = student.id
                        INNER JOIN tag ON tag.id = tag_student.ref_tag_id
                    WHERE
                        student.status IN (1, 2)
                        AND tag_student.ref_tag_id IN(
                            SELECT
                                DISTINCT ref_tag_id
                            FROM
                                student
                                INNER JOIN tag_student ON tag_student.ref_student_id = student.id
                            WHERE
                                student.id IN({0})
                        )
                    GROUP BY
                        class.grade_year,
                        tag_full_name    
                    ORDER BY
                        class.grade_year,
                        tag_full_name
                    ", string.Join(",", _StudentSourceIDs.ToArray()));

                    DataTable dtTag = new DataTable();
                    dtTag = qh.Select(strSQL);
                    foreach (DataRow dr in dtTag.Rows)
                    {
                        int count = 0;
                        string grade_year = dr["grade_year"] + "";
                        string tag_fullname = dr["tag_full_name"] + "";
                        string key = grade_year + "_" + tag_fullname;
                        if (int.TryParse(dr["人數"] + "", out count))
                        {
                            if (!_FullNameTagStudentCountDict.ContainsKey(key))
                                _FullNameTagStudentCountDict.Add(key, count);
                        }
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("計算年級類別完整名稱人數:" + ex.Message);
                }

                _StudentFullNameTagList.Clear();
                // 取得來源學生類別完整名稱
                try
                {
                    strSQL = string.Format(@"
                    SELECT
                        ref_student_id AS student_id,
                         tag.prefix || ':' || tag.name AS tag_full_name
                    FROM
                        tag_student 
                        INNER JOIN tag ON tag.id = tag_student.ref_tag_id
                    WHERE
                        tag_student.ref_student_id IN({0})
                    ", string.Join(",", _StudentSourceIDs.ToArray()));


                    DataTable dtTag = new DataTable();
                    dtTag = qh.Select(strSQL);
                    foreach (DataRow dr in dtTag.Rows)
                    {
                        string sid = dr["student_id"] + "";
                        string fullname = dr["tag_full_name"] + "";
                        if (!_StudentFullNameTagList.ContainsKey(sid))
                            _StudentFullNameTagList.Add(sid, new List<string>());

                        if (!_StudentFullNameTagList[sid].Contains(fullname))
                            _StudentFullNameTagList[sid].Add(fullname);
                    }

                }
                catch (Exception ex)
                {
                    Console.WriteLine("取得來源學生類別完整名稱:" + ex.Message);
                }

            }
        }

        // 透過班級名稱取得班級人數
        public int GetClassStudentCount(string ClassName)
        {
            int count = 0;
            if (_ClassStudentCountDict.ContainsKey(ClassName))
                count = _ClassStudentCountDict[ClassName];
            return count;
        }

        // 透過年級科別名稱，取得科別人數
        public int GetDeptStudentCount(string grade_year, string DeptName)
        {
            int count = 0;
            string key = grade_year + "_" + DeptName;
            if (_DeptStudentCountDict.ContainsKey(key))
                count = _DeptStudentCountDict[key];
            return count;
        }

        // 取得年級人數
        public int GetGradeStudentCount(string grade_year)
        {
            int count = 0;
            if (_GradeStudentCountDict.ContainsKey(grade_year))
                count = _GradeStudentCountDict[grade_year];
            return count;
        }

        // 透過年級與類別前置，取得年級類別人數
        public int GetTagStudentCount(string grade_year, string tag_prefix)
        {
            int count = 0;
            string key = grade_year + "_" + tag_prefix;
            if (_TagStudentCountDict.ContainsKey(key))
                count = _TagStudentCountDict[key];
            return count;

        }

        // 透過年級與類別完整名稱，取得年級類別人數
        public int GetFullNameTagStudentCount(string grade_year, string fullname)
        {
            int count = 0;
            string key = grade_year + "_" + fullname;
            if (_FullNameTagStudentCountDict.ContainsKey(key))
                count = _FullNameTagStudentCountDict[key];
            return count;

        }

        // 取得學生前置
        public List<string> GetStudentPrefixList(string StudentID)
        {
            if (_StudentTagPrefixList.ContainsKey(StudentID))
                return _StudentTagPrefixList[StudentID];
            else
                return new List<string>();
        }

        // 取得前置內容
        public List<string> GetPrefixList()
        {
            List<string> value = new List<string>();
            foreach (string id in _StudentTagPrefixList.Keys)
            {
                foreach (string name in _StudentTagPrefixList[id])
                {
                    if (!value.Contains(name))
                        value.Add(name);
                }
            }

            return value;
        }

        // 取得學生完整類別名稱
        public List<string> GetStudentFullNameTagList(string StudentID)
        {
            if (_StudentFullNameTagList.ContainsKey(StudentID))
                return _StudentFullNameTagList[StudentID];
            else
                return new List<string>();
        }

        // 取得完整類別內容
        public List<string> GetFullNameTagList()
        {
            List<string> value = new List<string>();
            foreach (string id in _StudentFullNameTagList.Keys)
            {
                foreach (string name in _StudentFullNameTagList[id])
                {
                    if (!value.Contains(name))
                        value.Add(name);
                }
            }

            return value;
        }

    }
}
