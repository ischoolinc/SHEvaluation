using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using K12.Data;
using FISCA.UDT;
using FISCA.Data;
using System.Data;
using System.Windows.Forms;
using System.Xml.Linq;

namespace RegularAssessmentTranscriptFixedRank
{
    public class Utility
    {

        /// <summary>
        /// 透過日期區間取得學生缺曠統計(傳入學生系統編號、開始日期、結束日期；回傳：學生系統編號、獎懲名稱,統計值
        /// </summary>
        /// <param name="StudIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, int>> GetAttendanceCountByDate(List<StudentRecord> StudRecordList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, Dictionary<string, int>> retVal = new Dictionary<string, Dictionary<string, int>>();

            List<PeriodMappingInfo> PeriodMappingList = PeriodMapping.SelectAll();
            // 節次>類別
            Dictionary<string, string> PeriodMappingDict = new Dictionary<string, string>();
            foreach (PeriodMappingInfo rec in PeriodMappingList)
            {
                if (!PeriodMappingDict.ContainsKey(rec.Name))
                    PeriodMappingDict.Add(rec.Name, rec.Type);
            }

            List<AttendanceRecord> attendList = K12.Data.Attendance.SelectByDate(StudRecordList, beginDate, endDate);

            // 計算統計資料
            foreach (AttendanceRecord rec in attendList)
            {
                if (!retVal.ContainsKey(rec.RefStudentID))
                    retVal.Add(rec.RefStudentID, new Dictionary<string, int>());

                foreach (AttendancePeriod per in rec.PeriodDetail)
                {
                    if (!PeriodMappingDict.ContainsKey(per.Period))
                        continue;

                    // ex.一般:曠課
                    string key = "區間" + PeriodMappingDict[per.Period] + "_" + per.AbsenceType;

                    if (!retVal[rec.RefStudentID].ContainsKey(key))
                        retVal[rec.RefStudentID].Add(key, 0);

                    retVal[rec.RefStudentID][key]++;
                }
            }

            return retVal;
        }

        /// <summary>
        /// 透過日期區間取得獎懲資料,傳入學生ID,開始日期,結束日期,回傳：學生ID,獎懲統計名稱,統計值
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, int>> GetDisciplineCountByDate(List<string> StudentIDList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, Dictionary<string, int>> retVal = new Dictionary<string, Dictionary<string, int>>();

            List<string> nameList = new string[] { "大功", "小功", "嘉獎", "大過", "小過", "警告", "留校" }.ToList();

            // 取得獎懲資料
            List<DisciplineRecord> dataList = Discipline.SelectByStudentIDs(StudentIDList);

            foreach (DisciplineRecord data in dataList)
            {
                if (data.OccurDate >= beginDate && data.OccurDate <= endDate)
                {
                    // 初始化
                    if (!retVal.ContainsKey(data.RefStudentID))
                    {
                        retVal.Add(data.RefStudentID, new Dictionary<string, int>());
                        foreach (string str in nameList)
                            retVal[data.RefStudentID].Add(str, 0);
                    }

                    // 獎勵
                    if (data.MeritFlag == "1")
                    {
                        if (data.MeritA.HasValue)
                            retVal[data.RefStudentID]["大功"] += data.MeritA.Value;

                        if (data.MeritB.HasValue)
                            retVal[data.RefStudentID]["小功"] += data.MeritB.Value;

                        if (data.MeritC.HasValue)
                            retVal[data.RefStudentID]["嘉獎"] += data.MeritC.Value;
                    }
                    else if (data.MeritFlag == "0")
                    { // 懲戒
                        if (data.Cleared != "是")
                        {
                            if (data.DemeritA.HasValue)
                                retVal[data.RefStudentID]["大過"] += data.DemeritA.Value;

                            if (data.DemeritB.HasValue)
                                retVal[data.RefStudentID]["小過"] += data.DemeritB.Value;

                            if (data.DemeritC.HasValue)
                                retVal[data.RefStudentID]["警告"] += data.DemeritC.Value;
                        }
                    }
                    else if (data.MeritFlag == "2")
                    {
                        // 留校察看
                        retVal[data.RefStudentID]["留校"]++;
                    }
                }
            }
            return retVal;
        }

        /// <summary>
        /// 透過學生編號取得學生服務學習時數 傳入學生編號、開始日期、結束日期,回傳：學生編號、內容、值
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, decimal>> GetServiceLearningByDate(List<string> StudentIDList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, Dictionary<string, decimal>> retVal = new Dictionary<string, Dictionary<string, decimal>>();

            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,school_year,semester,hours from $k12.service.learning.record where ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "') and occur_date >='" + beginDate.ToShortDateString() + "' and occur_date <='" + endDate.ToShortDateString() + "'order by ref_student_id,school_year,semester;";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr[0].ToString();
                    string key2 = dr[1].ToString() + "學年度第" + dr[2].ToString() + "學期";
                    decimal hr;
                    decimal.TryParse(dr[3].ToString(), out hr);

                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, new Dictionary<string, decimal>());

                    if (!retVal[sid].ContainsKey(key2))
                        retVal[sid].Add(key2, 0);

                    retVal[sid][key2] += hr;

                }
            }
            return retVal;
        }

        /// <summary>
        /// 透過學生編號、開始與結束日期，取的獎懲資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, List<DisciplineRecord>> GetDisciplineDetailByDate(List<string> StudentIDList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, List<DisciplineRecord>> retVal = new Dictionary<string, List<DisciplineRecord>>();

            // 取得獎懲資料
            List<DisciplineRecord> dataList = Discipline.SelectByStudentIDs(StudentIDList);
            // 依日期排序
            dataList = (from data in dataList orderby data.OccurDate select data).ToList();

            foreach (DisciplineRecord rec in dataList)
            {
                if (rec.OccurDate >= beginDate && rec.OccurDate <= endDate)
                {
                    if (!retVal.ContainsKey(rec.RefStudentID))
                        retVal.Add(rec.RefStudentID, new List<DisciplineRecord>());

                    retVal[rec.RefStudentID].Add(rec);
                }
            }

            return retVal;
        }

        /// <summary>
        /// 透過學生編號、開始與結束日期，取得缺曠資料
        /// </summary>
        /// <param name="StudRecordList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, List<AttendanceRecord>> GetAttendanceDetailByDate(List<StudentRecord> StudRecordList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, List<AttendanceRecord>> retVal = new Dictionary<string, List<AttendanceRecord>>();

            // 讀取資料
            List<AttendanceRecord> attendList = K12.Data.Attendance.SelectByDate(StudRecordList, beginDate, endDate);

            // 依日期排序
            attendList = (from data in attendList orderby data.OccurDate select data).ToList();

            foreach (AttendanceRecord rec in attendList)
            {
                if (!retVal.ContainsKey(rec.RefStudentID))
                    retVal.Add(rec.RefStudentID, new List<AttendanceRecord>());

                retVal[rec.RefStudentID].Add(rec);
            }

            return retVal;
        }


        /// <summary>
        /// 透過學生編號、開始與結束日期，取得學習服務 DataRow
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, List<DataRow>> GetServiceLearningDetailByDate(List<string> StudentIDList, DateTime beginDate, DateTime endDate)
        {
            Dictionary<string, List<DataRow>> retVal = new Dictionary<string, List<DataRow>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,occur_date,reason,hours from $k12.service.learning.record where ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "') and occur_date >='" + beginDate.ToShortDateString() + "' and occur_date <='" + endDate.ToShortDateString() + "'order by ref_student_id,occur_date;";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr[0].ToString();
                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, new List<DataRow>());

                    retVal[sid].Add(dr);

                }
            }
            return retVal;
        }


        static Dictionary<string, decimal> ServiceLearningByDateDict = new Dictionary<string, decimal>();
        static Dictionary<string, decimal> ServiceLearningBySemesterDict = new Dictionary<string, decimal>();

        /// <summary>
        /// 服務學習時數加總
        /// </summary>
        public static void SetServiceLearningSum(List<string> StudentIDList, DateTime beginDate, DateTime endDate, string SchoolYear, string Semester)
        {
            ServiceLearningByDateDict.Clear();
            ServiceLearningBySemesterDict.Clear();

            // 沒有學生不處理
            if (StudentIDList.Count == 0)
                return;

            // 依日期區間
            QueryHelper qh = new QueryHelper();
            string query1 = "SELECT " +
                "ref_student_id AS student_id" +
                ",hours " +
                "FROM $k12.service.learning.record " +
                "WHERE ref_student_id " +
                "IN('" + string.Join("','", StudentIDList.ToArray()) + "') " +
                "AND occur_date >='" + beginDate.ToShortDateString() + "' " +
                "AND occur_date <='" + endDate.ToShortDateString() + "' " +
                "ORDER BY ref_student_id;";

            DataTable dt1 = qh.Select(query1);
            if (dt1 != null)
            {
                foreach (DataRow dr in dt1.Rows)
                {
                    string sid = dr["student_id"].ToString();
                    if (!ServiceLearningByDateDict.ContainsKey(sid))
                        ServiceLearningByDateDict.Add(sid, 0);

                    decimal d1;
                    if (dr["hours"] != null)
                    {
                        if (decimal.TryParse(dr["hours"].ToString(), out d1))
                        {
                            ServiceLearningByDateDict[sid] += d1;
                        }
                    }

                }
            }

            // 依學年度學期
            string query2 = "SELECT " +
              "ref_student_id AS student_id" +
              ",hours " +
              "FROM $k12.service.learning.record " +
              "WHERE ref_student_id " +
              "IN('" + string.Join("','", StudentIDList.ToArray()) + "') " +
              "AND school_year=" + SchoolYear +
              "AND semester =" + Semester +
              "ORDER BY ref_student_id;";

            DataTable dt2 = qh.Select(query2);
            if (dt2 != null)
            {
                foreach (DataRow dr in dt2.Rows)
                {
                    string sid = dr["student_id"].ToString();
                    if (!ServiceLearningBySemesterDict.ContainsKey(sid))
                        ServiceLearningBySemesterDict.Add(sid, 0);

                    decimal d1;
                    if (dr["hours"] != null)
                    {
                        if (decimal.TryParse(dr["hours"].ToString(), out d1))
                        {
                            ServiceLearningBySemesterDict[sid] += d1;
                        }
                    }
                }
            }
        }

        public static Dictionary<string, decimal> GetServiceLearningByDateDict()
        {
            return Utility.ServiceLearningByDateDict;
        }

        public static Dictionary<string, decimal> GetServiceLearningBySemesterDict()
        {
            return Utility.ServiceLearningBySemesterDict;
        }

        /// <summary>
        /// 取得缺曠對照 List,一般_曠課..
        /// </summary>
        /// <returns></returns>
        public static List<string> GetATMappingKey()
        {
            List<string> retVal = new List<string>();
            List<string> key1List = new List<string>();
            List<string> Key2List = new List<string>();
            foreach (PeriodMappingInfo data in PeriodMapping.SelectAll())
                if (!key1List.Contains(data.Type))
                    key1List.Add(data.Type);

            foreach (AbsenceMappingInfo data in AbsenceMapping.SelectAll())
                if (!Key2List.Contains(data.Name))
                    Key2List.Add(data.Name);

            // 一般_曠課
            foreach (string key1 in key1List)
                foreach (string key2 in Key2List)
                    retVal.Add(key1 + "_" + key2);

            return retVal;
        }

        /// <summary>
        /// 取得學生排名、五標與組距資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, DataRow>> GetRankMatrixData(string SchoolYear, string Semester, string ExamID, List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, DataRow>> value = new Dictionary<string, Dictionary<string, DataRow>>();

            try
            {

                // 沒有學生不處理
                if (StudentIDList.Count == 0)
                    return value;

                QueryHelper qh = new QueryHelper();
                string query = "" +
                   " SELECT " +
    " 	rank_matrix.id AS rank_matrix_id" +
    " 	, rank_matrix.school_year" +
    " 	, rank_matrix.semester" +
    " 	, rank_matrix.grade_year" +
    " 	, rank_matrix.item_type" +
    " 	, rank_matrix.ref_exam_id AS exam_id" +
    " 	, rank_matrix.item_name" +
    " 	, rank_matrix.rank_type" +
    " 	, rank_matrix.rank_name" +
    " 	, class.class_name" +
    " 	, student.seat_no" +
    " 	, student.student_number" +
    " 	, student.name" +
    " 	, rank_detail.ref_student_id AS student_id " +
    " 	, rank_detail.rank" +
    "   , rank_matrix.matrix_count " +
    " 	, rank_detail.pr" +
    " 	, rank_detail.percentile" +
    "   , rank_matrix.avg_top_25" +
    "   , rank_matrix.avg_top_50" +
    "   , rank_matrix.avg" +
    "   , rank_matrix.avg_bottom_50" +
    "   , rank_matrix.avg_bottom_25" +
    "    , rank_matrix.pr_88" +
    "    ,rank_matrix.pr_75" +
    "    ,rank_matrix.pr_50" +
    "    ,rank_matrix.pr_25" +
    "    ,rank_matrix.pr_12" +
    "    ,rank_matrix.std_dev_pop" +
    " 	, rank_matrix.level_gte100" +
    " 	, rank_matrix.level_90" +
    " 	, rank_matrix.level_80" +
    " 	, rank_matrix.level_70" +
    " 	, rank_matrix.level_60" +
    " 	, rank_matrix.level_50" +
    " 	, rank_matrix.level_40" +
    " 	, rank_matrix.level_30" +
    " 	, rank_matrix.level_20" +
    " 	, rank_matrix.level_10" +
    " 	, rank_matrix.level_lt10" +
    "  , level_gte100+level_90 + level_80 + level_70 + level_60 AS level_60up" +
    "  , level_50+level_40 + level_30 + level_20 + level_10 + level_lt10 AS level_60down" +
    " FROM " +
    " 	rank_matrix" +
    " 	LEFT OUTER JOIN rank_detail" +
    " 		ON rank_detail.ref_matrix_id = rank_matrix.id" +
    " 	LEFT OUTER JOIN student" +
    " 		ON student.id = rank_detail.ref_student_id" +
    " 	LEFT OUTER JOIN class" +
    " 		ON class.id = student.ref_class_id" +
    " WHERE" +
    " 	rank_matrix.is_alive = true" +
    " 	AND rank_matrix.school_year = " + SchoolYear +
    "     AND rank_matrix.semester = " + Semester +
    " 	AND rank_matrix.item_type like '定期評量%'" +
    " 	AND rank_matrix.ref_exam_id = " + ExamID +
    "     AND ref_student_id IN (" + string.Join(",", StudentIDList.ToArray()) + ") " +
    " ORDER BY " +
    " 	rank_matrix.id" +
    " 	, rank_detail.rank" +
    " 	, class.grade_year" +
    " 	, class.display_order" +
    " 	, class.class_name" +
    " 	, student.seat_no" +
    " 	, student.id";

                DataTable dt = qh.Select(query);
                //dt.TableName = "d5";
                //dt.WriteXmlSchema(Application.StartupPath + "\\d5s.xml");
                //dt.WriteXml(Application.StartupPath + "\\d5d.xml");

                // student id key
                // key = item_type + item_name +  rank_name 
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["student_id"].ToString();
                    if (!value.ContainsKey(sid))
                        value.Add(sid, new Dictionary<string, DataRow>());

                    string key = dr["item_type"].ToString() + "_" + dr["item_name"].ToString() + "_" + dr["rank_type"].ToString();

                    if (!value[sid].ContainsKey(key))
                        value[sid].Add(key, dr);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        // 取得中英文對照表中文科目名稱(排序用)
        public static List<string> GetChineseSubjectNameList()
        {
            List<string> value = new List<string>();

            /* content 格式：
            <Content>
                <Subject Chinese="國文 " English="Chinese"/>
            </Content> 
             */

            // 讀取 SQL 
            string strSQL = string.Format(@"
            SELECT
                content
            FROM
                list
            WHERE
                name = '科目中英文對照表'
            ");

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(strSQL);
            if (dt.Rows.Count > 0)
            {
                try
                {
                    XElement elmRoot = XElement.Parse(dt.Rows[0]["content"] + "");
                    foreach (XElement elm in elmRoot.Elements("Subject"))
                    {
                        if (elm.Attribute("Chinese") != null)
                            value.Add(elm.Attribute("Chinese").Value);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }

            return value;
        }

        // 取得學生課程規畫表內9D科目名稱
        public static Dictionary<string, List<string>> GetStudent9DSubjectNameByID(List<string> studentIDs)
        {
            Dictionary<string, List<string>> value = new Dictionary<string, List<string>>();

            try
            {
                if (studentIDs.Count > 0)
                {
                    QueryHelper qh = new QueryHelper();
                    string query = string.Format(@"
                WITH student_data AS(
                    SELECT
                        student.id AS student_id,
                        COALESCE(
                            student.ref_graduation_plan_id,
                            class.ref_graduation_plan_id
                        ) AS graduation_plan_id
                    FROM
                        student
                        LEFT JOIN class ON student.ref_class_id = class.id
                    WHERE
                        student.id IN({0})
                ),
                graduation_plan_expand AS(
                    SELECT
                        graduation_plan_id,
                        array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: TEXT AS subject_name,
                        array_to_string(xpath('//Subject/@課程代碼', subject_ele), '') :: TEXT AS 課程代碼
                    FROM
                        (
                            SELECT
                                graduation_plan_id,
                                unnest(
                                    xpath(
                                        '//GraduationPlan/Subject',
                                        xmlparse(content graduation_plan.content)
                                    )
                                ) AS subject_ele
                            FROM
                                (
                                    SELECT
                                        DISTINCT graduation_plan_id
                                    FROM
                                        student_data
                                ) AS target_graduation_plan
                                INNER JOIN graduation_plan ON graduation_plan.id = target_graduation_plan.graduation_plan_id
                        ) AS graduation_plan_expand
                )
                SELECT
                    DISTINCT student_data.student_id,
                    graduation_plan_expand.subject_name                    
                FROM
                    student_data
                    LEFT JOIN graduation_plan_expand ON student_data.graduation_plan_id = graduation_plan_expand.graduation_plan_id
                WHERE
                    SUBSTRING(graduation_plan_expand.課程代碼, 17, 1) = '9'
                    AND SUBSTRING(graduation_plan_expand.課程代碼, 19, 1) = 'D'
                ORDER BY
                    subject_name
                ", string.Join(",", studentIDs.ToArray()));

                    DataTable dt = qh.Select(query);
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["student_id"] + "";
                        string SubjectName = dr["subject_name"] + "";
                        if (!value.ContainsKey(sid))
                            value.Add(sid, new List<string>());

                        if (!value[sid].Contains(SubjectName))
                            value[sid].Add(SubjectName);

                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }
    }
}
