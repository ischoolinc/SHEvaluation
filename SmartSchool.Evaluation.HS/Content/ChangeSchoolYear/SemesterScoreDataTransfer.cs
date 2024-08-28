using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Data;
using FISCA.Data;
using System.Windows.Forms;
using System.Security.Cryptography;

namespace SmartSchool.Evaluation.Content.ChangeSchoolYear
{
    public class SemesterScoreDataTransfer
    {

        // 檢查並更新學生學期分項成績學年度
        public int CheckAndUpdateStudentSemesterScoreSchoolYear(SemesterScoreData sourceData, string ChangeSchoolYear)
        {
            int value = -1;
            try
            {
                // 檢查欲調整學年度是否有資料

                string query1 = string.Format(@"
                SELECT
                    id
                FROM
                    sems_entry_score
                WHERE
                    ref_student_id = {0}
                    AND school_year = {1}
                    AND semester = {2}                    
                ", sourceData.StudentID, ChangeSchoolYear, sourceData.Semester);

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query1);
                // 沒有資料，搬過去
                if (dt.Rows.Count == 0)
                {
                    // 取得原來分項id來更新
                    string querySource = string.Format(@"
                    SELECT
                        id
                    FROM
                        sems_entry_score
                    WHERE
                        ref_student_id = {0}
                        AND school_year = {1}
                        AND semester = {2}                        
                ", sourceData.StudentID, sourceData.SchoolYear, sourceData.Semester);

                    DataTable dtSemsEnrty = qh.Select(querySource);

                    // 調整資料，回傳id
                    if (dtSemsEnrty.Rows.Count > 0)
                    {
                        int SemsEntryID;
                        if (int.TryParse(dtSemsEnrty.Rows[0]["id"].ToString(), out SemsEntryID) && !string.IsNullOrEmpty(ChangeSchoolYear))
                        {
                            string updateQuery = string.Format(@"
                        UPDATE
                            sems_entry_score
                        SET
                            school_year = {1} 
                        WHERE
                            id = {0} RETURNING id;
                        ", SemsEntryID, ChangeSchoolYear);

                            DataTable dtUpdate = qh.Select(updateQuery);

                            // 調整資料，回傳id
                            if (dtUpdate.Rows.Count > 0)
                            {
                                int.TryParse(dtUpdate.Rows[0]["id"].ToString(), out value);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        // 更新學生學期成績學年度
        public int UpdateStudentSemesterScoreSchoolYearBySemsID(string SemsScoreID, string SchoolYear)
        {
            int value = -1;
            try
            {
                string query = string.Format(@"
                UPDATE sems_subj_score SET school_year = {0}
                WHERE id = {1} RETURNING id;
                ", SchoolYear, SemsScoreID);

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                if (dt.Rows.Count > 0)
                {
                    int.TryParse(dt.Rows[0][0].ToString(), out value);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        // 新增成績資料
        public int InsertData(SemesterScoreData data)
        {
            int value = -1;
            try
            {
                string query = string.Format(@"
                INSERT INTO sems_subj_score(
                    ref_student_id
                    ,school_year
                    ,semester
                    ,grade_year
                    ,score_info
                )
                VALUES(
                    {0}
                    ,{1}
                    ,{2}
                    ,{3}
                    ,'{4}'
                ) RETURNING id;", data.StudentID, data.SchoolYear, data.Semester, data.GradeYear, data.ScoreInfo);
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                if (dt.Rows.Count > 0)
                {
                    int.TryParse(dt.Rows[0][0].ToString(), out value);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        // 依學生ID，學年度、學期、成績年級，取得成績資料
        public SemesterScoreData GetStudentScoreDataByIDSchoolYear(SemesterScoreData data)
        {
            try
            {
                string query = string.Format(@"
            SELECT
                sems_subj_score.id,
                ref_student_id AS student_id,
                school_year,
                semester,
                grade_year,
                score_info
            FROM
                sems_subj_score
            WHERE
                ref_student_id = {0}
                AND school_year = {1}
                AND semester = {2} ;
            ", data.StudentID, data.SchoolYear, data.Semester);

                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                if (dt.Rows.Count > 0)
                {
                    data.ID = dt.Rows[0]["id"] + "";
                    data.ScoreInfo = dt.Rows[0]["score_info"] + "";
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return data;
        }

        // 透過學生成績ID更新成績資料
        public int UpdateScoreDataBySemesScoreID(SemesterScoreData data)
        {
            int value = -1;
            try
            {
                string query = string.Format(@"
                UPDATE
                    sems_subj_score
                SET
                    score_info = '{1}'
                WHERE
                    id = {0} RETURNING id;
                ", data.ID, data.ScoreInfo);
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);
                if (dt.Rows.Count > 0)
                {
                    int.TryParse(dt.Rows[0][0].ToString(), out value);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        // 透過學生學期成績ID刪除學期成績
        public int DeleteSemesScoreBySemsID(string sems_id)
        {
            int value = -1;
            try
            {
                string query = string.Format(@"
                DELETE FROM
                    sems_subj_score
                WHERE
                    id = {0} RETURNING id;
            ", sems_id);
                QueryHelper qh = new QueryHelper();
                DataTable dt = qh.Select(query);

                if (dt.Rows.Count > 0)
                {
                    int.TryParse(dt.Rows[0][0].ToString(), out value);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }

            return value;
        }

        // 比對學期科目名稱與級別是否有相同
        public List<SubjectScoreInfo> CheckHasSubjectNameLevel(string sourceScoreInfo, string changeScoreInfo)
        {
            // 有相同回傳
            List<SubjectScoreInfo> value = new List<SubjectScoreInfo>();

            try
            {
                List<string> sourceSubjNameLit = new List<string>();
                XElement elmSource = XElement.Parse(sourceScoreInfo);
                XElement elmChange = XElement.Parse(changeScoreInfo);

                // 當預計寫入學年度有資料，使用預計寫入學年度學期成績科目+級別當 key 來檢查
                foreach (XElement elm in elmChange.Elements("Subject"))
                {
                    string key = elm.Attribute("科目").Value + "_" + elm.Attribute("科目級別").Value;
                    sourceSubjNameLit.Add(key);
                }

                // 檢查科目名稱+級別是否有相同
                foreach (XElement elm in elmSource.Elements("Subject"))
                {
                    string key = elm.Attribute("科目").Value + "_" + elm.Attribute("科目級別").Value;

                    if (sourceSubjNameLit.Contains(key))
                    {
                        SubjectScoreInfo ssi = new SubjectScoreInfo();
                        ssi.SubjectName = elm.Attribute("科目").Value;
                        ssi.SubjectLevel = elm.Attribute("科目級別").Value;
                        ssi.Credit = elm.Attribute("開課學分數").Value;
                        ssi.Required = elm.Attribute("修課必選修").Value;
                        ssi.RequiredBy = elm.Attribute("修課校部訂").Value;

                        if (ssi.RequiredBy == "部訂")
                            ssi.RequiredBy = "部定";

                        // 有值表示預計寫入已有相同科目名稱+級別
                        value.Add(ssi);
                    }
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        // 附加成績資料，將source家入change
        public string AppendScoreData(string sourceScoreInfo, string changeScoreInfo)
        {
            string value = "";
            try
            {
                XElement elmSource = XElement.Parse(sourceScoreInfo);
                XElement elmChange = null;
                if (string.IsNullOrEmpty(changeScoreInfo))
                    elmChange = new XElement("SemesterSubjectScoreInfo");
                else
                    elmChange = XElement.Parse(changeScoreInfo);

                foreach (XElement elm in elmSource.Elements("Subject"))
                {
                    elmChange.Add(elm);
                }

                value = elmChange.ToString();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            return value;
        }

        // 取得學生目前有成績學年度、學期、年級
        public static Dictionary<string, string> GetStudentHasSemsScoreSchoolYearSemesterByID(string StudentID)
        {
            // key:SchoolYear+Semester, value: GradeYear
            Dictionary<string, string> value = new Dictionary<string, string>();

            try
            {
                QueryHelper qh = new QueryHelper();
                string query = string.Format(@"
                SELECT
                    school_year,
                    semester,
                    grade_year
                FROM
                    sems_subj_score
                WHERE
                    ref_student_id = {0} 
                ", StudentID);

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string key = dr["school_year"] + "_" + dr["semester"];
                    if (!value.ContainsKey(key))
                        value.Add(key, dr["grade_year"] + "");
                }

            }
            catch (Exception ex)
            {
                Console.WriteLine("GetStudentHasSemsScoreSchoolYearSemesterByID," + ex.Message);
            }

            return value;
        }

    }
}
