using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml.Linq;

namespace SmartSchool.Evaluation.Process.Wizards.LearningHistory
{
    public class Utility
    {
        public static string GetAttribute(XElement elm, string name)
        {
            string value = "";
            if (elm.Attribute(name) != null)
                value = elm.Attribute(name).Value;

            return value;
        }

        /// <summary>
        /// 取得異動與身分對照
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetUpdateCodeMappingDict()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();

            try
            {
                XElement elmRoot = XElement.Parse(Properties.Resources.StudTypeMapping1);
                if (elmRoot != null)
                {
                    foreach (XElement elm in elmRoot.Elements("Item"))
                    {
                        string code = "";
                        string val = "";
                        if (elm.Attribute("異動代碼") != null)
                            code = elm.Attribute("異動代碼").Value;

                        if (elm.Attribute("身分別") != null)
                            val = elm.Attribute("身分別").Value;

                        if (!value.ContainsKey(code))
                            value.Add(code, val);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return value;

        }

        public static Dictionary<string, string> GetUpdateCodeMappingDict2()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();

            try
            {
                XElement elmRoot = XElement.Parse(Properties.Resources.StudTypeMapping2);
                if (elmRoot != null)
                {
                    foreach (XElement elm in elmRoot.Elements("Item"))
                    {
                        string code = "";
                        string val = "";
                        if (elm.Attribute("異動代碼") != null)
                            code = elm.Attribute("異動代碼").Value;

                        if (elm.Attribute("原因及事項") != null)
                            val = elm.Attribute("原因及事項").Value;

                        if (!value.ContainsKey(code))
                            value.Add(code, val);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return value;

        }

        /// <summary>
        /// 取得復學、轉學、轉科、重讀異動
        /// </summary>
        /// <returns></returns>
        public static Dictionary<string, string> GetUpdateCodeMappingDict3()
        {
            Dictionary<string, string> value = new Dictionary<string, string>();

            try
            {
                XElement elmRoot = XElement.Parse(Properties.Resources.StudTypeMapping3);
                if (elmRoot != null)
                {
                    foreach (XElement elm in elmRoot.Elements("Item"))
                    {
                        string code = "";
                        string val = "";
                        if (elm.Attribute("異動代碼") != null)
                            code = elm.Attribute("異動代碼").Value;

                        if (elm.Attribute("原因及事項") != null)
                            val = elm.Attribute("原因及事項").Value;

                        if (!value.ContainsKey(code))
                            value.Add(code, val);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

            return value;

        }

        public static Dictionary<string, string> GetStudentHasUpdateCodeDict(int SchoolYear, int Semester, List<string> StudentIDList, List<string> UpdateCodeList)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            string sql = "SELECT " +
                "ref_student_id AS student_id" +
                ",update_code FROM " +
                "update_record " +
                "WHERE school_year = " + SchoolYear + " AND semester =" + Semester + " " +
                "AND ref_student_id IN(" + string.Join(",", StudentIDList.ToArray()) + ") " +
                "AND update_code IN('" + string.Join("','", UpdateCodeList.ToArray()) + "') ORDER BY ref_student_id,update_code;";
            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(sql);
            if (dt != null)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["student_id"] + "";
                    string code = dr["update_code"] + "";
                    if (!value.ContainsKey(sid))
                        value.Add(sid, code);
                }
            }


            return value;
        }

        /// <summary>
        /// 將西元日期轉成民國數字
        /// </summary>
        /// <param name="dt"></param>
        /// <returns></returns>
        public static string ConvertChDateString(DateTime? dt)
        {
            string value = "";
            if (dt.HasValue)
            {
                value = string.Format("{0:000}", (dt.Value.Year - 1911)) + string.Format("{0:00}", dt.Value.Month) + string.Format("{0:00}", dt.Value.Day);
            }

            return value;
        }

        /// <summary>
        /// 透過異動代碼找到符合學生及學年度學期
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="UpdateCodeList"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetStudentHasUpdateCodeDict(List<string> StudentIDList, List<string> UpdateCodeList)
        {
            Dictionary<string, string> value = new Dictionary<string, string>();
            string sql = @"
SELECT * FROM  (
    SELECT 
        ROW_NUMBER() OVER (PARTITION BY ref_student_id ORDER BY update_date DESC) as ROW_ID 
        , ref_student_id AS student_id
        , update_code 
        , school_year
        , semester
        , update_date
        , school_year::text || semester:: text ::int AS year_sems
    FROM update_record 
    WHERE 
     ref_student_id IN( " + string.Join(",", StudentIDList.ToArray()) + "  )" +
     " AND update_code  IN('" + string.Join("','", UpdateCodeList.ToArray()) + "') " +
     " ORDER BY ref_student_id, update_date DESC ) AS data  WHERE ROW_ID = 1; ";

            QueryHelper qh = new QueryHelper();
            DataTable dt = qh.Select(sql);
            if (dt != null)
            {
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["student_id"] + "";
                    string year_sems = dr["year_sems"].ToString();
                    if (!value.ContainsKey(sid))
                        value.Add(sid, year_sems);
                }
            }


            return value;
        }

        /// <summary>
        /// 透過學生 ID 與 學年度，取得學生該學年度學年科目成績
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, decimal>> GetStudentYearScoreByStudentIDDict(int SchoolYear, List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, decimal>> value = new Dictionary<string, Dictionary<string, decimal>>();

            try
            {
                QueryHelper qh = new QueryHelper();
                string sql = "SELECT ref_student_id AS student_id,school_year,grade_year,score_info FROM year_subj_score WHERE ref_student_id IN(" + string.Join(",", StudentIDList.ToArray()) + ") AND school_year = " + SchoolYear + ";";
                DataTable dt = qh.Select(sql);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["student_id"] + "";
                        XElement elmRoot = XElement.Parse(dr["score_info"] + "");
                        if (elmRoot != null)
                        {
                            foreach (XElement elm in elmRoot.Elements("Subject"))
                            {
                                string subjName = "";
                                decimal score;
                                if (elm.Attribute("科目") != null)
                                    subjName = elm.Attribute("科目").Value;

                                if (elm.Attribute("學年成績") != null)
                                {
                                    if (decimal.TryParse(elm.Attribute("學年成績").Value, out score))
                                    {
                                        if (!value.ContainsKey(sid))
                                            value.Add(sid, new Dictionary<string, decimal>());

                                        if (!value[sid].ContainsKey(subjName))
                                        {
                                            value[sid].Add(subjName, score);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }

            return value;
        }

        /// <summary>
        /// 透過學生 ID 與 學年度，取得學生該學年度學年科目補考成績
        /// </summary>
        /// <param name="SchoolYear"></param>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, decimal>> GetStudentYearReScoreByStudentIDDict(int SchoolYear, List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, decimal>> value = new Dictionary<string, Dictionary<string, decimal>>();

            try
            {
                QueryHelper qh = new QueryHelper();
                string sql = "SELECT ref_student_id AS student_id,school_year,grade_year,score_info FROM year_subj_score WHERE ref_student_id IN(" + string.Join(",", StudentIDList.ToArray()) + ") AND school_year = " + SchoolYear + ";";
                DataTable dt = qh.Select(sql);
                if (dt != null)
                {
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["student_id"] + "";
                        XElement elmRoot = XElement.Parse(dr["score_info"] + "");
                        if (elmRoot != null)
                        {
                            foreach (XElement elm in elmRoot.Elements("Subject"))
                            {
                                string subjName = "";
                                decimal score;
                                if (elm.Attribute("科目") != null)
                                    subjName = elm.Attribute("科目").Value;

                                if (elm.Attribute("補考成績") != null)
                                {
                                    if (decimal.TryParse(elm.Attribute("補考成績").Value, out score))
                                    {
                                        if (!value.ContainsKey(sid))
                                            value.Add(sid, new Dictionary<string, decimal>());

                                        if (!value[sid].ContainsKey(subjName))
                                        {
                                            value[sid].Add(subjName, score);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

            }
            catch (Exception ex)
            {
                Console.Write(ex.Message);
            }

            return value;
        }
    }
}
