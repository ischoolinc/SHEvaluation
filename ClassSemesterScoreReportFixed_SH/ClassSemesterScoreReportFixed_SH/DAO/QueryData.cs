using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;
using System.Xml.Linq;

namespace ClassSemesterScoreReportFixed_SH.DAO
{
    public class QueryData
    {

        // 類別1名稱
        public static Dictionary<string, string> StudentTag1Dict = new Dictionary<string, string>();
        // 類別2名稱
        public static Dictionary<string, string> StudentTag2Dict = new Dictionary<string, string>();

        /// <summary>
        /// 透過班級編號取得該班級學生系統編號(一般生)
        /// </summary>
        /// <param name="ClassIDList"></param>
        /// <returns></returns>
        public static List<string> GetClassStudentIDByClassID(List<string> ClassIDList)
        {
            List<string> retVal = new List<string>();
            if (ClassIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strQuery = "select student.id from student inner join class on student.ref_class_id=class.id where student.status=1 and class.id in(" + string.Join(",", ClassIDList.ToArray()) + ")";
                DataTable dt = qh.Select(strQuery);
                foreach (DataRow dr in dt.Rows)
                    retVal.Add(dr[0].ToString());
            }
            return retVal;
        }

        /// <summary>
        /// 透過學生學生系統編號，取得學年度、學期，(科目排名)
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, DataRow> GetStudentSemesterSubjectScoreRowBySchoolYearSemester(List<string> StudentIDList, int SchoolYear, int Semester)
        {
            Dictionary<string, DataRow> retVal = new Dictionary<string, DataRow>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strQuery = "select ref_student_id,score_info,class_rating,dept_rating,year_rating,group_rating from sems_subj_score where ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ") and school_year=" + SchoolYear + " and semester=" + Semester;
                DataTable dt = qh.Select(strQuery);
                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr["ref_student_id"].ToString();
                    if (!retVal.ContainsKey(id))
                        retVal.Add(id, dr);
                }

                Global._TempSubjClassRankDict.Clear();
                Global._TempSubjDeptRankDict.Clear();
                Global._TempSubjGradeYearRankDict.Clear();
                Global._TempSubjGroup1RankDict.Clear();

                foreach (string id in retVal.Keys)
                {
                    DataRow dr = retVal[id];

                    #region 科目班排名
                    if (!Global._TempSubjClassRankDict.ContainsKey(id))
                        Global._TempSubjClassRankDict.Add(id, new Dictionary<string, ScoreItem>());

                    if (dr["class_rating"] != null)
                        if (!string.IsNullOrEmpty(dr["class_rating"].ToString()))
                        {
                            XElement elmClass = XElement.Parse(dr["class_rating"].ToString());
                            foreach (XElement elm in elmClass.Elements("Item"))
                            {
                                int g1, g2;
                                decimal c1;
                                string key = elm.Attribute("科目").Value + elm.Attribute("科目級別").Value;
                                ScoreItem si = new ScoreItem();
                                si.Name = key;
                                if (!Global._TempSubjClassRankDict[id].ContainsKey(key))
                                {
                                    int.TryParse(elm.Attribute("排名").Value, out g1);
                                    int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                    si.Rank = g1;
                                    si.RankT = g2;
                                    si.Score = c1;
                                    Global._TempSubjClassRankDict[id].Add(key, si);
                                }
                            }
                        }
                    #endregion

                    #region 科目科排名
                    if (!Global._TempSubjDeptRankDict.ContainsKey(id))
                        Global._TempSubjDeptRankDict.Add(id, new Dictionary<string, ScoreItem>());

                    if (dr["dept_rating"] != null)
                        if (!string.IsNullOrEmpty(dr["dept_rating"].ToString()))
                        {
                            XElement elmDept = XElement.Parse(dr["dept_rating"].ToString());
                            foreach (XElement elm in elmDept.Elements("Item"))
                            {
                                int g1, g2;
                                decimal c1;
                                string key = elm.Attribute("科目").Value + elm.Attribute("科目級別").Value;
                                ScoreItem si = new ScoreItem();
                                si.Name = key;
                                if (!Global._TempSubjDeptRankDict[id].ContainsKey(key))
                                {
                                    int.TryParse(elm.Attribute("排名").Value, out g1);
                                    int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                    si.Rank = g1;
                                    si.RankT = g2;
                                    si.Score = c1;
                                    Global._TempSubjDeptRankDict[id].Add(key, si);
                                }
                            }
                        }
                    #endregion

                    #region 科目年排名
                    if (!Global._TempSubjGradeYearRankDict.ContainsKey(id))
                        Global._TempSubjGradeYearRankDict.Add(id, new Dictionary<string, ScoreItem>());

                    if (dr["year_rating"] != null)
                        if (!string.IsNullOrEmpty(dr["year_rating"].ToString()))
                        {
                            XElement elmGradeYear = XElement.Parse(dr["year_rating"].ToString());
                            foreach (XElement elm in elmGradeYear.Elements("Item"))
                            {
                                int g1, g2;
                                decimal c1;
                                string key = elm.Attribute("科目").Value + elm.Attribute("科目級別").Value;
                                ScoreItem si = new ScoreItem();
                                si.Name = key;
                                if (!Global._TempSubjGradeYearRankDict[id].ContainsKey(key))
                                {
                                    int.TryParse(elm.Attribute("排名").Value, out g1);
                                    int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                    si.Rank = g1;
                                    si.RankT = g2;
                                    si.Score = c1;
                                    Global._TempSubjGradeYearRankDict[id].Add(key, si);
                                }
                            }
                        }
                    #endregion

                    #region 科目類1排名
                    if (!Global._TempSubjGroup1RankDict.ContainsKey(id))
                        Global._TempSubjGroup1RankDict.Add(id, new Dictionary<string, ScoreItem>());

                    if (dr["group_rating"] != null)
                        if (!string.IsNullOrEmpty(dr["group_rating"].ToString()))
                        {
                            int cnCount = 1;
                            XElement elmGroup1 = XElement.Parse(dr["group_rating"].ToString());
                            foreach (XElement elm in elmGroup1.Elements("Item"))
                            {
                                int g1, g2;
                                decimal c1;
                                string key = elm.Attribute("科目").Value + elm.Attribute("科目級別").Value + "_類別" + cnCount;
                                ScoreItem si = new ScoreItem();
                                si.Name = key;
                                if (!Global._TempSubjGroup1RankDict[id].ContainsKey(key))
                                {
                                    int.TryParse(elm.Attribute("排名").Value, out g1);
                                    int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                    si.Rank = g1;
                                    si.RankT = g2;
                                    si.Score = c1;
                                    Global._TempSubjGroup1RankDict[id].Add(key, si);
                                }
                            }
                            cnCount++;
                        }
                    #endregion
                }

            }
            return retVal;
        }

        /// <summary>
        /// 透過學生學生系統編號，取得學年度、學期，(分項排名)
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, DataRow> GetStudentSemesterScoreRowBySchoolYearSemester(List<string> StudentIDList, int SchoolYear, int Semester)
        {
            Dictionary<string, DataRow> retVal = new Dictionary<string, DataRow>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strQuery = "select ref_student_id,score_info,class_rating,dept_rating,year_rating,group_rating from sems_entry_score where entry_group=1 and ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ") and school_year=" + SchoolYear + " and semester=" + Semester;
                DataTable dt = qh.Select(strQuery);
                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr["ref_student_id"].ToString();
                    if (!retVal.ContainsKey(id))
                        retVal.Add(id, dr);
                }

                Global._TempStudentSemesScoreDict.Clear();
                Global._TempClassRankDict.Clear();
                Global._TempDeptRankDict.Clear();
                Global._TempGradeYearRankDict.Clear();
                Global._TempGroup1RankDict.Clear();


                foreach (string id in retVal.Keys)
                {
                    DataRow dr = retVal[id];

                    #region 學期各項成績
                    if (!Global._TempStudentSemesScoreDict.ContainsKey(id))
                        Global._TempStudentSemesScoreDict.Add(id, new Dictionary<string, decimal>());
                    if (dr["score_info"] != null)
                        if (!string.IsNullOrEmpty(dr["score_info"].ToString()))
                        {
                            XElement elmScore = XElement.Parse(dr["score_info"].ToString());
                            foreach (XElement elm in elmScore.Elements("Entry"))
                            {
                                decimal c1;
                                string key = elm.Attribute("分項").Value;
                               
                                if (!Global._TempStudentSemesScoreDict[id].ContainsKey(key))
                                {                                  
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);                                  
                                    Global._TempStudentSemesScoreDict[id].Add(key, c1);
                                }
                            }
                        }


                    #endregion


                    #region 學期班排名
                    if (!Global._TempClassRankDict.ContainsKey(id))
                        Global._TempClassRankDict.Add(id, new Dictionary<string, ScoreItem>());

                    if (dr["class_rating"] != null)
                        if (!string.IsNullOrEmpty(dr["class_rating"].ToString()))
                        {
                            XElement elmClass = XElement.Parse(dr["class_rating"].ToString());
                            foreach (XElement elm in elmClass.Elements("Item"))
                            {
                                int g1, g2;
                                decimal c1;
                                string key = elm.Attribute("分項").Value;
                                ScoreItem si = new ScoreItem();
                                si.Name = key;
                                if (!Global._TempClassRankDict[id].ContainsKey(key))
                                {
                                    int.TryParse(elm.Attribute("排名").Value, out g1);
                                    int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                    si.Rank = g1;
                                    si.RankT = g2;
                                    si.Score = c1;
                                    Global._TempClassRankDict[id].Add(key, si);
                                }
                            }
                        }
                    #endregion

                    #region 學期科排名
                    if (!Global._TempDeptRankDict.ContainsKey(id))
                        Global._TempDeptRankDict.Add(id, new Dictionary<string, ScoreItem>());

                    if (dr["dept_rating"] != null)
                        if (!string.IsNullOrEmpty(dr["dept_rating"].ToString()))
                        {
                            XElement elmDept = XElement.Parse(dr["dept_rating"].ToString());
                            foreach (XElement elm in elmDept.Elements("Item"))
                            {
                                int g1, g2;
                                decimal c1;
                                string key = elm.Attribute("分項").Value;
                                ScoreItem si = new ScoreItem();
                                si.Name = key;
                                if (!Global._TempDeptRankDict[id].ContainsKey(key))
                                {
                                    int.TryParse(elm.Attribute("排名").Value, out g1);
                                    int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                    si.Rank = g1;
                                    si.RankT = g2;
                                    si.Score = c1;
                                    Global._TempDeptRankDict[id].Add(key, si);
                                }
                            }
                        }
                    #endregion

                    #region 學期年排名
                    if (!Global._TempGradeYearRankDict.ContainsKey(id))
                        Global._TempGradeYearRankDict.Add(id, new Dictionary<string, ScoreItem>());

                    if (dr["year_rating"] != null)
                        if (!string.IsNullOrEmpty(dr["year_rating"].ToString()))
                        {
                            XElement elmGradeYear = XElement.Parse(dr["year_rating"].ToString());
                            foreach (XElement elm in elmGradeYear.Elements("Item"))
                            {
                                int g1, g2;
                                decimal c1;
                                string key = elm.Attribute("分項").Value;
                                ScoreItem si = new ScoreItem();
                                si.Name = key;
                                if (!Global._TempGradeYearRankDict[id].ContainsKey(key))
                                {
                                    int.TryParse(elm.Attribute("排名").Value, out g1);
                                    int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                    si.Rank = g1;
                                    si.RankT = g2;
                                    si.Score = c1;
                                    Global._TempGradeYearRankDict[id].Add(key, si);
                                }
                            }
                        }
                    #endregion

                    #region 學期年類1名
                    if (!Global._TempGroup1RankDict.ContainsKey(id))
                        Global._TempGroup1RankDict.Add(id, new Dictionary<string, ScoreItem>());

                    if (dr["group_rating"] != null)
                        if (!string.IsNullOrEmpty(dr["group_rating"].ToString()))
                        {
                            XElement elmGroup1 = XElement.Parse(dr["group_rating"].ToString());
                            int cnCount = 1;
                            foreach (XElement elmA in elmGroup1.Elements("Rating"))
                            {
                                string Group = elmA.Attribute("類別").Value;
                                foreach (XElement elm in elmA.Elements("Item"))
                                {
                                    int g1, g2;
                                    decimal c1;
                                    string key = elm.Attribute("分項").Value + "_類別" + cnCount;
                                    ScoreItem si = new ScoreItem();
                                    si.Name = Group;
                                    if (!Global._TempGroup1RankDict[id].ContainsKey(key))
                                    {
                                        int.TryParse(elm.Attribute("排名").Value, out g1);
                                        int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                        decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                        si.Rank = g1;
                                        si.RankT = g2;
                                        si.Score = c1;
                                        Global._TempGroup1RankDict[id].Add(key, si);
                                    }
                                }
                                cnCount++;
                            }
                        }
                    #endregion
                }

            }
            return retVal;
        }



        /// <summary>
        /// 取得學生學期排名、五標與組距資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, Dictionary<string, string>>> GetSemsScoreRankMatrixDataDict(string SchoolYear, string Semester, List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> value = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            // 沒有學生不處理
            if (StudentIDList.Count == 0)
                return value;

            List<string> r2List = new List<string>();
            r2List.Add("rank");
            r2List.Add("matrix_count");
            r2List.Add("pr");
            r2List.Add("percentile");
            r2List.Add("avg_top_25");
            r2List.Add("avg_top_50");
            r2List.Add("avg");
            r2List.Add("avg_bottom_50");
            r2List.Add("avg_bottom_25");
            r2List.Add("level_gte100");
            r2List.Add("level_90");
            r2List.Add("level_80");
            r2List.Add("level_70");
            r2List.Add("level_60");
            r2List.Add("level_50");
            r2List.Add("level_40");
            r2List.Add("level_30");
            r2List.Add("level_20");
            r2List.Add("level_10");
            r2List.Add("level_lt10");

            // 需要四捨五入
            List<string> r2ListNP = new List<string>();
            r2ListNP.Add("avg_top_25");
            r2ListNP.Add("avg_top_50");
            r2ListNP.Add("avg");
            r2ListNP.Add("avg_bottom_50");
            r2ListNP.Add("avg_bottom_25");

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
" 	AND rank_matrix.item_type like '學期%'" +
" 	AND rank_matrix.ref_exam_id = -1 " +
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

            foreach (DataRow dr in dt.Rows)
            {
                string sid = dr["student_id"].ToString();
                if (!value.ContainsKey(sid))
                    value.Add(sid, new Dictionary<string, Dictionary<string, string>>());

                string key = dr["item_type"].ToString() + "_" + dr["item_name"].ToString() + "_" + dr["rank_type"].ToString();
                if (key == "學期/分項成績_學業_類別1排名")
                {
                    if (dr["rank_name"] != null)
                    {
                        if (!StudentTag1Dict.ContainsKey(sid))
                            StudentTag1Dict.Add(sid, dr["rank_name"].ToString());
                    }
                }

                if (key == "學期/分項成績_學業_類別2排名")
                {
                    if (dr["rank_name"] != null)
                    {
                        if (!StudentTag2Dict.ContainsKey(sid))
                            StudentTag2Dict.Add(sid, dr["rank_name"].ToString());
                    }
                }
                if (!value[sid].ContainsKey(key))
                    value[sid].Add(key, new Dictionary<string, string>());

                foreach (string r2 in r2List)
                {
                    string dValue = "";
                    if (dr[r2] != null)
                    {
                        if (r2ListNP.Contains(r2))
                        {
                            decimal dd;
                            if (decimal.TryParse(dr[r2].ToString(), out dd))
                            {
                                dValue = Math.Round(dd, 2, MidpointRounding.AwayFromZero).ToString();
                            }

                        }
                        else
                        {
                            dValue = dr[r2].ToString();
                        }
                    }

                    if (!value[sid][key].ContainsKey(r2))
                        value[sid][key].Add(r2, dValue);
                }
            }
            return value;
        }


        public static Dictionary<string, Dictionary<string, Dictionary<string, string>>> GetSemsScoreRankMatrixDataByClassIDDict(string SchoolYear, string Semester, List<string> ClassIDList)
        {
            Dictionary<string, Dictionary<string, Dictionary<string, string>>> value = new Dictionary<string, Dictionary<string, Dictionary<string, string>>>();

            // 沒有班級不處理
            if (ClassIDList.Count == 0)
                return value;

            List<string> r2List = new List<string>();
            r2List.Add("matrix_count");
            r2List.Add("avg_top_25");
            r2List.Add("avg_top_50");
            r2List.Add("avg");
            r2List.Add("avg_bottom_50");
            r2List.Add("avg_bottom_25");
            r2List.Add("level_gte100");
            r2List.Add("level_90");
            r2List.Add("level_80");
            r2List.Add("level_70");
            r2List.Add("level_60");
            r2List.Add("level_50");
            r2List.Add("level_40");
            r2List.Add("level_30");
            r2List.Add("level_20");
            r2List.Add("level_10");
            r2List.Add("level_lt10");

            // 需要四捨五入
            List<string> r2ListNP = new List<string>();
            r2ListNP.Add("avg_top_25");
            r2ListNP.Add("avg_top_50");
            r2ListNP.Add("avg");
            r2ListNP.Add("avg_bottom_50");
            r2ListNP.Add("avg_bottom_25");

            QueryHelper qh = new QueryHelper();
            string query = "" +
              " SELECT  " +
"     DISTINCT rank_matrix.id AS rank_matrix_id " +
"       , class.id AS class_id " +
"     , rank_matrix.school_year " +
"     , rank_matrix.semester " +
"     , rank_matrix.grade_year " +
"     , rank_matrix.item_type " +
"     , rank_matrix.item_name " +
"     , rank_matrix.rank_type " +
"     , rank_matrix.rank_name " +
"     , class.class_name " +
"   , rank_matrix.matrix_count  " +
"   , rank_matrix.avg_top_25 " +
"   , rank_matrix.avg_top_50 " +
"   , rank_matrix.avg " +
"   , rank_matrix.avg_bottom_50 " +
"   , rank_matrix.avg_bottom_25 " +
"     , rank_matrix.level_gte100 " +
"     , rank_matrix.level_90 " +
"     , rank_matrix.level_80 " +
"     , rank_matrix.level_70 " +
"     , rank_matrix.level_60 " +
"     , rank_matrix.level_50 " +
"     , rank_matrix.level_40 " +
"     , rank_matrix.level_30 " +
"     , rank_matrix.level_20 " +
"     , rank_matrix.level_10 " +
"     , rank_matrix.level_lt10 " +
" FROM  " +
"     rank_matrix " +
"     LEFT OUTER JOIN rank_detail " +
"           ON rank_detail.ref_matrix_id = rank_matrix.id " +
"     LEFT OUTER JOIN student " +
"           ON student.id = rank_detail.ref_student_id " +
"     LEFT OUTER JOIN class " +
"           ON class.id = student.ref_class_id " +
" WHERE " +
"     rank_matrix.is_alive = true " +
"     AND rank_matrix.school_year = " + SchoolYear + " " +
"     AND rank_matrix.semester = " + Semester + " " +
"     AND rank_matrix.item_type like '學期%' " +
"     AND rank_matrix.ref_exam_id = -1  " +
"     AND class.id IN (" + string.Join(",", ClassIDList.ToArray()) + "); ";

            DataTable dt = qh.Select(query);

            foreach (DataRow dr in dt.Rows)
            {
                string cid = dr["class_id"].ToString();
                if (!value.ContainsKey(cid))
                    value.Add(cid, new Dictionary<string, Dictionary<string, string>>());

                string key = dr["item_type"].ToString() + "_" + dr["item_name"].ToString() + "_" + dr["rank_type"].ToString();
                if (key == "學期/分項成績_學業_類別1排名")
                {
                    if (dr["rank_name"] != null)
                    {
                        if (!StudentTag1Dict.ContainsKey(cid))
                            StudentTag1Dict.Add(cid, dr["rank_name"].ToString());
                    }
                }

                if (key == "學期/分項成績_學業_類別2排名")
                {
                    if (dr["rank_name"] != null)
                    {
                        if (!StudentTag2Dict.ContainsKey(cid))
                            StudentTag2Dict.Add(cid, dr["rank_name"].ToString());
                    }
                }
                if (!value[cid].ContainsKey(key))
                    value[cid].Add(key, new Dictionary<string, string>());

                foreach (string r2 in r2List)
                {
                    string dValue = "";
                    if (dr[r2] != null)
                    {
                        if (r2ListNP.Contains(r2))
                        {
                            decimal dd;
                            if (decimal.TryParse(dr[r2].ToString(), out dd))
                            {
                                dValue = Math.Round(dd, 2, MidpointRounding.AwayFromZero).ToString();
                            }

                        }
                        else
                        {
                            dValue = dr[r2].ToString();
                        }
                    }

                    if (!value[cid][key].ContainsKey(r2))
                        value[cid][key].Add(r2, dValue);
                }
            }
            return value;
        }





    }
}
