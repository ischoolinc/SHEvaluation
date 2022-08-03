using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

namespace SH_ClassYearScoreReport.DAO
{
    public class QueryData
    {
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
        /// 透過學生學生系統編號，取得學年科目排名
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, DataRow> GetStudentYearSubjectScoreRowBySchoolYear(List<string> StudentIDList, string SchoolYear)
        {
            Dictionary<string, DataRow> retVal = new Dictionary<string, DataRow>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strQuery = "select ref_student_id,score_info,class_rating,dept_rating,year_rating,group_rating from year_subj_score where ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ") and school_year=" + SchoolYear;
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
                //Global._TempSubjGroup1RankDict.Clear();

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
                                string key = elm.Attribute("科目").Value;// + elm.Attribute("科目級別").Value;
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
                                string key = elm.Attribute("科目").Value;// + elm.Attribute("科目級別").Value;
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
                                string key = elm.Attribute("科目").Value;// + elm.Attribute("科目級別").Value;
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
                    //if (!Global._TempSubjGroup1RankDict.ContainsKey(id))
                    //    Global._TempSubjGroup1RankDict.Add(id, new Dictionary<string, ScoreItem>());

                    //if (dr["group_rating"] != null)
                    //    if (!string.IsNullOrEmpty(dr["group_rating"].ToString()))
                    //    {
                    //        int cnCount = 1;
                    //        XElement elmGroup1 = XElement.Parse(dr["group_rating"].ToString());
                    //        foreach (XElement elm in elmGroup1.Elements("Item"))
                    //        {
                    //            int g1, g2;
                    //            decimal c1;
                    //            string key = elm.Attribute("科目").Value + elm.Attribute("科目級別").Value + "_類別" + cnCount;
                    //            ScoreItem si = new ScoreItem();
                    //            si.Name = key;
                    //            if (!Global._TempSubjGroup1RankDict[id].ContainsKey(key))
                    //            {
                    //                int.TryParse(elm.Attribute("排名").Value, out g1);
                    //                int.TryParse(elm.Attribute("成績人數").Value, out g2);
                    //                decimal.TryParse(elm.Attribute("成績").Value, out c1);
                    //                si.Rank = g1;
                    //                si.RankT = g2;
                    //                si.Score = c1;
                    //                Global._TempSubjGroup1RankDict[id].Add(key, si);
                    //            }
                    //        }
                    //        cnCount++;
                    //    }
                    #endregion
                }

            }
            return retVal;
        }

        /// <summary>
        /// 透過學生學生系統編號，取得學年分項排名
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, DataRow> GetStudentYearEntryScoreRowBySchoolYear(List<string> StudentIDList, string SchoolYear)
        {
            Dictionary<string, DataRow> retVal = new Dictionary<string, DataRow>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strQuery = "select ref_student_id,score_info,class_rating,dept_rating,year_rating,group_rating from year_entry_score where ref_student_id in(" + string.Join(",", StudentIDList.ToArray()) + ") and school_year=" + SchoolYear;
                DataTable dt = qh.Select(strQuery);
                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr["ref_student_id"].ToString();
                    if (!retVal.ContainsKey(id))
                        retVal.Add(id, dr);
                }

                Global._TempEntryClassRankDict.Clear();
                Global._TempEntryDeptRankDict.Clear();
                Global._TempEntryGradeYearRankDict.Clear();
               // Global._TempSubjGroup1RankDict.Clear();

                foreach (string id in retVal.Keys)
                {
                    DataRow dr = retVal[id];

                    #region 分項班排名
                    if (!Global._TempEntryClassRankDict.ContainsKey(id))
                        Global._TempEntryClassRankDict.Add(id, new Dictionary<string, ScoreItem>());

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
                                if (!Global._TempEntryClassRankDict[id].ContainsKey(key))
                                {
                                    int.TryParse(elm.Attribute("排名").Value, out g1);
                                    int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                    si.Rank = g1;
                                    si.RankT = g2;
                                    si.Score = c1;
                                    Global._TempEntryClassRankDict[id].Add(key, si);
                                }
                            }
                        }
                    #endregion

                    #region 分項科排名
                    if (!Global._TempEntryDeptRankDict.ContainsKey(id))
                        Global._TempEntryDeptRankDict.Add(id, new Dictionary<string, ScoreItem>());

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
                                if (!Global._TempEntryDeptRankDict[id].ContainsKey(key))
                                {
                                    int.TryParse(elm.Attribute("排名").Value, out g1);
                                    int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                    si.Rank = g1;
                                    si.RankT = g2;
                                    si.Score = c1;
                                    Global._TempEntryDeptRankDict[id].Add(key, si);
                                }
                            }
                        }
                    #endregion

                    #region 分項年排名
                    if (!Global._TempEntryGradeYearRankDict.ContainsKey(id))
                        Global._TempEntryGradeYearRankDict.Add(id, new Dictionary<string, ScoreItem>());

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
                                if (!Global._TempEntryGradeYearRankDict[id].ContainsKey(key))
                                {
                                    int.TryParse(elm.Attribute("排名").Value, out g1);
                                    int.TryParse(elm.Attribute("成績人數").Value, out g2);
                                    decimal.TryParse(elm.Attribute("成績").Value, out c1);
                                    si.Rank = g1;
                                    si.RankT = g2;
                                    si.Score = c1;
                                    Global._TempEntryGradeYearRankDict[id].Add(key, si);
                                }
                            }
                        }
                    #endregion

                }

            }
            return retVal;
        }

    }
}
