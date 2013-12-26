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
                string strQuery = "select ref_student_id,score_info,class_rating,dept_rating,year_rating,group_rating from sems_subj_score where ref_student_id in("+string.Join(",",StudentIDList.ToArray())+") and school_year="+SchoolYear+" and semester="+Semester;
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
                                string key = elm.Attribute("科目").Value + elm.Attribute("科目級別").Value+"_類別"+cnCount;
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
                string strQuery = "select ref_student_id,score_info,class_rating,dept_rating,year_rating,group_rating from sems_entry_score where entry_group=1 and ref_student_id in("+string.Join(",",StudentIDList.ToArray())+") and school_year="+SchoolYear+" and semester=" + Semester;
                DataTable dt = qh.Select(strQuery);
                foreach (DataRow dr in dt.Rows)
                {
                    string id = dr["ref_student_id"].ToString();
                    if (!retVal.ContainsKey(id))
                        retVal.Add(id, dr);
                }

                Global._TempClassRankDict.Clear();                
                Global._TempDeptRankDict.Clear();                
                Global._TempGradeYearRankDict.Clear();                
                Global._TempGroup1RankDict.Clear();
                

                foreach (string id in retVal.Keys)
                {
                    DataRow dr = retVal[id];

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



    }
}
