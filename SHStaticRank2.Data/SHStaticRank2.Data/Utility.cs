using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;
using System.Xml.Linq;

namespace SHStaticRank2.Data
{
    public class Utility
    {
        /// <summary>
        /// 不處理進位(word合併欄位中自行處理)
        /// </summary>
        /// <param name="score"></param>
        /// <returns></returns>
        public static decimal NoRound(decimal score)
        {
            return score;
        }

        /// <summary>
        /// 四捨五入到小數下二位
        /// </summary>
        /// <param name="score"></param>
        /// <returns></returns>
        public static decimal ParseD2(decimal score)
        {
            return Math.Round(score, 2, MidpointRounding.AwayFromZero);
        }

        /// <summary>
        /// 取得排名百分比：名次減一除母數後取左邊第一個整數
        /// </summary>
        /// <param name="rank"></param>
        /// <param name="total"></param>
        /// <returns></returns>
        public static int ParseRankPercent(int rank, int total)
        {
            int retVal = 0;
            if (total > 0 && rank > 0)
            {
                return (int)(Math.Floor(((decimal)rank - 1) * 100m / (decimal)total) + 1);

                //decimal rr = (decimal)(rank - 1);
                //decimal tt = (decimal)total;


                //decimal xR = Math.Round(rr * 100 / total, 0);
                //decimal x = rr * 100 / total + 1;

                //if (xR == x)
                //    retVal = (int)xR;
                //else
                //    retVal = (int)x;

                //retVal = (int)(Math.Floor((rr * 100) / tt));
            }
            return retVal;
        }


        /// <summary>
        /// 透過學生編號取得學測報名序號
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, string> GetStudentSATSerNoByStudentIDList(List<string> StudentIDList)
        {
            Dictionary<string, string> retValue = new Dictionary<string, string>();
            if(StudentIDList.Count>0)
            {
                try
                {
                    QueryHelper qh = new QueryHelper();
                    string query = "select ref_student_id as sid,sat_ser_no as sno from $sh.college.sat.student where ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "')";
                    DataTable dt = qh.Select(query);
                    foreach (DataRow dr in dt.Rows)
                    {
                        string sid = dr["sid"].ToString();
                        string sno = "";
                        if (dr["sno"] != null)
                            sno = dr["sno"].ToString();

                        if (!retValue.ContainsKey(sid))
                            retValue.Add(sid, sno);
                    }
                }
                catch (Exception ex)
                { }
            }
            return retValue;
        }

        public static Dictionary<string, List<StudSemsEntryRating>> GetStudSemsEntryRatingByStudentID(List<string> StudentIDList)
        {
            Dictionary<string, List<StudSemsEntryRating>> retValue = new Dictionary<string, List<StudSemsEntryRating>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id as sid,school_year,semester,grade_year,class_rating,dept_rating,year_rating,group_rating from sems_entry_score where ref_student_id in ("+string.Join(",",StudentIDList.ToArray())+") and entry_group=1  order by ref_student_id,grade_year,semester";                
              
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["sid"].ToString();

                    if (!retValue.ContainsKey(sid))
                        retValue.Add(sid, new List<StudSemsEntryRating>());

                    StudSemsEntryRating sser = new StudSemsEntryRating();
                    sser.StudentID = sid;
                    sser.SchoolYear = dr["school_year"].ToString();
                    sser.Semester = dr["semester"].ToString();
                    sser.GradeYear = dr["grade_year"].ToString();
                    sser.EntryName = "學業";
                    // Parse XML
                    try
                    {
                        string cStr = dr["class_rating"].ToString();
                        string dStr = dr["dept_rating"].ToString();
                        string yStr = dr["year_rating"].ToString();
                        string g1Str = dr["group_rating"].ToString();

                        // 班排
                        if (!string.IsNullOrEmpty(cStr))
                        {
                            XElement elmC = XElement.Parse(cStr);
                            foreach (XElement elm in elmC.Elements("Item"))
                            {
                                if (elm.Attribute("分項") != null && elm.Attribute("分項").Value == "學業")
                                {
                                    if (elm.Attribute("成績人數") != null && elm.Attribute("成績人數").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("成績人數").Value, out x))
                                            sser.ClassCount = x;

                                    }
                                    if (elm.Attribute("排名") != null && elm.Attribute("排名").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("排名").Value, out x))
                                            sser.ClassRank = x;
                                    }
                                }
                            }
                        }
                        
                        // 科排
                        if (!string.IsNullOrEmpty(dStr))
                        {
                            XElement elmD = XElement.Parse(dStr);
                            foreach (XElement elm in elmD.Elements("Item"))
                            {
                                if (elm.Attribute("分項") != null && elm.Attribute("分項").Value == "學業")
                                {
                                    if (elm.Attribute("成績人數") != null && elm.Attribute("成績人數").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("成績人數").Value, out x))
                                            sser.DeptCount = x;
                                    }
                                    if (elm.Attribute("排名") != null && elm.Attribute("排名").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("排名").Value, out x))
                                            sser.DeptRank = x;
                                    }
                                }
                            }
                        }

                        // 校排
                        if (!string.IsNullOrEmpty(yStr))
                        {
                            XElement elmY = XElement.Parse(yStr);
                            foreach (XElement elm in elmY.Elements("Item"))
                            {
                                if (elm.Attribute("分項") != null && elm.Attribute("分項").Value == "學業")
                                {
                                    if (elm.Attribute("成績人數") != null && elm.Attribute("成績人數").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("成績人數").Value, out x))
                                            sser.YearCount = x;
                                    }

                                    if (elm.Attribute("排名") != null && elm.Attribute("排名").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("排名").Value, out x))
                                            sser.YearRank = x;
                                    }
                                }
                            }
                        }

                        // 類1排
                        if (!string.IsNullOrEmpty(g1Str))
                        {
                            XElement elmG1 = XElement.Parse(g1Str);
                            foreach(XElement elmR in elmG1.Elements("Rating"))
                            foreach (XElement elm in elmR.Elements("Item"))
                            {
                                if (elm.Attribute("分項") != null && elm.Attribute("分項").Value == "學業")
                                {
                                    if (elm.Attribute("成績人數") != null && elm.Attribute("成績人數").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("成績人數").Value, out x))
                                            sser.Group1Count = x;
                                    }

                                    if (elm.Attribute("排名") != null && elm.Attribute("排名").Value != "")
                                    {
                                        int x;
                                        if (int.TryParse(elm.Attribute("排名").Value, out x))
                                            sser.Group1Rank = x;
                                    }
                                        
                                }
                            }
                        }
                    }
                    catch (Exception ex) { }

                    retValue[sid].Add(sser);
                }
            }
            return retValue;
        }

        public static Dictionary<string, List<StudSemsSubjRating>> GetStudSemsSubjRatingByStudentID(List<string> StudentIDList)
        {
            Dictionary<string, List<StudSemsSubjRating>> retValue = new Dictionary<string, List<StudSemsSubjRating>>();
            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id as sid,school_year,semester,grade_year,class_rating,dept_rating,year_rating,group_rating from sems_subj_score where ref_student_id in (" + string.Join(",", StudentIDList.ToArray()) + ")  order by ref_student_id,grade_year,semester";

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["sid"].ToString();

                    if (!retValue.ContainsKey(sid))
                        retValue.Add(sid, new List<StudSemsSubjRating>());

                    StudSemsSubjRating sser = new StudSemsSubjRating();
                    sser.StudentID = sid;
                    sser.SchoolYear = dr["school_year"].ToString();
                    sser.Semester = dr["semester"].ToString();
                    sser.GradeYear = dr["grade_year"].ToString();                    
                    // Parse XML
                    try
                    {
                        string cStr = dr["class_rating"].ToString();
                        string dStr = dr["dept_rating"].ToString();
                        string yStr = dr["year_rating"].ToString();
                        string g1Str = dr["group_rating"].ToString();

                        // 班排
                        if (!string.IsNullOrEmpty(cStr))
                        {
                            XElement elmC = XElement.Parse(cStr);
                            foreach (XElement elm in elmC.Elements("Item"))
                            {
                                if (elm.Attribute("科目") != null)
                                {
                                    if (elm.Attribute("成績人數") != null)
                                        sser.AddClassCount(elm.Attribute("科目").Value, int.Parse(elm.Attribute("成績人數").Value));

                                    if (elm.Attribute("排名") != null)
                                        sser.AddClassRank(elm.Attribute("科目").Value, int.Parse(elm.Attribute("排名").Value));
                                }
                            }
                        }

                        // 科排
                        if (!string.IsNullOrEmpty(dStr))
                        {
                            XElement elmC = XElement.Parse(dStr);
                            foreach (XElement elm in elmC.Elements("Item"))
                            {
                                if (elm.Attribute("科目") != null)
                                {
                                    if (elm.Attribute("成績人數") != null)
                                        sser.AddDeptCount(elm.Attribute("科目").Value, int.Parse(elm.Attribute("成績人數").Value));

                                    if (elm.Attribute("排名") != null)
                                        sser.AddDeptRank(elm.Attribute("科目").Value, int.Parse(elm.Attribute("排名").Value));
                                }
                            }
                        }
                        // 年排
                        if (!string.IsNullOrEmpty(yStr))
                        {
                            XElement elmC = XElement.Parse(yStr);
                            foreach (XElement elm in elmC.Elements("Item"))
                            {
                                if (elm.Attribute("科目") != null)
                                {
                                    if (elm.Attribute("成績人數") != null)
                                        sser.AddYearCount(elm.Attribute("科目").Value, int.Parse(elm.Attribute("成績人數").Value));

                                    if (elm.Attribute("排名") != null)
                                        sser.AddYearRank(elm.Attribute("科目").Value, int.Parse(elm.Attribute("排名").Value));
                                }
                            }
                        }

                        // 類1排
                        if (!string.IsNullOrEmpty(g1Str))
                        {
                            XElement elmC = XElement.Parse(g1Str);
                            foreach (XElement elmR in elmC.Elements("Rating"))
                            foreach (XElement elm in elmR.Elements("Item"))
                            {
                                if (elm.Attribute("科目") != null)
                                {
                                    if (elm.Attribute("成績人數") != null)
                                        sser.AddGroup1Count(elm.Attribute("科目").Value, int.Parse(elm.Attribute("成績人數").Value));

                                    if (elm.Attribute("排名") != null)
                                        sser.AddGroup1Rank(elm.Attribute("科目").Value, int.Parse(elm.Attribute("排名").Value));
                                }
                            }
                        } 
                    }
                    catch (Exception ex) { }

                    retValue[sid].Add(sser);
                }
            }
            return retValue;
        }
    }
}
