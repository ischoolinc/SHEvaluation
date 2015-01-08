using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;

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
        
    }
}
