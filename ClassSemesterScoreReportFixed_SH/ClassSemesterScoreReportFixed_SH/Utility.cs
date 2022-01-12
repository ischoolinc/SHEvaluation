using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using FISCA.Data;
using System.Data;
using System.IO;

namespace ClassSemesterScoreReportFixed_SH
{
    public class Utility
    {
        public static string ParseFileName(string fileName)
        {
            string name = fileName;

            if (fileName == null)
                throw new ArgumentNullException();

            if (name.Length == 0)
                throw new ArgumentException();

            if (name.Length > 245)
                throw new PathTooLongException();

            foreach (char c in System.IO.Path.GetInvalidFileNameChars())
            {
                name = name.Replace(c, '_');
            }

            return name;
        }

        public static List<string> GetSubjectOrder()
        {
            List<string> result = new List<string>();
            QueryHelper qh = new QueryHelper();
            string sql =
                @"
            SELECT
	            array_to_string(xpath('//Subject/@Chinese', each_period.period), '')::text as subj_chinese_name
	            , array_to_string(xpath('//Subject/@English', each_period.period), '')::text as subj_english_name
	            , row_number() OVER () as order
            FROM (
	            SELECT unnest(xpath('//Content/Subject', xmlparse(content content))) as period
	            FROM list 
	            WHERE name = '科目中英文對照表'
            ) as each_period
";

            DataTable dt = qh.Select(sql);

            foreach (DataRow dr in dt.Rows)
            {
                string subject = dr["subj_chinese_name"].ToString();
                result.Add(subject);
            }
            return result;
        }
    }
}
