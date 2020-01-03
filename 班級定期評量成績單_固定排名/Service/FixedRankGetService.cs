using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using 班級定期評量成績單_固定排名.Model;

namespace 班級定期評量成績單_固定排名.Service
{
    class FixedRankGetService
    {
        /// <summary>
        /// 學年度
        /// </summary>
        public string SchoolYear { get; set; }

        /// <summary>
        /// 學期
        /// </summary>
        public string Semester { get; set; }

        /// <summary>
        /// 試別ID
        /// </summary>
        public string ExamID { get; set; }


        /// <summary>
        ///  某班級的固定評量資訊
        /// </summary>
       // private Dictionary<string, StudentFixedRankInfo> DicStuFixedRankInfo { get; set; }

        /// <summary>
        /// 多個班級固定評量 
        /// </summary>
        private Dictionary<string, Dictionary<string, StudentFixedRankInfo>> DicStuFixedRankInfoMany { get; set; }


        private QueryHelper _Qp;


        /// <summary>
        /// 建構子
        /// </summary>
        public FixedRankGetService(string schoolYear, string semester, string examID)
        {
            //   DicStuFixedRankInfo = new Dictionary<string, StudentFixedRankInfo>();
            DicStuFixedRankInfoMany = new Dictionary<string, Dictionary<string, StudentFixedRankInfo>>();
            this.SchoolYear = schoolYear;
            this.Semester = semester;
            this.ExamID = examID;
            this._Qp = new QueryHelper();
        }


        /// <summary>
        /// Load 所有班級的資料 by 班級
        /// </summary>
        /// <param name="classID"></param>
        public Dictionary<string, StudentFixedRankInfo> GetFixedRankInfo(string classID)
        {
            Dictionary<string, StudentFixedRankInfo> dicStuFixedRankInfo = new Dictionary<string, StudentFixedRankInfo>();

            // this.DicStuFixedRankInfo.Clear();

            GetFixedRankByClass(classID, "定期評量/科目成績", ref dicStuFixedRankInfo); //科目成績排名(班、科、年)

            GetFixedRankByClass(classID, "定期評量/總計成績", ref dicStuFixedRankInfo); //總分(班、科、年排)

            return dicStuFixedRankInfo;
        }


        /// <summary>
        /// 取得多個班級
        /// </summary>
        /// <param name="classID"></param>
        /// <returns></returns>
        public Dictionary<string, Dictionary<string, StudentFixedRankInfo>> GetFixedRankManyClass(List<string> classID)
        {
            if (classID != null)
            {
                foreach (string claID in classID)
                {
                    Dictionary<string, StudentFixedRankInfo> oneClassFixRankInfo = GetFixedRankInfo(claID);
                    this.DicStuFixedRankInfoMany.Add(claID, oneClassFixRankInfo);
                }
            }

            return this.DicStuFixedRankInfoMany;
        }

        /// <summary>
        /// 取得某學期學生排名 by class
        /// </summary>
        /// <param name="class_id"></param>
        /// <param name="itemType">定期評量/科目成績、定期評量/總計成績</param>
        private void GetFixedRankByClass(string classID, string itemType, ref Dictionary<string, StudentFixedRankInfo> dicFixRankInfo)
        {
            string sql = @"
WITH  sp_rank_matrix  AS
(
	SELECT 	
		*
	FROM   
		rank_matrix 
	WHERE 
		school_year = {0}
		AND semester = {1}
		AND item_type = '{4}'	
		AND is_alive = true 
		--AND( rank_name IN (SELECT  class_name FROM  class WHERE id = {3} ) OR  rank_type LIKE '%類別%') --班級名稱
		AND ref_exam_id = {2} --某次定期評量

)SELECT 
		sp_rank_matrix.*
		, rank_detail.ref_student_id 
		, rank_detail.rank
FROM  sp_rank_matrix
	INNER JOIN   ( SELECT * FROM rank_detail  WHERE ref_student_id IN  (SELECT id FROM student WHERE ref_class_id ={3}) )AS rank_detail    
    	ON rank_detail.ref_matrix_id = sp_rank_matrix.id 
";
            sql = string.Format(sql, this.SchoolYear, this.Semester, this.ExamID, classID, itemType);

            DataTable dt = _Qp.Select(sql);

            foreach (DataRow dr in dt.Rows)
            {
                string stuID = "" + dr["ref_student_id"];
                string rankType = "" + dr["rank_type"];  //班排名、科排名、年排名
                string rankName = "" + dr["rank_name"];  //普101、普通科、1年級
                string itemName = "" + dr["item_name"];  //加權平均(總分)、算術平均(總分)
                string rank = "" + dr["rank"];
                string matrixCount = "" + dr["matrix_count"]; //母體數量。
                string gradeYear = "" + dr["grade_year"]; //年級 

                //看是 科目成績 OR 總計成績
                string itemtype = itemType.Split('/')[1];
                if (!dicFixRankInfo.ContainsKey(stuID))
                {
                    //1. new stuFixRankInfo
                    StudentFixedRankInfo stuFixRankInfo = new StudentFixedRankInfo(stuID);
                    dicFixRankInfo.Add(stuID, stuFixRankInfo);
                }
                Dictionary<string, RankInfo> rankInfo;
                string key; //Dictionary 用
                if (itemType == "定期評量/科目成績")
                {
                    rankInfo = dicFixRankInfo[stuID].DicSubjectFixRank;
                    //Dictionary Key 

                }
                else
                {//定期評量/總計成績

                    rankInfo = dicFixRankInfo[stuID].DicSubjectTotalFixRank;
                }

                //取用處邏輯// "科目成績" + "^^^" + "班排名" + "grade" + stuRec.RefClass.GradeYear + "^^^" + stuRec.RefClass.ClassName + "^^^" + sceTakeRecord.Subject;//@@@@@ 取固定排名的key值
                key = itemtype + "^^^" + rankType+"^^^" + "grade" + gradeYear + "^^^" + rankName + "^^^" + itemName;

                if (!rankInfo.ContainsKey(key))
                {
                    rankInfo.Add(key, new RankInfo(itemName, rank, matrixCount));
                }
                else
                {
                    // 有錯誤
                }
            }
        }
    }
}
