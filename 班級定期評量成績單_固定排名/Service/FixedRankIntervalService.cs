using FISCA.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.Customization.Data;
using System.Data;
using 班級定期評量成績單.Model;

namespace 班級定期評量成績單.Service
{
    /// <summary>
    /// 取得固定排名相關資料使用
    /// </summary>
    class FixedRankIntervalService
    {
        /// <summary>
        /// 取得學年度
        /// </summary>
        public string School_year { get; set; }

        /// <summary>
        /// 學期 
        /// </summary>
        public string Semester { get; set; }

        /// <summary>
        /// 試別
        /// </summary>
        public string ExamID { get; set; }

        /// <summary>
        ///  排名名稱(班級名稱)對應Rank_matrix 
        /// </summary>
        public List<string> ClassName { get; set; }

        /// <summary>
        ///  存放班級資訊 
        /// </summary>
        public List<ClassRecord> ListClasses;

        /// <summary>
        /// 班級對應科別資料 (普通科、資訊科...等等)
        /// 因為傳入班級資料 缺科別名稱
        /// </summary>
        public Dictionary<string, string> DicClassDeptMapping;

        /// <summary>
        /// Dictionary((班級),Dictionary(科目,科目下級距))
        /// </summary>
        public Dictionary<string, Dictionary<string, FixRankIntervalInfo>> DicEachClassSubjectInteval;

        // public Dictionary<string, IntervalInfoByClass> DicInteval;

        private QueryHelper _Qp;


        public FixedRankIntervalService(string schoolYear, string semester, string examID)
        {
            this._Qp = new QueryHelper();
            this.School_year = schoolYear;
            this.Semester = semester;
            this.ExamID = examID;
            this.DicEachClassSubjectInteval = new Dictionary<string, Dictionary<string, FixRankIntervalInfo>>();
            GetDepInfo(); //取得班級科別對照表
        }


        /// <summary>
        /// 載入班級>科別對照表
        /// </summary>
        public void GetDepInfo()
        {
            this.DicClassDeptMapping = new Dictionary<string, string>();
            string sql = @"
 SELECT  
	class.id 
	, class.class_name
	, dept.name AS dept_name
FROM class
	LEFT JOIN dept 
	 	ON dept.id= class.ref_dept_id 
";

            DataTable dt = _Qp.Select(sql);

            foreach (DataRow dr in dt.Rows)
            {
                this.DicClassDeptMapping.Add("" + dr["id"], "" + dr["dept_name"]);
            }

        }




        /// <summary>
        /// 取得該班級之級距資料 及標(頂標、高標等)
        /// 在(班級、科、年級)之每位學生之(科目*n、加權平均、加權總分、平均(算數)、總分(算數))級距(100以上、90分以上、80分以上)有多少人
        /// </summary>
        /// <param name="classInfo">單一班級資訊</param>
        private void GetIntervalByClass(ClassRecord classInfo)
        {
            string className = classInfo.ClassID;

            Dictionary<string, FixRankIntervalInfo> oneClassInterval = new Dictionary<string, FixRankIntervalInfo>();

            //總計成績的資料
            string sql = @"
            WITH   rank_matrix_sp  AS 
            (
	            SELECT
		            *
	            FROM 
		            rank_matrix   
	            WHERE
		            item_type LIKE '%定期評量%' 
			            AND ref_exam_id = {2}
			            AND  school_year ='{0}'
			            AND  semester = '{1}'
                        AND ( rank_name = '{3}' OR rank_name = '{4}' OR rank_name = '{5}' OR rank_type LIKE'%類別%')
                        AND is_alive = true 
            )SELECT 
	           *
             FROM 
               rank_matrix_sp
	
        ";

            sql = string.Format(sql, this.School_year, this.Semester, this.ExamID, classInfo.ClassName, classInfo.GradeYear + "年級", this.DicClassDeptMapping[classInfo.ClassID]);

            DataTable dt = _Qp.Select(sql);

            foreach (DataRow dr in dt.Rows)
            {
                try
                {
                    //取得key值 要用得變數
                    string itemName = "" + dr["item_name"];
                    string rankType = "" + dr["rank_type"];
                    string rankName = "" + dr["rank_name"];
                    string gradeYear = "" + dr["grade_year"];

                    
                    FixRankIntervalInfo intervalInfo = new FixRankIntervalInfo(rankType, rankName, gradeYear, itemName);
                    intervalInfo.Matrix_count = "" + dr["matrix_Count"];
                    intervalInfo.ItemName = "" + dr["item_name"];
                    intervalInfo.RankName = rankName;
                    intervalInfo.RankType = rankType;

                    intervalInfo.Avg_top_25 = "" + dr["avg_top_25"];
                    intervalInfo.Avg_top_50 = "" + dr["avg_top_50"];
                    intervalInfo.Avg = "" + dr["avg"];
                    intervalInfo.Avg_bottom_50 = "" + dr["avg_bottom_50"];
                    intervalInfo.Avg_bottom_25 = "" + dr["avg_bottom_25"];
                    intervalInfo.Level_gte100 = "" + dr["level_gte100"];
                    intervalInfo.Level_90 = "" + dr["level_90"];
                    intervalInfo.Level_80 = "" + dr["level_80"];
                    intervalInfo.Level_70 = "" + dr["level_70"];
                    intervalInfo.Level_60 = "" + dr["level_60"];
                    intervalInfo.Level_50 = "" + dr["level_50"];
                    intervalInfo.Level_40 = "" + dr["level_40"];
                    intervalInfo.Level_30 = "" + dr["level_30"];
                    intervalInfo.Level_20 = "" + dr["level_20"];
                    intervalInfo.Level_10 = "" + dr["level_10"];
                    intervalInfo.Level_lt10 = "" + dr["level_lt10"];

                    //某班級下 級距dictionary 的 key(ex:2年級(rank_name)某科的加權總分{90-100分:6人 ;80:90分:人 ...etc })

                    //取得 總計成績 OR 科目成績
                    string itemType = ("" + dr["item_type"]).Split('/')[1];

                    string key = $"{itemType}^^^{intervalInfo.RankType}^^^grade{intervalInfo.GradeYear}^^^{intervalInfo.RankName}^^^{intervalInfo.ItemName}";

                    oneClassInterval.Add(key, intervalInfo); //裝進各班的組距資訊
                }
                catch (Exception ex)
                {
                    Console.WriteLine("error" + ex);
                }
            }


            if (!this.DicEachClassSubjectInteval.ContainsKey(classInfo.ClassID))
            {
                this.DicEachClassSubjectInteval.Add(classInfo.ClassID, oneClassInterval);
            }
            else
            {
                //已有同樣班級(不太可能發生)
            }
        }

        /// <summary>
        ///  取得所有班級之級距
        ///  班排名、科排名、校排名
        /// </summary>
        public Dictionary<string, Dictionary<string, FixRankIntervalInfo>> GetAllClassInterval(List<ClassRecord> courseList)
        {
            foreach (ClassRecord classInfo in courseList)
            {
                this.GetIntervalByClass(classInfo);
            }

            return DicEachClassSubjectInteval;
        }
    }
}
