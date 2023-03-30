using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using ClassSemesterScoreReportFixed_SH.DAO;

namespace ClassSemesterScoreReportFixed_SH
{
    public class Global
    {
        #region 設定檔記錄用

        /// <summary>
        /// UDT TableName
        /// </summary>
        public const string _UDTTableName = "ischool.ClassSemesterScoreReportFixed_SH.Configure";

        public static string _ProjectName = "高中班級學期成績單";

        public static string _DefaultConfTypeName = "預設設定檔";

        public static string _UserConfTypeName = "使用者選擇設定檔";

        /// <summary>
        /// 設定檔預設名稱
        /// </summary>
        /// <returns></returns>
        public static List<string> DefaultConfigNameList()
        {
            List<string> retVal = new List<string>();
            retVal.Add("班級學期成績單13科");
            retVal.Add("班級學期成績單24科");
          
            return retVal;
        }

        #endregion

        /// <summary>
        /// 學生各項學期成績
        /// </summary>
        public static Dictionary<string, Dictionary<string, decimal>> _TempStudentSemesScoreDict = new Dictionary<string, Dictionary<string, decimal>>();


        /// <summary>
        /// 班排名
        /// </summary>
        public static Dictionary<string, Dictionary<string, ScoreItem>> _TempClassRankDict = new Dictionary<string, Dictionary<string, ScoreItem>>();
        /// <summary>

        /// <summary>
        /// 科排名
        /// </summary>
        public static Dictionary<string, Dictionary<string, ScoreItem>> _TempDeptRankDict = new Dictionary<string, Dictionary<string, ScoreItem>>();

        /// <summary>
        /// 年排名
        /// </summary>
        public static Dictionary<string, Dictionary<string, ScoreItem>> _TempGradeYearRankDict = new Dictionary<string, Dictionary<string, ScoreItem>>();


        /// <summary>
        /// 類1排名
        /// </summary>
        public static Dictionary<string, Dictionary<string, ScoreItem>> _TempGroup1RankDict = new Dictionary<string, Dictionary<string, ScoreItem>>();
        
        /// 班排名(科目)
        /// </summary>
        public static Dictionary<string, Dictionary<string, ScoreItem>> _TempSubjClassRankDict = new Dictionary<string, Dictionary<string, ScoreItem>>();
        

        /// <summary>
        /// 科排名(科目)
        /// </summary>
        public static Dictionary<string, Dictionary<string, ScoreItem>> _TempSubjDeptRankDict = new Dictionary<string, Dictionary<string, ScoreItem>>();


        /// <summary>
        /// 年排名(科目)
        /// </summary>
        public static Dictionary<string, Dictionary<string, ScoreItem>> _TempSubjGradeYearRankDict = new Dictionary<string, Dictionary<string, ScoreItem>>();

        /// <summary>
        /// 類1排名(科目)
        /// </summary>
        public static Dictionary<string, Dictionary<string, ScoreItem>> _TempSubjGroup1RankDict = new Dictionary<string, Dictionary<string, ScoreItem>>();

        /// <summary>
        /// 學生學分
        /// </summary>
        public static Dictionary<string, DAO.StudCredit> _StudCreditDict = new Dictionary<string, StudCredit>();


        /// <summary>
        /// 班級 科目習修人數 
        /// Key 科目+級別 int :習修人數
        /// </summary>
        public static Dictionary<string, Dictionary<string, int>> SubjectStudiedCountDic = new Dictionary<string, Dictionary<string, int>>();

        /// <summary>
        /// 班級 科目取得學分人數 
        /// Key 科目+級別 int :取得學分人數
        /// </summary>
        public static Dictionary<string, Dictionary<string, int>> SubjectPassCountDic = new Dictionary<string, Dictionary<string, int>>();
    }
}
