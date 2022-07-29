using FISCA.Data;
using SH_ClassYearScoreReport.DAO;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SH_ClassYearScoreReport
{
    public class Global
    {
        #region 設定檔記錄用

        /// <summary>
        /// UDT TableName
        /// </summary>
        public const string _UDTTableName = "ischool.ClassYearScoreReport_SH.Configure";

        public static string _ProjectName = "高中班級學年成績單";

        public static string _DefaultConfTypeName = "預設設定檔";

        public static string _UserConfTypeName = "使用者選擇設定檔";

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


        public static List<string> _SubjectListOrder = GetSubjectOrder();


        /// <summary>
        /// 學生學分
        /// </summary>
        //public static Dictionary<string, DAO.StudCredit> _StudCreditDict = new Dictionary<string, StudCredit>();

        internal static void CreateFieldTemplate()
        {
            #region 產生欄位表

            Aspose.Words.Document doc = new Aspose.Words.Document(new System.IO.MemoryStream(Properties.Resources.Template));
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
            int maxSubjectNum = 35;  //40
            int maxStuNum = 60;  //60

            builder.Font.Size = 8;

            #region 基本欄位
            builder.Writeln("基本欄位");
            builder.StartTable();
            foreach (string field in new string[] { "學年度", "學校名稱", "學校地址", "學校電話", "科別名稱", "班級", "班導師" })//, "類別排名1", "類別排名2", "學期成績排名採計成績欄位"
            {
                builder.InsertCell();
                builder.Write(field);
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + field + " \\* MERGEFORMAT ", "«" + field + "»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 科目成績
            builder.Writeln("科目成績");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("科目名稱");
            for (int i = 1; i <= maxSubjectNum; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目名稱" + i + " \\* MERGEFORMAT ", "«科目名稱" + i + "»");
            }
            builder.EndRow();

            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.Write("學分數");
            //for (int i = 1; i <= maxSubjectNum; i++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 學分數" + i + " \\* MERGEFORMAT ", "«C" + i + "»");
            //}
            //builder.EndRow();

            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 科目成績" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«S" + i + "»");
                }
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 科目學年學分
            builder.Writeln("科目學年學分");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("科目名稱");
            for (int i = 1; i <= maxSubjectNum; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 科目名稱" + i + " \\* MERGEFORMAT ", "«科目名稱" + i + "»");
            }
            builder.EndRow();

            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 科目學年學分" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«C" + i + "»");
                }
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 科目成績排名
            foreach (string key in new string[] { "班", "科", "年" })//, "類別1", "類別2" 
            {
                builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
                builder.Writeln("科目成績" + key + "排名");
                builder.StartTable();
                builder.InsertCell();
                builder.InsertCell();
                builder.InsertCell();
                builder.Write("科目名稱");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 科目名稱" + i + " \\* MERGEFORMAT ", "«科目名稱" + i + "»");
                }
                builder.EndRow();

                for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 學號" + stuIndex + "\\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 座號" + stuIndex + "\\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 姓名" + stuIndex + "\\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                    for (int i = 1; i <= maxSubjectNum; i++)
                    {
                        builder.InsertCell();
                        builder.InsertField("MERGEFIELD " + key + "排名" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«RS»");
                        builder.InsertField("MERGEFIELD " + key + "排名母數" + stuIndex + "-" + i + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                    }
                    builder.EndRow();
                }

                #region 科目成績 五標
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("頂標");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_avg_top_25" + " \\* MERGEFORMAT ", "«T" + i + "»");
                //}
                //builder.EndRow();

                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("高標");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_avg_top_50" + " \\* MERGEFORMAT ", "«T" + i + "»");
                //}
                //builder.EndRow();

                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("均標");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_avg" + " \\* MERGEFORMAT ", "«T" + i + "»");
                //}
                //builder.EndRow();

                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("低標");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_avg_bottom_50" + " \\* MERGEFORMAT ", "«T" + i + "»");
                //}
                //builder.EndRow();

                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("底標");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_avg_bottom_25" + " \\* MERGEFORMAT ", "«T" + i + "»");
                //}
                //builder.EndRow();

                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("新頂標");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_pr_88" + " \\* MERGEFORMAT ", "«P" + i + "»");
                //}
                //builder.EndRow();

                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("新前標");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_pr_75" + " \\* MERGEFORMAT ", "«P" + i + "»");
                //}
                //builder.EndRow();

                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("新均標");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_pr_50" + " \\* MERGEFORMAT ", "«P" + i + "»");
                //}
                //builder.EndRow();

                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("新後標");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_pr_25" + " \\* MERGEFORMAT ", "«P" + i + "»");
                //}
                //builder.EndRow();

                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("新底標");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_pr_12" + " \\* MERGEFORMAT ", "«P" + i + "»");
                //}
                //builder.EndRow();

                //builder.InsertCell();
                //builder.InsertCell();
                //builder.InsertCell();
                //builder.Write("標準差");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                //    builder.InsertCell();
                //    builder.InsertField("MERGEFIELD 科目成績" + i + "_" + key + "排名_std_dev_pop" + " \\* MERGEFORMAT ", "«STD" + i + "»");
                //}
                //builder.EndRow();
                #endregion

                builder.EndTable();
            }
            #endregion

            #region 單科班級總計成績 ?? 這裡應該只有平均才合理，先註解
            //builder.Writeln("單科班級總計成績");
            //builder.StartTable();

            //builder.InsertCell();
            //builder.Write("科目名稱");
            //for (int i = 1; i <= maxSubjectNum; i++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 科目名稱" + i + " \\* MERGEFORMAT ", "«科目名稱" + i + "»");
            //}
            //builder.EndRow();

            //builder.InsertCell();
            //builder.Write("加權平均");
            //for (int i = 1; i <= maxSubjectNum; i++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 科目加權平均" + i + " \\* MERGEFORMAT ", "«W" + i + "»");
            //}
            //builder.EndRow();

            //builder.InsertCell();
            //builder.Write("平均");
            //for (int i = 1; i <= maxSubjectNum; i++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 科目平均" + i + " \\* MERGEFORMAT ", "«A" + i + "»");
            //}
            //builder.EndRow();

            //builder.InsertCell();
            //builder.Write("加權總分");
            //for (int i = 1; i <= maxSubjectNum; i++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 科目加權總分" + i + " \\* MERGEFORMAT ", "«S" + i + "»");
            //}
            //builder.EndRow();

            //builder.InsertCell();
            //builder.Write("總分");
            //for (int i = 1; i <= maxSubjectNum; i++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 科目總分" + i + " \\* MERGEFORMAT ", "«T" + i + "»");
            //}
            //builder.EndRow();

            //builder.EndTable();
            #endregion

            #region 學生分項成績
            builder.Writeln("學生分項成績");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("分項成績");
            //for (int i = 1; i <= maxSubjectNum; i++)
            //{
                builder.InsertCell();
                builder.Write("學業成績");
                //builder.InsertField("MERGEFIELD 學年學業成績" + i + " \\* MERGEFORMAT ", "«學業成績" + i + "»");
                builder.InsertCell();
                builder.Write("專業科目成績");
                //builder.InsertField("MERGEFIELD 學年專業成績" + i + " \\* MERGEFORMAT ", "«專業成績" + i + "»");
                builder.InsertCell();
                builder.Write("實習科目成績");
                //builder.InsertField("MERGEFIELD 學年實習成績" + i + " \\* MERGEFORMAT ", "«實習成績" + i + "»");
            //}
            builder.EndRow();



            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                //for (int i = 1; i <= maxSubjectNum; i++)
                //{
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學年學業成績" + stuIndex + " \\* MERGEFORMAT ", "«學業成績" + stuIndex + "»");

                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學年專業科目成績" + stuIndex + " \\* MERGEFORMAT ", "«專業科目成績" + stuIndex + "»");

                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學年實習科目成績" + stuIndex + " \\* MERGEFORMAT ", "«實習科目成績" + stuIndex + "»");
                //builder.InsertField("MERGEFIELD 科目成績" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«S" + i + "»");

                //}
                builder.EndRow();
            }
            builder.EndTable();
            #endregion


            #region 科目成績(原始)
            //builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            //builder.Writeln("科目成績(原始)");
            //builder.StartTable();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.Write("科目名稱(原始)");
            //for (int i = 1; i <= maxSubjectNum; i++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 科目名稱" + i + " \\* MERGEFORMAT ", "«科目名稱" + i + "»");
            //}
            //builder.EndRow();

            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.Write("學分數");
            //for (int i = 1; i <= maxSubjectNum; i++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 學分數" + i + " \\* MERGEFORMAT ", "«C" + i + "»");
            //}
            //builder.EndRow();

            //for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«S" + i + "»");
            //    }
            //    builder.EndRow();
            //}
            //builder.EndTable();
            #endregion

            #region 科目成績排名(原始)
            //foreach (string key in new string[] { "班", "科", "全校", "類別1", "類別2" })
            //{
            //    builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            //    builder.Writeln("科目成績" + key + "排名(原始)");
            //    builder.StartTable();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("科目名稱");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目名稱" + i + " \\* MERGEFORMAT ", "«科目名稱" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("學分數");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 學分數" + i + " \\* MERGEFORMAT ", "«C" + i + "»");
            //    }
            //    builder.EndRow();

            //    for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 學號" + stuIndex + "\\* MERGEFORMAT ", "«學號" + stuIndex + "»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 座號" + stuIndex + "\\* MERGEFORMAT ", "«座號" + stuIndex + "»");
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 姓名" + stuIndex + "\\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
            //        for (int i = 1; i <= maxSubjectNum; i++)
            //        {
            //            builder.InsertCell();
            //            builder.InsertField("MERGEFIELD " + key + "排名(原始)" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«RS»");
            //            builder.InsertField("MERGEFIELD " + key + "排名(原始)母數" + stuIndex + "-" + i + " \\b /  \\* MERGEFORMAT ", "/«TS»");
            //        }
            //        builder.EndRow();
            //    }
            //    // 2022/1/7 科目成績(原始) 五標 Cynthia
            //    #region 科目成績(原始) 五標
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("頂標");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_avg_top_25" + " \\* MERGEFORMAT ", "«T" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("高標");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_avg_top_50" + " \\* MERGEFORMAT ", "«T" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("均標");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_avg" + " \\* MERGEFORMAT ", "«T" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("低標");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_avg_bottom_50" + " \\* MERGEFORMAT ", "«T" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("底標");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_avg_bottom_25" + " \\* MERGEFORMAT ", "«T" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("新頂標");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_pr_88" + " \\* MERGEFORMAT ", "«P" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("新前標");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_pr_75" + " \\* MERGEFORMAT ", "«P" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("新均標");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_pr_50" + " \\* MERGEFORMAT ", "«P" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("新後標");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_pr_25" + " \\* MERGEFORMAT ", "«P" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("新底標");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_pr_12" + " \\* MERGEFORMAT ", "«P" + i + "»");
            //    }
            //    builder.EndRow();

            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.InsertCell();
            //    builder.Write("標準差");
            //    for (int i = 1; i <= maxSubjectNum; i++)
            //    {
            //        builder.InsertCell();
            //        builder.InsertField("MERGEFIELD 科目成績(原始)" + i + "_" + key + "排名_std_dev_pop" + " \\* MERGEFORMAT ", "«STD" + i + "»");
            //    }
            //    builder.EndRow();
            //    #endregion
            //    builder.EndTable();
            //}
            #endregion

            #region 加權總分、加權平均及排名
            //builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            //builder.Writeln("加權平均及排名");
            //builder.StartTable();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();

            //builder.Write("加權平均");
            //builder.InsertCell();
            //builder.Write("加權平均班排名");
            //builder.InsertCell();
            //builder.Write("加權平均科排名");
            //builder.InsertCell();
            //builder.Write("加權平均校排名");
            //builder.EndRow();
            //for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 加權平均" + stuIndex + " \\* MERGEFORMAT ", "«平均»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 加權平均班排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
            //    builder.InsertField("MERGEFIELD 加權平均班排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 加權平均科排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
            //    builder.InsertField("MERGEFIELD 加權平均科排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 加權平均全校排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
            //    builder.InsertField("MERGEFIELD 加權平均全校排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
            //    builder.EndRow();
            //}
            //builder.EndTable();
            #endregion

            #region 加權總分、加權平均及排名(原始)
            //builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            //builder.Writeln("加權平均及排名(原始)");
            //builder.StartTable();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();

            //builder.Write("加權平均(原始)");
            //builder.InsertCell();
            //builder.Write("加權平均班排名(原始)");
            //builder.InsertCell();
            //builder.Write("加權平均科排名(原始)");
            //builder.InsertCell();
            //builder.Write("加權平均校排名(原始)");
            //builder.EndRow();
            //for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 加權平均(原始)" + stuIndex + " \\* MERGEFORMAT ", "«平均»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 加權平均班排名(原始)" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
            //    builder.InsertField("MERGEFIELD 加權平均班排名(原始)母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 加權平均科排名(原始)" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
            //    builder.InsertField("MERGEFIELD 加權平均科排名(原始)母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 加權平均全校排名(原始)" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
            //    builder.InsertField("MERGEFIELD 加權平均全校排名(原始)母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
            //    builder.EndRow();
            //}
            //builder.EndTable();
            #endregion



            #region 應得與實得學分

            //builder.Writeln("應得與實得學分");
            //builder.StartTable();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.Write("應得學分");
            //builder.InsertCell();
            //builder.Write("實得學分");
            //builder.InsertCell();
            //builder.Write("應得學分累計");
            //builder.InsertCell();
            //builder.Write("實得學分累計");
            //builder.EndRow();
            //for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 應得學分" + stuIndex + " \\* MERGEFORMAT ", "«SC»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 實得學分" + stuIndex + " \\* MERGEFORMAT ", "«GC»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 應得學分累計" + stuIndex + " \\* MERGEFORMAT ", "«ST»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 實得學分累計" + stuIndex + " \\* MERGEFORMAT ", "«GT»");

            //    builder.EndRow();
            //}
            //builder.EndTable();

            #endregion

            #region 類別1排名
            //builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            //builder.Writeln("類別1排名");
            //builder.StartTable();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();

            ////builder.Write("加權平均");
            ////builder.InsertCell();
            //builder.Write("加權平均排名");
            //builder.InsertCell();
            //builder.Write("加權平均排名(原始)");
            //builder.EndRow();
            //for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

            //    //builder.InsertCell();
            //    //builder.InsertField("MERGEFIELD 類別1加權平均" + stuIndex + " \\* MERGEFORMAT ", "«類1加均»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 類別1加權平均排名" + stuIndex + " \\* MERGEFORMAT ", "«RP»");
            //    builder.InsertField("MERGEFIELD 類別1加權平均排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TP»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 類別1加權平均排名(原始)" + stuIndex + " \\* MERGEFORMAT ", "«RP»");
            //    builder.InsertField("MERGEFIELD 類別1加權平均排名(原始)母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TP»");
            //    builder.EndRow();
            //}
            //builder.EndTable();
            #endregion

            #region 類別2排名
            //builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            //builder.Writeln("類別2排名");
            //builder.StartTable();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();
            //builder.InsertCell();

            ////builder.Write("加權平均");
            ////builder.InsertCell();
            //builder.Write("加權平均排名");
            //builder.InsertCell();
            //builder.Write("加權平均排名(原始)");
            //builder.EndRow();
            //for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            //{
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

            //    //builder.InsertCell();
            //    //builder.InsertField("MERGEFIELD 類別2加權平均" + stuIndex + " \\* MERGEFORMAT ", "«類2加均»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 類別2加權平均排名" + stuIndex + " \\* MERGEFORMAT ", "«RP»");
            //    builder.InsertField("MERGEFIELD 類別2加權平均排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TP»");
            //    builder.InsertCell();
            //    builder.InsertField("MERGEFIELD 類別2加權平均排名(原始)" + stuIndex + " \\* MERGEFORMAT ", "«RP»");
            //    builder.InsertField("MERGEFIELD 類別2加權平均排名(原始)母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TP»");
            //    builder.EndRow();
            //}
            //builder.EndTable();
            #endregion

            #region 刪除
            //List<string> r1 = new List<string>();
            //List<string> r2 = new List<string>();

            //r1.Add("pr");
            //r1.Add("percentile");
            //r1.Add("avg_top_25");
            //r1.Add("avg_top_50");
            //r1.Add("avg");
            //r1.Add("avg_bottom_50");
            //r1.Add("avg_bottom_25");
            //r1.Add("pr_88");
            //r1.Add("pr_75");
            //r1.Add("pr_50");
            //r1.Add("pr_25");
            //r1.Add("pr_12");
            //r1.Add("std_dev_pop");
            //r2.Add("level_gte100");
            //r2.Add("level_90");
            //r2.Add("level_80");
            //r2.Add("level_70");
            //r2.Add("level_60");
            //r2.Add("level_50");
            //r2.Add("level_40");
            //r2.Add("level_30");
            //r2.Add("level_20");
            //r2.Add("level_10");
            //r2.Add("level_lt10");

            //List<string> ta1 = new List<string>();
            //ta1.Add("學業");
            //ta1.Add("學業(原始)");

            //List<string> ta2 = new List<string>();
            //ta2.Add("班排名");
            //ta2.Add("科排名");
            //ta2.Add("全校排名");
            //ta2.Add("類別1");
            //ta2.Add("類別2");

            //builder.InsertBreak(Aspose.Words.BreakType.PageBreak);

            //foreach (string s1 in ta1)
            //{
            //    builder.Writeln();
            //    builder.Writeln(s1 + " 百分比、五標");

            //    builder.StartTable();

            //    builder.InsertCell();
            //    builder.Write("名稱_序號");
            //    builder.InsertCell();
            //    builder.Write("PR");
            //    builder.InsertCell();
            //    builder.Write("百分比");
            //    builder.InsertCell();
            //    builder.Write("頂標");
            //    builder.InsertCell();
            //    builder.Write("高標");
            //    builder.InsertCell();
            //    builder.Write("均標");
            //    builder.InsertCell();
            //    builder.Write("低標");
            //    builder.InsertCell();
            //    builder.Write("底標");
            //    builder.InsertCell();
            //    builder.Write("新頂標");
            //    builder.InsertCell();
            //    builder.Write("新前標");
            //    builder.InsertCell();
            //    builder.Write("新均標");
            //    builder.InsertCell();
            //    builder.Write("新後標");
            //    builder.InsertCell();
            //    builder.Write("新底標");
            //    builder.InsertCell();
            //    builder.Write("標準差");
            //    builder.EndRow();
            //    foreach (string s2 in ta2)
            //    {

            //        for (int i = 1; i <= maxStuNum; i++)
            //        {

            //            builder.InsertCell();
            //            builder.Write(s2 + "_" + i);
            //            foreach (string s3 in r1)
            //            {
            //                builder.InsertCell();
            //                string nn = s1 + s2 + "_" + s3 + "_" + i;
            //                builder.InsertField("MERGEFIELD " + nn + " \\* MERGEFORMAT ", "«R»");
            //            }
            //            builder.EndRow();
            //        }

            //    }
            //    builder.EndTable();

            //}

            //foreach (string s1 in ta1)
            //{
            //    builder.Writeln();
            //    builder.Writeln(s1 + " 組距");

            //    builder.StartTable();

            //    builder.InsertCell();
            //    builder.Write("名稱_序號");
            //    builder.InsertCell();
            //    builder.Write("100以上");
            //    builder.InsertCell();
            //    builder.Write("90以上小於100");
            //    builder.InsertCell();
            //    builder.Write("80以上小於90");
            //    builder.InsertCell();
            //    builder.Write("70以上小於80");
            //    builder.InsertCell();
            //    builder.Write("60以上小於70");
            //    builder.InsertCell();
            //    builder.Write("50以上小於60");
            //    builder.InsertCell();
            //    builder.Write("40以上小於50");
            //    builder.InsertCell();
            //    builder.Write("30以上小於40");
            //    builder.InsertCell();
            //    builder.Write("20以上小於30");
            //    builder.InsertCell();
            //    builder.Write("10以上小於20");
            //    builder.InsertCell();
            //    builder.Write("10以下");
            //    builder.EndRow();
            //    foreach (string s2 in ta2)
            //    {
            //        for (int i = 1; i <= maxStuNum; i++)
            //        {

            //            builder.InsertCell();
            //            builder.Write(s2 + "_" + i);
            //            foreach (string s3 in r2)
            //            {
            //                builder.InsertCell();
            //                string nn = s1 + s2 + "_" + s3 + "_" + i;
            //                builder.InsertField("MERGEFIELD " + nn + " \\* MERGEFORMAT ", "«R»");
            //            }
            //            builder.EndRow();

            //        }
            //    }
            //    builder.EndTable();
            //    builder.Writeln();

            //}



            //builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            //#region 學業分項 百分比、五標、組距
            //List<string> t1 = new List<string>();
            //t1.Add("學業");
            //t1.Add("學業(原始)");
            //t1.Add("實習科目");
            //t1.Add("專業科目(原始)");
            //t1.Add("專業科目");
            //t1.Add("實習科目(原始)");

            //List<string> t2 = new List<string>();
            //t2.Add("班排名");
            //t2.Add("科排名");
            //t2.Add("年排名");
            //t2.Add("類別1排名");
            //t2.Add("類別2排名");




            //foreach (string s1 in t1)
            //{
            //    builder.Writeln();
            //    builder.Writeln(s1 + " 百分比、五標");

            //    builder.StartTable();
            //    builder.InsertCell();
            //    builder.Write("分類");
            //    builder.InsertCell();
            //    builder.Write("頂標");
            //    builder.InsertCell();
            //    builder.Write("高標");
            //    builder.InsertCell();
            //    builder.Write("均標");
            //    builder.InsertCell();
            //    builder.Write("低標");
            //    builder.InsertCell();
            //    builder.Write("底標");
            //    builder.InsertCell();
            //    builder.Write("新頂標");
            //    builder.InsertCell();
            //    builder.Write("新前標");
            //    builder.InsertCell();
            //    builder.Write("新均標");
            //    builder.InsertCell();
            //    builder.Write("新後標");
            //    builder.InsertCell();
            //    builder.Write("新底標");
            //    builder.InsertCell();
            //    builder.Write("標準差");
            //    builder.EndRow();
            //    foreach (string s2 in t2)
            //    {
            //        builder.InsertCell();
            //        builder.Write(s2);
            //        foreach (string s3 in r1)
            //        {
            //            if (s3 == "pr" || s3 == "percentile")
            //                continue;

            //            builder.InsertCell();
            //            string nn = s1 + "_" + s2 + "_" + s3;
            //            builder.InsertField("MERGEFIELD " + nn + " \\* MERGEFORMAT ", "«R»");
            //        }
            //        builder.EndRow();
            //    }
            //    builder.EndTable();
            //    builder.Writeln();

            //}


            //foreach (string s1 in t1)
            //{
            //    builder.Writeln();
            //    builder.Writeln(s1 + " 組距");

            //    builder.StartTable();

            //    builder.InsertCell();
            //    builder.Write("分類");
            //    builder.InsertCell();
            //    builder.Write("100以上");
            //    builder.InsertCell();
            //    builder.Write("90以上小於100");
            //    builder.InsertCell();
            //    builder.Write("80以上小於90");
            //    builder.InsertCell();
            //    builder.Write("70以上小於80");
            //    builder.InsertCell();
            //    builder.Write("60以上小於70");
            //    builder.InsertCell();
            //    builder.Write("50以上小於60");
            //    builder.InsertCell();
            //    builder.Write("40以上小於50");
            //    builder.InsertCell();
            //    builder.Write("30以上小於40");
            //    builder.InsertCell();
            //    builder.Write("20以上小於30");
            //    builder.InsertCell();
            //    builder.Write("10以上小於20");
            //    builder.InsertCell();
            //    builder.Write("10以下");
            //    builder.EndRow();
            //    foreach (string s2 in t2)
            //    {
            //        builder.InsertCell();
            //        builder.Write(s2);
            //        foreach (string s3 in r2)
            //        {
            //            builder.InsertCell();
            //            string nn = s1 + "_" + s2 + "_" + s3;
            //            builder.InsertField("MERGEFIELD " + nn + " \\* MERGEFORMAT ", "«R»");
            //        }
            //        builder.EndRow();
            //    }
            //    builder.EndTable();
            //    builder.Writeln();

            //}
            //#endregion


            //#region 學生學業 類別1,2名稱 

            ////builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            ////builder.Writeln("學生類別排名1名稱,類別排名2名稱");
            ////builder.StartTable();
            ////builder.InsertCell();
            ////builder.InsertCell();
            ////builder.InsertCell();
            ////builder.InsertCell();
            ////builder.Write("類別排名1名稱");
            ////builder.InsertCell();
            ////builder.Write("類別排名2名稱");
            ////builder.EndRow();

            ////for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            ////{
            ////    builder.InsertCell();
            ////    builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
            ////    builder.InsertCell();
            ////    builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
            ////    builder.InsertCell();
            ////    builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");

            ////    builder.InsertCell();
            ////    builder.InsertField("MERGEFIELD 學生類別排名1名稱" + stuIndex + " \\* MERGEFORMAT ", "«類1_" + stuIndex + "»");
            ////    builder.InsertCell();
            ////    builder.InsertField("MERGEFIELD 學生類別排名2名稱" + stuIndex + " \\* MERGEFORMAT ", "«類2_" + stuIndex + "»");

            ////    builder.EndRow();
            ////}
            ////builder.EndTable();

            //#endregion

            #endregion


            #endregion

            #region 儲存檔案
            string inputReportName = "班級學年成績單合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".docx");

            if (File.Exists(path))
            {
                int i = 1;
                while (true)
                {
                    string newPath = Path.GetDirectoryName(path) + "\\" + Path.GetFileNameWithoutExtension(path) + (i++) + Path.GetExtension(path);
                    if (!File.Exists(newPath))
                    {
                        path = newPath;
                        break;
                    }
                }
            }

            try
            {
                doc.Save(path, Aspose.Words.SaveFormat.Docx);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".docx";
                sd.Filter = "Word檔案 (*.docx)|*.docx|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        doc.Save(path, Aspose.Words.SaveFormat.Doc);
                    }
                    catch
                    {
                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
        }

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
