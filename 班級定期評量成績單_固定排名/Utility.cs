using FISCA.Data;
using FISCA.Presentation.Controls;
using SmartSchool.Customization.Data;
//using K12.Data;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;

namespace 班級定期評量成績單_固定排名
{
    /// <summary>
    /// 套表列印功能 通常
    /// step 1. 功能變數總表
    /// step 2. 產生DataTable 並填入 Header
    /// step 3. 不再這裡做 因為很多變數
    /// </summary>

    public class Utility
    {
        /// <summary>
        /// 產生功能變數總表 (類別1 ,類別2 動態產生)
        /// </summary>
        internal static void CreateFieldTemplate()
        {
            #region 產生欄位表

            Aspose.Words.Document doc = new Aspose.Words.Document(new System.IO.MemoryStream(Properties.Resources.Template));
            Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
            int maxSubjectNum = 15;
            int maxStuNum = 60;

            builder.Font.Size = 12;
            #region 基本欄位
            builder.Writeln("基本欄位");
            builder.StartTable();


            // 班級為單位下的基本欄位
            List<string> BaseColumns = new List<string>{
                                                    "學年度"
                                                    , "學期"
                                                    , "學校名稱"
                                                    , "學校地址"
                                                    , "學校電話"
                                                    , "校長名稱"
                                                    , "學務主任"
                                                    , "教務主任"
                                                    , "科別名稱"
                                                    , "定期評量"
                                                    , "班級"
                                                    , "班導師"

            };
            // 加入 類別下 分組欄位 ex :  類別1_分組1名稱 類別1_分組2名稱 ..
            BaseColumns.AddRange(GetTagItemNames());






            foreach (string field in BaseColumns)
            {
                builder.InsertCell();
                builder.Write(field);
                builder.InsertCell();
                builder.InsertField("MERGEFIELD " + field + " \\* MERGEFORMAT ", "«" + field + "»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion


            #region 類別1 類學生所屬分組

            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("學生類別_分組");
            builder.StartTable();

            builder.InsertCell();
            builder.Write("學號");
            builder.InsertCell();
            builder.Write("座號");
            builder.InsertCell();
            builder.Write("姓名");
            builder.InsertCell();
            builder.Write("類別1_分組");
            builder.InsertCell();
            builder.Write("類別2_分組");
            builder.EndRow();
            for (int i = 1; i <= maxStuNum; i++)
            {

                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + i + "\\* MERGEFORMAT ", "«學號" + i + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + i + "\\* MERGEFORMAT ", "«座號" + i + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + i + "\\* MERGEFORMAT ", "«姓名" + i + "»");


                foreach (string tagNames in new string[] { "類別1", "類別2" })
                {
                    builder.InsertCell();
                    builder.InsertField($"MERGEFIELD  {tagNames}_分組名稱{i}  \\* MERGEFORMAT ", $"«{tagNames}_分組名稱{i}»");
                }
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
            builder.Write("領域名稱");
            for (int i = 1; i <= maxSubjectNum; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 領域名稱" + i + " \\* MERGEFORMAT ", "«領域名稱" + i + "»");
            }
            builder.EndRow();
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

            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("分項類別");
            for (int i = 1; i <= maxSubjectNum; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 分項類別" + i + " \\* MERGEFORMAT ", "«E" + i + "»");
            }
            builder.EndRow();

            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("學分數");
            for (int i = 1; i <= maxSubjectNum; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學分數" + i + " \\* MERGEFORMAT ", "«C" + i + "»");
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
                    builder.InsertField("MERGEFIELD 校部定/必選修" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«RB/R" + i + "» ");
                    builder.InsertField("MERGEFIELD 科目成績" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«S" + i + "»");
                    builder.InsertField("MERGEFIELD 科目缺考原因" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«R" + i + "»");
                }
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 科目成績排名
            foreach (string key in new string[] { "班", "科", "年", "類別1", "類別2" })
            {

                builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
                builder.Writeln(" " + key + "排名");
                builder.StartTable();


                builder.InsertCell();
                builder.InsertCell();
                builder.InsertCell();
                builder.Write("領域名稱");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 領域名稱" + i + " \\* MERGEFORMAT ", "«領域名稱" + i + "»");
                }
                builder.EndRow();

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

                builder.InsertCell();
                builder.InsertCell();
                builder.InsertCell();
                builder.Write("分項類別");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 分項類別" + i + " \\* MERGEFORMAT ", "«E" + i + "»");
                }
                builder.EndRow();

                builder.InsertCell();
                builder.InsertCell();
                builder.InsertCell();
                builder.Write("學分數");
                for (int i = 1; i <= maxSubjectNum; i++)
                {
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 學分數" + i + " \\* MERGEFORMAT ", "«C" + i + "»");
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
                builder.EndTable();
            }
            #endregion

            #region 前次成績
            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("前次成績(參考成績試別)");
            builder.StartTable();

            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("領域名稱");
            for (int i = 1; i <= maxSubjectNum; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 領域名稱" + i + " \\* MERGEFORMAT ", "«領域名稱" + i + "»");
            }
            builder.EndRow();

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

            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("分項類別");
            for (int i = 1; i <= maxSubjectNum; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 分項類別" + i + " \\* MERGEFORMAT ", "«E" + i + "»");
            }
            builder.EndRow();

            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("學分數");
            for (int i = 1; i <= maxSubjectNum; i++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學分數" + i + " \\* MERGEFORMAT ", "«C" + i + "»");
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
                    builder.InsertField("MERGEFIELD 校部定/必選修" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«RB/R" + i + "» ");
                    builder.InsertField("MERGEFIELD 前次成績" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«S" + i + "»");
                    builder.InsertField("MERGEFIELD 前次科目缺考原因" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«R" + i + "»");

                }
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 總分、平均及排名
            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("總分、平均及排名");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("總分");
            builder.InsertCell();
            builder.Write("總分班排名");
            builder.InsertCell();
            builder.Write("總分科排名");
            builder.InsertCell();
            builder.Write("總分年排名");
            builder.InsertCell();
            builder.Write("平均");
            builder.InsertCell();
            builder.Write("平均班排名");
            builder.InsertCell();
            builder.Write("平均科排名");
            builder.InsertCell();
            builder.Write("平均年排名");
            builder.EndRow();
            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + "\\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 總分" + stuIndex + " \\* MERGEFORMAT ", "«總分»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 總分班排名" + stuIndex + " \\* MERGEFORMAT ", "«RS»");
                builder.InsertField("MERGEFIELD 總分班排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 總分科排名" + stuIndex + " \\* MERGEFORMAT ", "«RS»");
                builder.InsertField("MERGEFIELD 總分科排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 總分年排名" + stuIndex + " \\* MERGEFORMAT ", "«RS»");
                builder.InsertField("MERGEFIELD 總分年排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 平均" + stuIndex + " \\* MERGEFORMAT ", "«平均»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 平均班排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 平均班排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 平均科排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 平均科排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 平均年排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 平均年排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 加權總分、加權平均及排名
            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("加權總分、加權平均及排名");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("加權總分");
            builder.InsertCell();
            builder.Write("加權總分班排名");
            builder.InsertCell();
            builder.Write("加權總分科排名");
            builder.InsertCell();
            builder.Write("加權總分年排名");
            builder.InsertCell();
            builder.Write("加權平均");
            builder.InsertCell();
            builder.Write("加權平均班排名");
            builder.InsertCell();
            builder.Write("加權平均科排名");
            builder.InsertCell();
            builder.Write("加權平均年排名");
            builder.EndRow();
            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權總分" + stuIndex + " \\* MERGEFORMAT ", "«總分»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權總分班排名" + stuIndex + " \\* MERGEFORMAT ", "«RS»");
                builder.InsertField("MERGEFIELD 加權總分班排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權總分科排名" + stuIndex + " \\* MERGEFORMAT ", "«RS»");
                builder.InsertField("MERGEFIELD 加權總分科排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權總分年排名" + stuIndex + " \\* MERGEFORMAT ", "«RS»");
                builder.InsertField("MERGEFIELD 加權總分年排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均" + stuIndex + " \\* MERGEFORMAT ", "«平均»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均班排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 加權平均班排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均科排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 加權平均科排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 加權平均年排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 加權平均年排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 類別1排名
            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("類別1排名");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("  ");
            builder.InsertCell();
            builder.Write("總分排名");
            builder.InsertCell();
            builder.Write("  ");
            builder.InsertCell();
            builder.Write("平均排名");
            builder.InsertCell();
            builder.Write("  ");
            builder.InsertCell();
            builder.Write("加權總分排名");
            builder.InsertCell();
            builder.Write("  ");
            builder.InsertCell();
            builder.Write("加權平均排名");
            builder.EndRow();
            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別1總分" + stuIndex + " \\* MERGEFORMAT ", "«類1總»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別1總分排名" + stuIndex + " \\* MERGEFORMAT ", "«RS»");
                builder.InsertField("MERGEFIELD 類別1總分排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TS»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別1平均" + stuIndex + " \\* MERGEFORMAT ", "«類1均»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別1平均排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 類別1平均排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別1加權總分" + stuIndex + " \\* MERGEFORMAT ", "«類1加總»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別1加權總分排名" + stuIndex + " \\* MERGEFORMAT ", "«RP»");
                builder.InsertField("MERGEFIELD 類別1加權總分排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TP»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別1加權平均" + stuIndex + " \\* MERGEFORMAT ", "«類1加均»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別1加權平均排名" + stuIndex + " \\* MERGEFORMAT ", "«RP»");
                builder.InsertField("MERGEFIELD 類別1加權平均排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TP»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 類別2排名
            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("類別2排名");
            builder.StartTable();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.InsertCell();
            builder.Write("  ");
            builder.InsertCell();
            builder.Write("總分排名");
            builder.InsertCell();
            builder.Write("  ");
            builder.InsertCell();
            builder.Write("平均排名");
            builder.InsertCell();
            builder.Write("  ");
            builder.InsertCell();
            builder.Write("加權總分排名");
            builder.InsertCell();
            builder.Write("");
            builder.InsertCell();
            builder.Write("加權平均排名");
            builder.EndRow();
            for (int stuIndex = 1; stuIndex <= maxStuNum; stuIndex++)
            {
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 學號" + stuIndex + " \\* MERGEFORMAT ", "«學號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 座號" + stuIndex + " \\* MERGEFORMAT ", "«座號" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 姓名" + stuIndex + " \\* MERGEFORMAT ", "«姓名" + stuIndex + "»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別2總分" + stuIndex + " \\* MERGEFORMAT ", "«類1總»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別2總分排名" + stuIndex + " \\* MERGEFORMAT ", "«RS»");
                builder.InsertField("MERGEFIELD 類別2總分排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TS»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別2平均" + stuIndex + " \\* MERGEFORMAT ", "«類1均»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別2平均排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                builder.InsertField("MERGEFIELD 類別2平均排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TA»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別2加權總分" + stuIndex + " \\* MERGEFORMAT ", "«類1加總»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別2加權總分排名" + stuIndex + " \\* MERGEFORMAT ", "«RP»");
                builder.InsertField("MERGEFIELD 類別2加權總分排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TP»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別2加權平均" + stuIndex + " \\* MERGEFORMAT ", "«類1加均»");
                builder.InsertCell();
                builder.InsertField("MERGEFIELD 類別2加權平均排名" + stuIndex + " \\* MERGEFORMAT ", "«RP»");
                builder.InsertField("MERGEFIELD 類別2加權平均排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "«/TP»");
                builder.EndRow();
            }
            builder.EndTable();
            #endregion

            #region 【科目成績】成績分析 

            // 取得類1 tag 相關資訊

            // todo 

            List<string> rankTypes = new List<string>() { "班", "科", "年", };

            List<string> rankTypesTags = new List<string>();

            // 動態產生
            foreach (string RankTagName in new string[] { "類別1", "類別2" })
            {
                for (int i = 1; i < 10; i++) // 產生10個 
                {
                    rankTypesTags.Add($"{RankTagName}_分組{i}");
                }
            }
            rankTypes.AddRange(rankTypesTags);

            foreach (string key in rankTypes)
            {
                builder.InsertBreak(Aspose.Words.BreakType.PageBreak);

                builder.Writeln(key + "成績分析及組距");

                builder.StartTable();

                builder.InsertCell(); builder.Write("領域名稱");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD 領域名稱" + subjectIndex + " \\* MERGEFORMAT ", "«D" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("科目名稱");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD 科目名稱" + subjectIndex + " \\* MERGEFORMAT ", "«N" + subjectIndex + "»");
                }
                builder.EndRow();

                //Jean 增加固定排名邏輯 (五標)

                builder.InsertCell(); builder.Write("頂標");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "頂標" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();

                builder.InsertCell(); builder.Write("高標");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "高標" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();

                builder.InsertCell(); builder.Write("均標");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "均標" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();

                builder.InsertCell(); builder.Write("低標");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "低標" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();


                builder.InsertCell(); builder.Write("底標");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "底標" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }


                builder.EndRow();

                //2021-12
                builder.InsertCell(); builder.Write("新頂標");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "新頂標" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();

                builder.InsertCell(); builder.Write("新前標");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "新前標" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();

                builder.InsertCell(); builder.Write("新均標");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "新均標" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();

                builder.InsertCell(); builder.Write("新後標");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "新後標" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();

                builder.InsertCell(); builder.Write("新底標");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "新底標" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();

                builder.InsertCell(); builder.Write("標準差");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "標準差" + subjectIndex + " \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();

                builder.InsertCell(); builder.Write("100以上");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count100Up \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("90以上");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count90Up \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("90以上小於100");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count90 \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("小於90");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count90Down \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("80以上");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count80Up \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("80以上小於90");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count80 \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("小於80");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count80Down \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("70以上");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count70Up \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("70以上小於80");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count70 \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("小於70");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count70Down \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("60以上");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count60Up \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("60以上小於70");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count60 \\* MERGEFORMAT ", "«C" + subjectIndex + "»");

                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("小於60");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count60Down \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("50以上");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count50Up \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("50以上小於60");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count50 \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("小於50");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count50Down \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("40以上");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count40Up \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("40以上小於50");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count40 \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("小於40");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count40Down \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("30以上");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count30Up \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("30以上小於40");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count30 \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("小於30");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count30Down \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("20以上");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count20Up \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("20以上小於30");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count20 \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("小於20");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count20Down \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("10以上");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count10Up \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("10以上小於20");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count10 \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.InsertCell(); builder.Write("小於10");
                for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距" + subjectIndex + "count10Down \\* MERGEFORMAT ", "«C" + subjectIndex + "»");
                }
                builder.EndRow();
                builder.EndTable();
            }
            #endregion

            #region 加總成績分析

            // todo 
            builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
            builder.Writeln("加總成績分析及組距");
            builder.StartTable();
            builder.InsertCell(); builder.Write("項目");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /*"類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" */})
            {
                builder.InsertCell(); builder.Write(key);
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("頂標");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"/* "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" */})
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "頂標 \\* MERGEFORMAT ", "«C»");
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("高標");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /*"類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "高標 \\* MERGEFORMAT ", "«C»");
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("均標");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"/*"類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "均標 \\* MERGEFORMAT ", "«C»");
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("低標");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" })
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "低標 \\* MERGEFORMAT ", "«C»");
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("底標");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" })
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "底標 \\* MERGEFORMAT ", "«C»");
            }


            builder.EndRow();

            //2021-12
            builder.InsertCell(); builder.Write("新頂標");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" })
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "新頂標 \\* MERGEFORMAT ", "«C»");
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("新前標");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" })
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "新前標 \\* MERGEFORMAT ", "«C»");
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("新均標");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" })
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "新均標 \\* MERGEFORMAT ", "«C»");
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("新後標");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" })
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "新後標 \\* MERGEFORMAT ", "«C»");
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("新底標");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" })
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "新底標 \\* MERGEFORMAT ", "«C»");
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("標準差");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" })
            {
                builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "標準差 \\* MERGEFORMAT ", "«C»");
            }
            builder.EndRow();


            builder.InsertCell(); builder.Write("100以上");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /*, "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count100Up \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("90以上");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" */})
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count90Up \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("90以上小於100");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /*, "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count90 \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("小於90");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count90Down \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("80以上");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /*, "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count80Up \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("80以上小於90");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /*, "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count80 \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("小於80");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count80Down \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("70以上");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" */})
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count70Up \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }

            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("70以上小於80");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /*, "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count70 \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("小於70");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count70Down \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("60以上");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count60Up \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("60以上小於70");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count60 \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("小於60");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count60Down \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("50以上");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count50Up \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("50以上小於60");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count50 \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }

            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("小於50");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"   /*, "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count50Down \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("40以上");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"   /*  , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count40Up \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("40以上小於50");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /*  , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count40 \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("小於40");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /*  , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count40Down \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("30以上");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count30Up \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }

            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("30以上小於40");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"   /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" */})
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count30 \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("小於30");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年" /*  , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" */})
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count30Down \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("20以上");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"   /*, "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" */})
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count20Up \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("20以上小於30");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /*  , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" */})
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count20 \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("小於20");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"  /*, "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count20Down \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("10以上");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"    /*, "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" */})
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count10Up \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("10以上小於20");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"    /*, "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均"*/ })
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count10 \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }
            }
            builder.EndRow();
            builder.InsertCell(); builder.Write("小於10");
            foreach (string key in new string[] { "總分班", "總分科", "總分年", "平均班", "平均科", "平均年", "加權總分班", "加權總分科", "加權總分年", "加權平均班", "加權平均科", "加權平均年"   /* , "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" */})
            {
                if (!key.Contains("總分"))
                {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count10Down \\* MERGEFORMAT ", "«C»");
                }
                else
                {
                    builder.InsertCell();
                }

            }
            builder.EndRow();
            builder.EndTable();


            #endregion



            #region 加總成績分析【類別1 與 類別2】

            foreach (string tagName in new string[] { "類別1", "類別2" })
            {
                string[] TotalScoreTypes = new string[] { "總分", "平均", "加權總分", "加權平均" }; // 總計成績 有4種

                List<string> columnsHeader = new List<string>();
                foreach (string totalScoreType in TotalScoreTypes)
                {
                    for (int i = 1; i < 10; i++) // 產生10個 
                    {
                        columnsHeader.Add($"{totalScoreType}{tagName}_分組{i}");
                    }
                }


                // 【總計成績】 跟平均都會 有的 分析項目
                List<string> allNeeds = new List<string> {
                         "頂標"
                        ,"高標"
                        ,"均標"
                        ,"低標"
                        ,"底標"
                        ,"新頂標"
                        ,"新前標"
                        ,"新均標"
                        ,"新後標"
                        ,"新底標"
                        ,"標準差"
                };

                Dictionary<string, string> rows = new Dictionary<string, string> {
                        { "頂標" ,"頂標" }
                        ,{"高標"                          ,"高標"   }
                        ,{"均標"                          ,"均標"   }
                        ,{"低標"                          ,"低標"   }
                        ,{"底標"                          ,"底標"   }
                        ,{ "新頂標" ,"新頂標" }
                        ,{"新前標"                          ,"新前標"   }
                        ,{"新均標"                          ,"新均標"   }
                        ,{"新後標"                          ,"新後標"   }
                        ,{"新底標"                          ,"新底標"   }
                        ,{"標準差"                          ,"標準差"   }
                        ,{"100以上"                       ,"組距count100Up"   }
                        ,{"90以上"                        ,"組距count90Up"   }
                        ,{"90以上小於100"                 ,"組距count90"   }
                        ,{"小於90"                        ,"組距count90Down"   }
                        ,{"80以上"                        ,"組距count80Up"   }
                        ,{"80以上小於90"                  ,"組距count80"   }
                        ,{"小於80"                        ,"組距count80Down"   }
                        ,{"70以上"                        ,"組距count70Up "   }
                        ,{"70以上小於80"                  , "組距count70"   }
                        ,{"小於70"                        ,"組距count70Down"   }
                        ,{"60以上"                        ,"組距count60Up"   }
                        ,{"60以上小於70"                  ,"組距count60"   }
                        ,{"小於60"                        ,"組距count60Down"   }
                        ,{"50以上"                        ,"組距count50Up"   }
                        ,{"50以上小於60"                  ,"組距count50"   }
                        ,{"小於50"                        ,"組距count50Down"   }
                        ,{"40以上"                        ,"組距count40Up"   }
                        ,{"40以上小於50"                  ,"組距count40"   }
                        ,{"小於40"                        ,"組距count40Down"   }
                        ,{"30以上"                        ,"組距count30Up"   }
                        ,{"30以上小於40"                  ,"組距count30"   }
                        ,{"小於30"                        ,"組距count30Down"   }
                        ,{"20以上"                        ,"組距count20Up"   }
                        ,{"20以上小於30"                  ,"組距count20"   }
                        ,{"小於20"                        ,"組距count20Down"   }
                        ,{"10以上"                        ,"組距count10Up"   }
                        ,{"10以上小於20"                  ,"組距count10"   }
                        ,{"小於10"                        ,"組距count10Down" }
            };


                // todo 總計成績組距

                builder.InsertBreak(Aspose.Words.BreakType.PageBreak);

                // builder.InsertCell(); builder.Write($"【{tagName}】加總成績分析及組距");
                // builder.Writeln("加總成績分析及組距");
                builder.StartTable();

                builder.InsertCell(); builder.Write("項目");
                foreach (string key in columnsHeader)  // 處理 header 
                {

                    builder.InsertCell(); builder.Write(key);

                }
                builder.EndRow();


                foreach (string rowKey in rows.Keys)
                {
                    builder.InsertCell();
                    builder.Write(rowKey);

                    foreach (string key in columnsHeader)
                    {
                        builder.InsertCell();
                        if (key.Contains("總分") && !allNeeds.Contains(rowKey))
                        {
                            continue;
                        }
                        builder.InsertField("MERGEFIELD " + key + rows[rowKey] + " \\* MERGEFORMAT ", "«C»");
                    }
                    builder.EndRow();
                }







                builder.EndTable();

            }

            #endregion

            #endregion

            // todo 



            #region 儲存檔案
            string inputReportName = "班級定期評量成績單(固定排名)合併欄位總表";
            string reportName = inputReportName;

            string path = Path.Combine(System.Windows.Forms.Application.StartupPath, "Reports");
            if (!Directory.Exists(path))
                Directory.CreateDirectory(path);
            path = Path.Combine(path, reportName + ".doc");

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
                doc.Save(path, Aspose.Words.SaveFormat.Doc);
                System.Diagnostics.Process.Start(path);
            }
            catch
            {
                System.Windows.Forms.SaveFileDialog sd = new System.Windows.Forms.SaveFileDialog();
                sd.Title = "另存新檔";
                sd.FileName = reportName + ".doc";
                sd.Filter = "Excel檔案 (*.doc)|*.doc|所有檔案 (*.*)|*.*";
                if (sd.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    try
                    {
                        doc.Save(path, Aspose.Words.SaveFormat.Doc);
                    }
                    catch
                    {
                        MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                        return;
                    }
                }
            }
            #endregion
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

        /// <summary>
        /// 動產生
        /// </summary>
        /// <param name="configure"></param>
        /// <param name="selectClasses"></param>
        /// <returns></returns>
        internal static DataTable CreateMergeDataTabel(Configure configure, List<SmartSchool.Customization.Data.ClassRecord> selectClasses)
        {
            DataTable table = new DataTable();

            #region 所有的合併欄位
            table.Columns.Add("學校名稱");
            table.Columns.Add("學校地址");
            table.Columns.Add("學校電話");
            table.Columns.Add("校長名稱");
            table.Columns.Add("學務主任");
            table.Columns.Add("教務主任");
            table.Columns.Add("科別名稱");
            table.Columns.Add("試別");
            table.Columns.Add("定期評量");
            table.Columns.Add("班級科別名稱");
            table.Columns.Add("班級");
            table.Columns.Add("班導師");
            table.Columns.Add("學年度");
            table.Columns.Add("學期");
            table.Columns.Add("類別排名1");
            table.Columns.Add("類別排名2");


            // 【固定排名】類別 1 類別 2 下 名稱
            foreach (var item in Utility.GetTagItemNames())
            {
                table.Columns.Add(item);

            }
            // 【固定排名】 類別1 類別2
            for (int i = 1; i <= configure.StudentLimit; i++)
            {
                foreach (string tagNames in new string[] { "類別1", "類別2" })
                {
                    table.Columns.Add($"{tagNames}_分組名稱{i}");
                }
            }

            // table.Columns.Add("類別1名稱");
            // table.Columns.Add("類別2名稱");
            //«通訊地址»«通訊地址郵遞區號»«通訊地址內容»
            //«戶籍地址»«戶籍地址郵遞區號»«戶籍地址內容»
            //«監護人»«父親»«母親»«科別名稱»
            for (int subjectIndex = 1; subjectIndex <= configure.SubjectLimit; subjectIndex++)
            {
                table.Columns.Add("領域名稱" + subjectIndex);
                table.Columns.Add("科目名稱" + subjectIndex);
                table.Columns.Add("分項類別" + subjectIndex);
                table.Columns.Add("學分數" + subjectIndex);
            }
            for (int i = 1; i <= configure.StudentLimit; i++)
            {
                table.Columns.Add("座號" + i);
                table.Columns.Add("學號" + i);
                table.Columns.Add("姓名" + i);
            }

            for (int Num = 1; Num <= configure.StudentLimit; Num++)
            {
                for (int subjectIndex = 1; subjectIndex <= configure.SubjectLimit; subjectIndex++)
                {
                    table.Columns.Add("校部定/必選修" + Num + "-" + subjectIndex);
                    table.Columns.Add("前次成績" + Num + "-" + subjectIndex);
                    table.Columns.Add("科目成績" + Num + "-" + subjectIndex);
                    table.Columns.Add("班排名" + Num + "-" + subjectIndex);
                    table.Columns.Add("班排名母數" + Num + "-" + subjectIndex);
                    table.Columns.Add("科排名" + Num + "-" + subjectIndex);
                    table.Columns.Add("科排名母數" + Num + "-" + subjectIndex);
                    table.Columns.Add("類別1排名" + Num + "-" + subjectIndex);
                    table.Columns.Add("類別1排名母數" + Num + "-" + subjectIndex);
                    table.Columns.Add("類別2排名" + Num + "-" + subjectIndex);
                    table.Columns.Add("類別2排名母數" + Num + "-" + subjectIndex);
                    // table.Columns.Add("全校排名" + Num + "-" + subjectIndex);
                    // table.Columns.Add("全校排名母數" + Num + "-" + subjectIndex);
                    table.Columns.Add("年排名" + Num + "-" + subjectIndex);
                    table.Columns.Add("年排名母數" + Num + "-" + subjectIndex);

                    // 因缺考新增
                    table.Columns.Add("前次科目缺考原因" + Num + "-" + subjectIndex);
                    table.Columns.Add("科目缺考原因" + Num + "-" + subjectIndex);
                }
                table.Columns.Add("總分" + Num);
                table.Columns.Add("總分班排名" + Num);
                table.Columns.Add("總分班排名母數" + Num);
                table.Columns.Add("總分科排名" + Num);
                table.Columns.Add("總分科排名母數" + Num);
                table.Columns.Add("總分年排名" + Num);
                table.Columns.Add("總分年排名母數" + Num);
                table.Columns.Add("平均" + Num);
                table.Columns.Add("平均班排名" + Num);
                table.Columns.Add("平均班排名母數" + Num);
                table.Columns.Add("平均科排名" + Num);
                table.Columns.Add("平均科排名母數" + Num);
                table.Columns.Add("平均年排名" + Num);
                table.Columns.Add("平均年排名母數" + Num);
                table.Columns.Add("加權總分" + Num);
                table.Columns.Add("加權總分班排名" + Num);
                table.Columns.Add("加權總分班排名母數" + Num);
                table.Columns.Add("加權總分科排名" + Num);
                table.Columns.Add("加權總分科排名母數" + Num);
                table.Columns.Add("加權總分年排名" + Num);
                table.Columns.Add("加權總分年排名母數" + Num);
                table.Columns.Add("加權平均" + Num);
                table.Columns.Add("加權平均班排名" + Num);
                table.Columns.Add("加權平均班排名母數" + Num);
                table.Columns.Add("加權平均科排名" + Num);
                table.Columns.Add("加權平均科排名母數" + Num);
                table.Columns.Add("加權平均年排名" + Num);
                table.Columns.Add("加權平均年排名母數" + Num);

                table.Columns.Add("類別排名1" + Num);
                table.Columns.Add("類別1總分" + Num);
                table.Columns.Add("類別1總分排名" + Num);
                table.Columns.Add("類別1總分排名母數" + Num);
                table.Columns.Add("類別1平均" + Num);
                table.Columns.Add("類別1平均排名" + Num);
                table.Columns.Add("類別1平均排名母數" + Num);
                table.Columns.Add("類別1加權總分" + Num);
                table.Columns.Add("類別1加權總分排名" + Num);
                table.Columns.Add("類別1加權總分排名母數" + Num);
                table.Columns.Add("類別1加權平均" + Num);
                table.Columns.Add("類別1加權平均排名" + Num);
                table.Columns.Add("類別1加權平均排名母數" + Num);

                table.Columns.Add("類別排名2" + Num);
                table.Columns.Add("類別2總分" + Num);
                table.Columns.Add("類別2總分排名" + Num);
                table.Columns.Add("類別2總分排名母數" + Num);
                table.Columns.Add("類別2平均" + Num);
                table.Columns.Add("類別2平均排名" + Num);
                table.Columns.Add("類別2平均排名母數" + Num);
                table.Columns.Add("類別2加權總分" + Num);
                table.Columns.Add("類別2加權總分排名" + Num);
                table.Columns.Add("類別2加權總分排名母數" + Num);
                table.Columns.Add("類別2加權平均" + Num);
                table.Columns.Add("類別2加權平均排名" + Num);
                table.Columns.Add("類別2加權平均排名母數" + Num);
            }
            #region 瘋狂的組距及分析 【Header】


            #region 【總計成績】【班、科、年、類1( ex:類1_分組1 類1_分組2 )、類2( 類2_分組1、類2_分組2)】各科目組距及分析 

            // 註:類1組距下可能有很多母群 ex:自然組、社會組

            List<string> rankTypes = new List<string> { "班", "科", "年" };

            // 【總計成績】取的 類1 類2 組距 相關資訊 
            foreach (string item in GetTagItemIntervalList())
            {
                table.Columns.Add(item);

            }

            //【科目成績】取得類1 類2組距 相關資訊
            foreach (string item in GetTagItems())
            {
                rankTypes.Add(item);
            }

            #region 【科目成績】【 產生 班 科 校 年 類別1 類別2  之統計值】


            for (int subjectIndex = 1; subjectIndex <= configure.SubjectLimit; subjectIndex++)
            {
                // 改成迴圈
                foreach (string rankType in rankTypes)   //班 科 校 年 類別1 類別2 
                {
                    // 五標
                    table.Columns.Add($"{rankType}頂標{subjectIndex}");
                    table.Columns.Add($"{rankType}高標{subjectIndex}");
                    table.Columns.Add($"{rankType}均標{subjectIndex}");
                    table.Columns.Add($"{rankType}低標{subjectIndex}");
                    table.Columns.Add($"{rankType}底標{subjectIndex}");

                    //新五標
                    table.Columns.Add($"{rankType}新頂標{subjectIndex}");
                    table.Columns.Add($"{rankType}新前標{subjectIndex}");
                    table.Columns.Add($"{rankType}新均標{subjectIndex}");
                    table.Columns.Add($"{rankType}新後標{subjectIndex}");
                    table.Columns.Add($"{rankType}新底標{subjectIndex}");

                    //標準差
                    table.Columns.Add($"{rankType}標準差{subjectIndex}");

                    // 組距

                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count90");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count80");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count70");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count60");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count50");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count40");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count30");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count20");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count10");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count100Up");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count90Up");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count80Up");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count70Up");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count60Up");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count50Up");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count40Up");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count30Up");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count20Up");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count10Up");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count90Down");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count80Down");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count70Down");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count60Down");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count50Down");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count40Down");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count30Down");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count20Down");
                    table.Columns.Add($"{rankType}組距" + subjectIndex + "count10Down");

                }

                #endregion


                //table.Columns.Add("班頂標" + subjectIndex);
                //table.Columns.Add("科頂標" + subjectIndex);
                //table.Columns.Add("年頂標" + subjectIndex);

                //table.Columns.Add("類1頂標" + subjectIndex);
                //table.Columns.Add("類2頂標" + subjectIndex);

                //table.Columns.Add("班高標" + subjectIndex);
                //table.Columns.Add("科高標" + subjectIndex);
                //table.Columns.Add("年高標" + subjectIndex);

                //table.Columns.Add("類1高標" + subjectIndex);
                //table.Columns.Add("類2高標" + subjectIndex);

                //table.Columns.Add("班均標" + subjectIndex);
                //table.Columns.Add("科均標" + subjectIndex);
                //table.Columns.Add("年均標" + subjectIndex);

                //table.Columns.Add("類1均標" + subjectIndex);
                //table.Columns.Add("類2均標" + subjectIndex);

                //table.Columns.Add("班低標" + subjectIndex);
                //table.Columns.Add("科低標" + subjectIndex);
                //table.Columns.Add("年低標" + subjectIndex);

                //table.Columns.Add("類1低標" + subjectIndex);
                //table.Columns.Add("類2低標" + subjectIndex);

                //table.Columns.Add("班底標" + subjectIndex);
                //table.Columns.Add("科底標" + subjectIndex);
                //table.Columns.Add("年底標" + subjectIndex);

                //table.Columns.Add("類1底標" + subjectIndex);
                //table.Columns.Add("類2底標" + subjectIndex);

                // 轉成迴圈



                ////標準差拿掉
                //table.Columns.Add("班組距" + subjectIndex + "count90"); table.Columns.Add("科組距" + subjectIndex + "count90"); table.Columns.Add("年組距" + subjectIndex + "count90"); table.Columns.Add("類1組距" + subjectIndex + "count90"); table.Columns.Add("類2組距" + subjectIndex + "count90");
                //table.Columns.Add("班組距" + subjectIndex + "count80"); table.Columns.Add("科組距" + subjectIndex + "count80"); table.Columns.Add("年組距" + subjectIndex + "count80"); table.Columns.Add("類1組距" + subjectIndex + "count80"); table.Columns.Add("類2組距" + subjectIndex + "count80");
                //table.Columns.Add("班組距" + subjectIndex + "count70"); table.Columns.Add("科組距" + subjectIndex + "count70"); table.Columns.Add("年組距" + subjectIndex + "count70"); table.Columns.Add("類1組距" + subjectIndex + "count70"); table.Columns.Add("類2組距" + subjectIndex + "count70");
                //table.Columns.Add("班組距" + subjectIndex + "count60"); table.Columns.Add("科組距" + subjectIndex + "count60"); table.Columns.Add("年組距" + subjectIndex + "count60"); table.Columns.Add("類1組距" + subjectIndex + "count60"); table.Columns.Add("類2組距" + subjectIndex + "count60");
                //table.Columns.Add("班組距" + subjectIndex + "count50"); table.Columns.Add("科組距" + subjectIndex + "count50"); table.Columns.Add("年組距" + subjectIndex + "count50"); table.Columns.Add("類1組距" + subjectIndex + "count50"); table.Columns.Add("類2組距" + subjectIndex + "count50");
                //table.Columns.Add("班組距" + subjectIndex + "count40"); table.Columns.Add("科組距" + subjectIndex + "count40"); table.Columns.Add("年組距" + subjectIndex + "count40"); table.Columns.Add("類1組距" + subjectIndex + "count40"); table.Columns.Add("類2組距" + subjectIndex + "count40");
                //table.Columns.Add("班組距" + subjectIndex + "count30"); table.Columns.Add("科組距" + subjectIndex + "count30"); table.Columns.Add("年組距" + subjectIndex + "count30"); table.Columns.Add("類1組距" + subjectIndex + "count30"); table.Columns.Add("類2組距" + subjectIndex + "count30");
                //table.Columns.Add("班組距" + subjectIndex + "count20"); table.Columns.Add("科組距" + subjectIndex + "count20"); table.Columns.Add("年組距" + subjectIndex + "count20"); table.Columns.Add("類1組距" + subjectIndex + "count20"); table.Columns.Add("類2組距" + subjectIndex + "count20");
                //table.Columns.Add("班組距" + subjectIndex + "count10"); table.Columns.Add("科組距" + subjectIndex + "count10"); table.Columns.Add("年組距" + subjectIndex + "count10"); table.Columns.Add("類1組距" + subjectIndex + "count10"); table.Columns.Add("類2組距" + subjectIndex + "count10");
                //table.Columns.Add("班組距" + subjectIndex + "count100Up"); table.Columns.Add("科組距" + subjectIndex + "count100Up"); table.Columns.Add("年組距" + subjectIndex + "count100Up"); table.Columns.Add("類1組距" + subjectIndex + "count100Up"); table.Columns.Add("類2組距" + subjectIndex + "count100Up");
                //table.Columns.Add("班組距" + subjectIndex + "count90Up"); table.Columns.Add("科組距" + subjectIndex + "count90Up"); table.Columns.Add("年組距" + subjectIndex + "count90Up"); table.Columns.Add("類1組距" + subjectIndex + "count90Up"); table.Columns.Add("類2組距" + subjectIndex + "count90Up");
                //table.Columns.Add("班組距" + subjectIndex + "count80Up"); table.Columns.Add("科組距" + subjectIndex + "count80Up"); table.Columns.Add("年組距" + subjectIndex + "count80Up"); table.Columns.Add("類1組距" + subjectIndex + "count80Up"); table.Columns.Add("類2組距" + subjectIndex + "count80Up");
                //table.Columns.Add("班組距" + subjectIndex + "count70Up"); table.Columns.Add("科組距" + subjectIndex + "count70Up"); table.Columns.Add("年組距" + subjectIndex + "count70Up"); table.Columns.Add("類1組距" + subjectIndex + "count70Up"); table.Columns.Add("類2組距" + subjectIndex + "count70Up");
                //table.Columns.Add("班組距" + subjectIndex + "count60Up"); table.Columns.Add("科組距" + subjectIndex + "count60Up"); table.Columns.Add("年組距" + subjectIndex + "count60Up"); table.Columns.Add("類1組距" + subjectIndex + "count60Up"); table.Columns.Add("類2組距" + subjectIndex + "count60Up");
                //table.Columns.Add("班組距" + subjectIndex + "count50Up"); table.Columns.Add("科組距" + subjectIndex + "count50Up"); table.Columns.Add("年組距" + subjectIndex + "count50Up"); table.Columns.Add("類1組距" + subjectIndex + "count50Up"); table.Columns.Add("類2組距" + subjectIndex + "count50Up");
                //table.Columns.Add("班組距" + subjectIndex + "count40Up"); table.Columns.Add("科組距" + subjectIndex + "count40Up"); table.Columns.Add("年組距" + subjectIndex + "count40Up"); table.Columns.Add("類1組距" + subjectIndex + "count40Up"); table.Columns.Add("類2組距" + subjectIndex + "count40Up");
                //table.Columns.Add("班組距" + subjectIndex + "count30Up"); table.Columns.Add("科組距" + subjectIndex + "count30Up"); table.Columns.Add("年組距" + subjectIndex + "count30Up"); table.Columns.Add("類1組距" + subjectIndex + "count30Up"); table.Columns.Add("類2組距" + subjectIndex + "count30Up");
                //table.Columns.Add("班組距" + subjectIndex + "count20Up"); table.Columns.Add("科組距" + subjectIndex + "count20Up"); table.Columns.Add("年組距" + subjectIndex + "count20Up"); table.Columns.Add("類1組距" + subjectIndex + "count20Up"); table.Columns.Add("類2組距" + subjectIndex + "count20Up");
                //table.Columns.Add("班組距" + subjectIndex + "count10Up"); table.Columns.Add("科組距" + subjectIndex + "count10Up"); table.Columns.Add("年組距" + subjectIndex + "count10Up"); table.Columns.Add("類1組距" + subjectIndex + "count10Up"); table.Columns.Add("類2組距" + subjectIndex + "count10Up");
                //table.Columns.Add("班組距" + subjectIndex + "count90Down"); table.Columns.Add("科組距" + subjectIndex + "count90Down"); table.Columns.Add("年組距" + subjectIndex + "count90Down"); table.Columns.Add("類1組距" + subjectIndex + "count90Down"); table.Columns.Add("類2組距" + subjectIndex + "count90Down");
                //table.Columns.Add("班組距" + subjectIndex + "count80Down"); table.Columns.Add("科組距" + subjectIndex + "count80Down"); table.Columns.Add("年組距" + subjectIndex + "count80Down"); table.Columns.Add("類1組距" + subjectIndex + "count80Down"); table.Columns.Add("類2組距" + subjectIndex + "count80Down");
                //table.Columns.Add("班組距" + subjectIndex + "count70Down"); table.Columns.Add("科組距" + subjectIndex + "count70Down"); table.Columns.Add("年組距" + subjectIndex + "count70Down"); table.Columns.Add("類1組距" + subjectIndex + "count70Down"); table.Columns.Add("類2組距" + subjectIndex + "count70Down");
                //table.Columns.Add("班組距" + subjectIndex + "count60Down"); table.Columns.Add("科組距" + subjectIndex + "count60Down"); table.Columns.Add("年組距" + subjectIndex + "count60Down"); table.Columns.Add("類1組距" + subjectIndex + "count60Down"); table.Columns.Add("類2組距" + subjectIndex + "count60Down");
                //table.Columns.Add("班組距" + subjectIndex + "count50Down"); table.Columns.Add("科組距" + subjectIndex + "count50Down"); table.Columns.Add("年組距" + subjectIndex + "count50Down"); table.Columns.Add("類1組距" + subjectIndex + "count50Down"); table.Columns.Add("類2組距" + subjectIndex + "count50Down");
                //table.Columns.Add("班組距" + subjectIndex + "count40Down"); table.Columns.Add("科組距" + subjectIndex + "count40Down"); table.Columns.Add("年組距" + subjectIndex + "count40Down"); table.Columns.Add("類1組距" + subjectIndex + "count40Down"); table.Columns.Add("類2組距" + subjectIndex + "count40Down");
                //table.Columns.Add("班組距" + subjectIndex + "count30Down"); table.Columns.Add("科組距" + subjectIndex + "count30Down"); table.Columns.Add("年組距" + subjectIndex + "count30Down"); table.Columns.Add("類1組距" + subjectIndex + "count30Down"); table.Columns.Add("類2組距" + subjectIndex + "count30Down");
                //table.Columns.Add("班組距" + subjectIndex + "count20Down"); table.Columns.Add("科組距" + subjectIndex + "count20Down"); table.Columns.Add("年組距" + subjectIndex + "count20Down"); table.Columns.Add("類1組距" + subjectIndex + "count20Down"); table.Columns.Add("類2組距" + subjectIndex + "count20Down");
                //table.Columns.Add("班組距" + subjectIndex + "count10Down"); table.Columns.Add("科組距" + subjectIndex + "count10Down"); table.Columns.Add("年組距" + subjectIndex + "count10Down"); table.Columns.Add("類1組距" + subjectIndex + "count10Down"); table.Columns.Add("類2組距" + subjectIndex + "count10Down");
            }
            #endregion

            // 【總計成績】【級距分析】
            //{  平均(算數) 、總分(算數)、加權平均 、平均(算數) } *
            //{  頂標、高標 、均標、低標、底標 } *
            //{  班、 科、校 } *

            // 產稱 總計成績 五標
            string[] items = new string[] { "總分", "平均", "加權平均", "加權總分" };

            string[] rankNames = new string[] { "班", "科", "年" };

            string[] intervals = new string[] { "頂標", "高標", "均標", "低標", "底標", "新頂標", "新前標", "新均標", "新後標", "新底標", "標準差" };

            foreach (string item in items) //加權平均、平均、加權總分、總分
            {
                foreach (string rankName in rankNames)
                {
                    foreach (string interval in intervals)
                    {
                        string columsHeader = $"{item}{rankName}{interval}";

                        table.Columns.Add(columsHeader);
                    }
                }
            }

            #region Jean 用迴圈替換掉


            #endregion
            table.Columns.Add("類1總分底標");
            table.Columns.Add("類1平均底標");
            table.Columns.Add("類1加權總分底標");
            table.Columns.Add("類1加權平均底標");
            table.Columns.Add("類2總分底標");
            table.Columns.Add("類2平均底標");
            table.Columns.Add("類2加權總分底標");
            table.Columns.Add("類2加權平均底標");

            //刪除標準差

            table.Columns.Add("總分班組距count90"); table.Columns.Add("總分科組距count90"); table.Columns.Add("總分年組距count90"); table.Columns.Add("平均班組距count90"); table.Columns.Add("平均科組距count90"); table.Columns.Add("平均年組距count90"); table.Columns.Add("加權總分班組距count90"); table.Columns.Add("加權總分科組距count90"); table.Columns.Add("加權總分年組距count90"); table.Columns.Add("加權平均班組距count90"); table.Columns.Add("加權平均科組距count90"); table.Columns.Add("加權平均年組距count90"); table.Columns.Add("類1總分組距count90"); table.Columns.Add("類1平均組距count90"); table.Columns.Add("類1加權總分組距count90"); table.Columns.Add("類1加權平均組距count90"); table.Columns.Add("類2總分組距count90"); table.Columns.Add("類2平均組距count90"); table.Columns.Add("類2加權總分組距count90"); table.Columns.Add("類2加權平均組距count90");
            table.Columns.Add("總分班組距count80"); table.Columns.Add("總分科組距count80"); table.Columns.Add("總分年組距count80"); table.Columns.Add("平均班組距count80"); table.Columns.Add("平均科組距count80"); table.Columns.Add("平均年組距count80"); table.Columns.Add("加權總分班組距count80"); table.Columns.Add("加權總分科組距count80"); table.Columns.Add("加權總分年組距count80"); table.Columns.Add("加權平均班組距count80"); table.Columns.Add("加權平均科組距count80"); table.Columns.Add("加權平均年組距count80"); table.Columns.Add("類1總分組距count80"); table.Columns.Add("類1平均組距count80"); table.Columns.Add("類1加權總分組距count80"); table.Columns.Add("類1加權平均組距count80"); table.Columns.Add("類2總分組距count80"); table.Columns.Add("類2平均組距count80"); table.Columns.Add("類2加權總分組距count80"); table.Columns.Add("類2加權平均組距count80");
            table.Columns.Add("總分班組距count70"); table.Columns.Add("總分科組距count70"); table.Columns.Add("總分年組距count70"); table.Columns.Add("平均班組距count70"); table.Columns.Add("平均科組距count70"); table.Columns.Add("平均年組距count70"); table.Columns.Add("加權總分班組距count70"); table.Columns.Add("加權總分科組距count70"); table.Columns.Add("加權總分年組距count70"); table.Columns.Add("加權平均班組距count70"); table.Columns.Add("加權平均科組距count70"); table.Columns.Add("加權平均年組距count70"); table.Columns.Add("類1總分組距count70"); table.Columns.Add("類1平均組距count70"); table.Columns.Add("類1加權總分組距count70"); table.Columns.Add("類1加權平均組距count70"); table.Columns.Add("類2總分組距count70"); table.Columns.Add("類2平均組距count70"); table.Columns.Add("類2加權總分組距count70"); table.Columns.Add("類2加權平均組距count70");
            table.Columns.Add("總分班組距count60"); table.Columns.Add("總分科組距count60"); table.Columns.Add("總分年組距count60"); table.Columns.Add("平均班組距count60"); table.Columns.Add("平均科組距count60"); table.Columns.Add("平均年組距count60"); table.Columns.Add("加權總分班組距count60"); table.Columns.Add("加權總分科組距count60"); table.Columns.Add("加權總分年組距count60"); table.Columns.Add("加權平均班組距count60"); table.Columns.Add("加權平均科組距count60"); table.Columns.Add("加權平均年組距count60"); table.Columns.Add("類1總分組距count60"); table.Columns.Add("類1平均組距count60"); table.Columns.Add("類1加權總分組距count60"); table.Columns.Add("類1加權平均組距count60"); table.Columns.Add("類2總分組距count60"); table.Columns.Add("類2平均組距count60"); table.Columns.Add("類2加權總分組距count60"); table.Columns.Add("類2加權平均組距count60");
            table.Columns.Add("總分班組距count50"); table.Columns.Add("總分科組距count50"); table.Columns.Add("總分年組距count50"); table.Columns.Add("平均班組距count50"); table.Columns.Add("平均科組距count50"); table.Columns.Add("平均年組距count50"); table.Columns.Add("加權總分班組距count50"); table.Columns.Add("加權總分科組距count50"); table.Columns.Add("加權總分年組距count50"); table.Columns.Add("加權平均班組距count50"); table.Columns.Add("加權平均科組距count50"); table.Columns.Add("加權平均年組距count50"); table.Columns.Add("類1總分組距count50"); table.Columns.Add("類1平均組距count50"); table.Columns.Add("類1加權總分組距count50"); table.Columns.Add("類1加權平均組距count50"); table.Columns.Add("類2總分組距count50"); table.Columns.Add("類2平均組距count50"); table.Columns.Add("類2加權總分組距count50"); table.Columns.Add("類2加權平均組距count50");
            table.Columns.Add("總分班組距count40"); table.Columns.Add("總分科組距count40"); table.Columns.Add("總分年組距count40"); table.Columns.Add("平均班組距count40"); table.Columns.Add("平均科組距count40"); table.Columns.Add("平均年組距count40"); table.Columns.Add("加權總分班組距count40"); table.Columns.Add("加權總分科組距count40"); table.Columns.Add("加權總分年組距count40"); table.Columns.Add("加權平均班組距count40"); table.Columns.Add("加權平均科組距count40"); table.Columns.Add("加權平均年組距count40"); table.Columns.Add("類1總分組距count40"); table.Columns.Add("類1平均組距count40"); table.Columns.Add("類1加權總分組距count40"); table.Columns.Add("類1加權平均組距count40"); table.Columns.Add("類2總分組距count40"); table.Columns.Add("類2平均組距count40"); table.Columns.Add("類2加權總分組距count40"); table.Columns.Add("類2加權平均組距count40");
            table.Columns.Add("總分班組距count30"); table.Columns.Add("總分科組距count30"); table.Columns.Add("總分年組距count30"); table.Columns.Add("平均班組距count30"); table.Columns.Add("平均科組距count30"); table.Columns.Add("平均年組距count30"); table.Columns.Add("加權總分班組距count30"); table.Columns.Add("加權總分科組距count30"); table.Columns.Add("加權總分年組距count30"); table.Columns.Add("加權平均班組距count30"); table.Columns.Add("加權平均科組距count30"); table.Columns.Add("加權平均年組距count30"); table.Columns.Add("類1總分組距count30"); table.Columns.Add("類1平均組距count30"); table.Columns.Add("類1加權總分組距count30"); table.Columns.Add("類1加權平均組距count30"); table.Columns.Add("類2總分組距count30"); table.Columns.Add("類2平均組距count30"); table.Columns.Add("類2加權總分組距count30"); table.Columns.Add("類2加權平均組距count30");
            table.Columns.Add("總分班組距count20"); table.Columns.Add("總分科組距count20"); table.Columns.Add("總分年組距count20"); table.Columns.Add("平均班組距count20"); table.Columns.Add("平均科組距count20"); table.Columns.Add("平均年組距count20"); table.Columns.Add("加權總分班組距count20"); table.Columns.Add("加權總分科組距count20"); table.Columns.Add("加權總分年組距count20"); table.Columns.Add("加權平均班組距count20"); table.Columns.Add("加權平均科組距count20"); table.Columns.Add("加權平均年組距count20"); table.Columns.Add("類1總分組距count20"); table.Columns.Add("類1平均組距count20"); table.Columns.Add("類1加權總分組距count20"); table.Columns.Add("類1加權平均組距count20"); table.Columns.Add("類2總分組距count20"); table.Columns.Add("類2平均組距count20"); table.Columns.Add("類2加權總分組距count20"); table.Columns.Add("類2加權平均組距count20");
            table.Columns.Add("總分班組距count10"); table.Columns.Add("總分科組距count10"); table.Columns.Add("總分年組距count10"); table.Columns.Add("平均班組距count10"); table.Columns.Add("平均科組距count10"); table.Columns.Add("平均年組距count10"); table.Columns.Add("加權總分班組距count10"); table.Columns.Add("加權總分科組距count10"); table.Columns.Add("加權總分年組距count10"); table.Columns.Add("加權平均班組距count10"); table.Columns.Add("加權平均科組距count10"); table.Columns.Add("加權平均年組距count10"); table.Columns.Add("類1總分組距count10"); table.Columns.Add("類1平均組距count10"); table.Columns.Add("類1加權總分組距count10"); table.Columns.Add("類1加權平均組距count10"); table.Columns.Add("類2總分組距count10"); table.Columns.Add("類2平均組距count10"); table.Columns.Add("類2加權總分組距count10"); table.Columns.Add("類2加權平均組距count10");
            table.Columns.Add("總分班組距count100Up"); table.Columns.Add("總分科組距count100Up"); table.Columns.Add("總分年組距count100Up"); table.Columns.Add("平均班組距count100Up"); table.Columns.Add("平均科組距count100Up"); table.Columns.Add("平均年組距count100Up"); table.Columns.Add("加權總分班組距count100Up"); table.Columns.Add("加權總分科組距count100Up"); table.Columns.Add("加權總分年組距count100Up"); table.Columns.Add("加權平均班組距count100Up"); table.Columns.Add("加權平均科組距count100Up"); table.Columns.Add("加權平均年組距count100Up"); table.Columns.Add("類1總分組距count100Up"); table.Columns.Add("類1平均組距count100Up"); table.Columns.Add("類1加權總分組距count100Up"); table.Columns.Add("類1加權平均組距count100Up"); table.Columns.Add("類2總分組距count100Up"); table.Columns.Add("類2平均組距count100Up"); table.Columns.Add("類2加權總分組距count100Up"); table.Columns.Add("類2加權平均組距count100Up");
            table.Columns.Add("總分班組距count90Up"); table.Columns.Add("總分科組距count90Up"); table.Columns.Add("總分年組距count90Up"); table.Columns.Add("平均班組距count90Up"); table.Columns.Add("平均科組距count90Up"); table.Columns.Add("平均年組距count90Up"); table.Columns.Add("加權總分班組距count90Up"); table.Columns.Add("加權總分科組距count90Up"); table.Columns.Add("加權總分年組距count90Up"); table.Columns.Add("加權平均班組距count90Up"); table.Columns.Add("加權平均科組距count90Up"); table.Columns.Add("加權平均年組距count90Up"); table.Columns.Add("類1總分組距count90Up"); table.Columns.Add("類1平均組距count90Up"); table.Columns.Add("類1加權總分組距count90Up"); table.Columns.Add("類1加權平均組距count90Up"); table.Columns.Add("類2總分組距count90Up"); table.Columns.Add("類2平均組距count90Up"); table.Columns.Add("類2加權總分組距count90Up"); table.Columns.Add("類2加權平均組距count90Up");
            table.Columns.Add("總分班組距count80Up"); table.Columns.Add("總分科組距count80Up"); table.Columns.Add("總分年組距count80Up"); table.Columns.Add("平均班組距count80Up"); table.Columns.Add("平均科組距count80Up"); table.Columns.Add("平均年組距count80Up"); table.Columns.Add("加權總分班組距count80Up"); table.Columns.Add("加權總分科組距count80Up"); table.Columns.Add("加權總分年組距count80Up"); table.Columns.Add("加權平均班組距count80Up"); table.Columns.Add("加權平均科組距count80Up"); table.Columns.Add("加權平均年組距count80Up"); table.Columns.Add("類1總分組距count80Up"); table.Columns.Add("類1平均組距count80Up"); table.Columns.Add("類1加權總分組距count80Up"); table.Columns.Add("類1加權平均組距count80Up"); table.Columns.Add("類2總分組距count80Up"); table.Columns.Add("類2平均組距count80Up"); table.Columns.Add("類2加權總分組距count80Up"); table.Columns.Add("類2加權平均組距count80Up");
            table.Columns.Add("總分班組距count70Up"); table.Columns.Add("總分科組距count70Up"); table.Columns.Add("總分年組距count70Up"); table.Columns.Add("平均班組距count70Up"); table.Columns.Add("平均科組距count70Up"); table.Columns.Add("平均年組距count70Up"); table.Columns.Add("加權總分班組距count70Up"); table.Columns.Add("加權總分科組距count70Up"); table.Columns.Add("加權總分年組距count70Up"); table.Columns.Add("加權平均班組距count70Up"); table.Columns.Add("加權平均科組距count70Up"); table.Columns.Add("加權平均年組距count70Up"); table.Columns.Add("類1總分組距count70Up"); table.Columns.Add("類1平均組距count70Up"); table.Columns.Add("類1加權總分組距count70Up"); table.Columns.Add("類1加權平均組距count70Up"); table.Columns.Add("類2總分組距count70Up"); table.Columns.Add("類2平均組距count70Up"); table.Columns.Add("類2加權總分組距count70Up"); table.Columns.Add("類2加權平均組距count70Up");
            table.Columns.Add("總分班組距count60Up"); table.Columns.Add("總分科組距count60Up"); table.Columns.Add("總分年組距count60Up"); table.Columns.Add("平均班組距count60Up"); table.Columns.Add("平均科組距count60Up"); table.Columns.Add("平均年組距count60Up"); table.Columns.Add("加權總分班組距count60Up"); table.Columns.Add("加權總分科組距count60Up"); table.Columns.Add("加權總分年組距count60Up"); table.Columns.Add("加權平均班組距count60Up"); table.Columns.Add("加權平均科組距count60Up"); table.Columns.Add("加權平均年組距count60Up"); table.Columns.Add("類1總分組距count60Up"); table.Columns.Add("類1平均組距count60Up"); table.Columns.Add("類1加權總分組距count60Up"); table.Columns.Add("類1加權平均組距count60Up"); table.Columns.Add("類2總分組距count60Up"); table.Columns.Add("類2平均組距count60Up"); table.Columns.Add("類2加權總分組距count60Up"); table.Columns.Add("類2加權平均組距count60Up");
            table.Columns.Add("總分班組距count50Up"); table.Columns.Add("總分科組距count50Up"); table.Columns.Add("總分年組距count50Up"); table.Columns.Add("平均班組距count50Up"); table.Columns.Add("平均科組距count50Up"); table.Columns.Add("平均年組距count50Up"); table.Columns.Add("加權總分班組距count50Up"); table.Columns.Add("加權總分科組距count50Up"); table.Columns.Add("加權總分年組距count50Up"); table.Columns.Add("加權平均班組距count50Up"); table.Columns.Add("加權平均科組距count50Up"); table.Columns.Add("加權平均年組距count50Up"); table.Columns.Add("類1總分組距count50Up"); table.Columns.Add("類1平均組距count50Up"); table.Columns.Add("類1加權總分組距count50Up"); table.Columns.Add("類1加權平均組距count50Up"); table.Columns.Add("類2總分組距count50Up"); table.Columns.Add("類2平均組距count50Up"); table.Columns.Add("類2加權總分組距count50Up"); table.Columns.Add("類2加權平均組距count50Up");
            table.Columns.Add("總分班組距count40Up"); table.Columns.Add("總分科組距count40Up"); table.Columns.Add("總分年組距count40Up"); table.Columns.Add("平均班組距count40Up"); table.Columns.Add("平均科組距count40Up"); table.Columns.Add("平均年組距count40Up"); table.Columns.Add("加權總分班組距count40Up"); table.Columns.Add("加權總分科組距count40Up"); table.Columns.Add("加權總分年組距count40Up"); table.Columns.Add("加權平均班組距count40Up"); table.Columns.Add("加權平均科組距count40Up"); table.Columns.Add("加權平均年組距count40Up"); table.Columns.Add("類1總分組距count40Up"); table.Columns.Add("類1平均組距count40Up"); table.Columns.Add("類1加權總分組距count40Up"); table.Columns.Add("類1加權平均組距count40Up"); table.Columns.Add("類2總分組距count40Up"); table.Columns.Add("類2平均組距count40Up"); table.Columns.Add("類2加權總分組距count40Up"); table.Columns.Add("類2加權平均組距count40Up");
            table.Columns.Add("總分班組距count30Up"); table.Columns.Add("總分科組距count30Up"); table.Columns.Add("總分年組距count30Up"); table.Columns.Add("平均班組距count30Up"); table.Columns.Add("平均科組距count30Up"); table.Columns.Add("平均年組距count30Up"); table.Columns.Add("加權總分班組距count30Up"); table.Columns.Add("加權總分科組距count30Up"); table.Columns.Add("加權總分年組距count30Up"); table.Columns.Add("加權平均班組距count30Up"); table.Columns.Add("加權平均科組距count30Up"); table.Columns.Add("加權平均年組距count30Up"); table.Columns.Add("類1總分組距count30Up"); table.Columns.Add("類1平均組距count30Up"); table.Columns.Add("類1加權總分組距count30Up"); table.Columns.Add("類1加權平均組距count30Up"); table.Columns.Add("類2總分組距count30Up"); table.Columns.Add("類2平均組距count30Up"); table.Columns.Add("類2加權總分組距count30Up"); table.Columns.Add("類2加權平均組距count30Up");
            table.Columns.Add("總分班組距count20Up"); table.Columns.Add("總分科組距count20Up"); table.Columns.Add("總分年組距count20Up"); table.Columns.Add("平均班組距count20Up"); table.Columns.Add("平均科組距count20Up"); table.Columns.Add("平均年組距count20Up"); table.Columns.Add("加權總分班組距count20Up"); table.Columns.Add("加權總分科組距count20Up"); table.Columns.Add("加權總分年組距count20Up"); table.Columns.Add("加權平均班組距count20Up"); table.Columns.Add("加權平均科組距count20Up"); table.Columns.Add("加權平均年組距count20Up"); table.Columns.Add("類1總分組距count20Up"); table.Columns.Add("類1平均組距count20Up"); table.Columns.Add("類1加權總分組距count20Up"); table.Columns.Add("類1加權平均組距count20Up"); table.Columns.Add("類2總分組距count20Up"); table.Columns.Add("類2平均組距count20Up"); table.Columns.Add("類2加權總分組距count20Up"); table.Columns.Add("類2加權平均組距count20Up");
            table.Columns.Add("總分班組距count10Up"); table.Columns.Add("總分科組距count10Up"); table.Columns.Add("總分年組距count10Up"); table.Columns.Add("平均班組距count10Up"); table.Columns.Add("平均科組距count10Up"); table.Columns.Add("平均年組距count10Up"); table.Columns.Add("加權總分班組距count10Up"); table.Columns.Add("加權總分科組距count10Up"); table.Columns.Add("加權總分年組距count10Up"); table.Columns.Add("加權平均班組距count10Up"); table.Columns.Add("加權平均科組距count10Up"); table.Columns.Add("加權平均年組距count10Up"); table.Columns.Add("類1總分組距count10Up"); table.Columns.Add("類1平均組距count10Up"); table.Columns.Add("類1加權總分組距count10Up"); table.Columns.Add("類1加權平均組距count10Up"); table.Columns.Add("類2總分組距count10Up"); table.Columns.Add("類2平均組距count10Up"); table.Columns.Add("類2加權總分組距count10Up"); table.Columns.Add("類2加權平均組距count10Up");
            table.Columns.Add("總分班組距count90Down"); table.Columns.Add("總分科組距count90Down"); table.Columns.Add("總分年組距count90Down"); table.Columns.Add("平均班組距count90Down"); table.Columns.Add("平均科組距count90Down"); table.Columns.Add("平均年組距count90Down"); table.Columns.Add("加權總分班組距count90Down"); table.Columns.Add("加權總分科組距count90Down"); table.Columns.Add("加權總分年組距count90Down"); table.Columns.Add("加權平均班組距count90Down"); table.Columns.Add("加權平均科組距count90Down"); table.Columns.Add("加權平均年組距count90Down"); table.Columns.Add("類1總分組距count90Down"); table.Columns.Add("類1平均組距count90Down"); table.Columns.Add("類1加權總分組距count90Down"); table.Columns.Add("類1加權平均組距count90Down"); table.Columns.Add("類2總分組距count90Down"); table.Columns.Add("類2平均組距count90Down"); table.Columns.Add("類2加權總分組距count90Down"); table.Columns.Add("類2加權平均組距count90Down");
            table.Columns.Add("總分班組距count80Down"); table.Columns.Add("總分科組距count80Down"); table.Columns.Add("總分年組距count80Down"); table.Columns.Add("平均班組距count80Down"); table.Columns.Add("平均科組距count80Down"); table.Columns.Add("平均年組距count80Down"); table.Columns.Add("加權總分班組距count80Down"); table.Columns.Add("加權總分科組距count80Down"); table.Columns.Add("加權總分年組距count80Down"); table.Columns.Add("加權平均班組距count80Down"); table.Columns.Add("加權平均科組距count80Down"); table.Columns.Add("加權平均年組距count80Down"); table.Columns.Add("類1總分組距count80Down"); table.Columns.Add("類1平均組距count80Down"); table.Columns.Add("類1加權總分組距count80Down"); table.Columns.Add("類1加權平均組距count80Down"); table.Columns.Add("類2總分組距count80Down"); table.Columns.Add("類2平均組距count80Down"); table.Columns.Add("類2加權總分組距count80Down"); table.Columns.Add("類2加權平均組距count80Down");
            table.Columns.Add("總分班組距count70Down"); table.Columns.Add("總分科組距count70Down"); table.Columns.Add("總分年組距count70Down"); table.Columns.Add("平均班組距count70Down"); table.Columns.Add("平均科組距count70Down"); table.Columns.Add("平均年組距count70Down"); table.Columns.Add("加權總分班組距count70Down"); table.Columns.Add("加權總分科組距count70Down"); table.Columns.Add("加權總分年組距count70Down"); table.Columns.Add("加權平均班組距count70Down"); table.Columns.Add("加權平均科組距count70Down"); table.Columns.Add("加權平均年組距count70Down"); table.Columns.Add("類1總分組距count70Down"); table.Columns.Add("類1平均組距count70Down"); table.Columns.Add("類1加權總分組距count70Down"); table.Columns.Add("類1加權平均組距count70Down"); table.Columns.Add("類2總分組距count70Down"); table.Columns.Add("類2平均組距count70Down"); table.Columns.Add("類2加權總分組距count70Down"); table.Columns.Add("類2加權平均組距count70Down");
            table.Columns.Add("總分班組距count60Down"); table.Columns.Add("總分科組距count60Down"); table.Columns.Add("總分年組距count60Down"); table.Columns.Add("平均班組距count60Down"); table.Columns.Add("平均科組距count60Down"); table.Columns.Add("平均年組距count60Down"); table.Columns.Add("加權總分班組距count60Down"); table.Columns.Add("加權總分科組距count60Down"); table.Columns.Add("加權總分年組距count60Down"); table.Columns.Add("加權平均班組距count60Down"); table.Columns.Add("加權平均科組距count60Down"); table.Columns.Add("加權平均年組距count60Down"); table.Columns.Add("類1總分組距count60Down"); table.Columns.Add("類1平均組距count60Down"); table.Columns.Add("類1加權總分組距count60Down"); table.Columns.Add("類1加權平均組距count60Down"); table.Columns.Add("類2總分組距count60Down"); table.Columns.Add("類2平均組距count60Down"); table.Columns.Add("類2加權總分組距count60Down"); table.Columns.Add("類2加權平均組距count60Down");
            table.Columns.Add("總分班組距count50Down"); table.Columns.Add("總分科組距count50Down"); table.Columns.Add("總分年組距count50Down"); table.Columns.Add("平均班組距count50Down"); table.Columns.Add("平均科組距count50Down"); table.Columns.Add("平均年組距count50Down"); table.Columns.Add("加權總分班組距count50Down"); table.Columns.Add("加權總分科組距count50Down"); table.Columns.Add("加權總分年組距count50Down"); table.Columns.Add("加權平均班組距count50Down"); table.Columns.Add("加權平均科組距count50Down"); table.Columns.Add("加權平均年組距count50Down"); table.Columns.Add("類1總分組距count50Down"); table.Columns.Add("類1平均組距count50Down"); table.Columns.Add("類1加權總分組距count50Down"); table.Columns.Add("類1加權平均組距count50Down"); table.Columns.Add("類2總分組距count50Down"); table.Columns.Add("類2平均組距count50Down"); table.Columns.Add("類2加權總分組距count50Down"); table.Columns.Add("類2加權平均組距count50Down");
            table.Columns.Add("總分班組距count40Down"); table.Columns.Add("總分科組距count40Down"); table.Columns.Add("總分年組距count40Down"); table.Columns.Add("平均班組距count40Down"); table.Columns.Add("平均科組距count40Down"); table.Columns.Add("平均年組距count40Down"); table.Columns.Add("加權總分班組距count40Down"); table.Columns.Add("加權總分科組距count40Down"); table.Columns.Add("加權總分年組距count40Down"); table.Columns.Add("加權平均班組距count40Down"); table.Columns.Add("加權平均科組距count40Down"); table.Columns.Add("加權平均年組距count40Down"); table.Columns.Add("類1總分組距count40Down"); table.Columns.Add("類1平均組距count40Down"); table.Columns.Add("類1加權總分組距count40Down"); table.Columns.Add("類1加權平均組距count40Down"); table.Columns.Add("類2總分組距count40Down"); table.Columns.Add("類2平均組距count40Down"); table.Columns.Add("類2加權總分組距count40Down"); table.Columns.Add("類2加權平均組距count40Down");
            table.Columns.Add("總分班組距count30Down"); table.Columns.Add("總分科組距count30Down"); table.Columns.Add("總分年組距count30Down"); table.Columns.Add("平均班組距count30Down"); table.Columns.Add("平均科組距count30Down"); table.Columns.Add("平均年組距count30Down"); table.Columns.Add("加權總分班組距count30Down"); table.Columns.Add("加權總分科組距count30Down"); table.Columns.Add("加權總分年組距count30Down"); table.Columns.Add("加權平均班組距count30Down"); table.Columns.Add("加權平均科組距count30Down"); table.Columns.Add("加權平均年組距count30Down"); table.Columns.Add("類1總分組距count30Down"); table.Columns.Add("類1平均組距count30Down"); table.Columns.Add("類1加權總分組距count30Down"); table.Columns.Add("類1加權平均組距count30Down"); table.Columns.Add("類2總分組距count30Down"); table.Columns.Add("類2平均組距count30Down"); table.Columns.Add("類2加權總分組距count30Down"); table.Columns.Add("類2加權平均組距count30Down");
            table.Columns.Add("總分班組距count20Down"); table.Columns.Add("總分科組距count20Down"); table.Columns.Add("總分年組距count20Down"); table.Columns.Add("平均班組距count20Down"); table.Columns.Add("平均科組距count20Down"); table.Columns.Add("平均年組距count20Down"); table.Columns.Add("加權總分班組距count20Down"); table.Columns.Add("加權總分科組距count20Down"); table.Columns.Add("加權總分年組距count20Down"); table.Columns.Add("加權平均班組距count20Down"); table.Columns.Add("加權平均科組距count20Down"); table.Columns.Add("加權平均年組距count20Down"); table.Columns.Add("類1總分組距count20Down"); table.Columns.Add("類1平均組距count20Down"); table.Columns.Add("類1加權總分組距count20Down"); table.Columns.Add("類1加權平均組距count20Down"); table.Columns.Add("類2總分組距count20Down"); table.Columns.Add("類2平均組距count20Down"); table.Columns.Add("類2加權總分組距count20Down"); table.Columns.Add("類2加權平均組距count20Down");
            table.Columns.Add("總分班組距count10Down"); table.Columns.Add("總分科組距count10Down"); table.Columns.Add("總分年組距count10Down"); table.Columns.Add("平均班組距count10Down"); table.Columns.Add("平均科組距count10Down"); table.Columns.Add("平均年組距count10Down"); table.Columns.Add("加權總分班組距count10Down"); table.Columns.Add("加權總分科組距count10Down"); table.Columns.Add("加權總分年組距count10Down"); table.Columns.Add("加權平均班組距count10Down"); table.Columns.Add("加權平均科組距count10Down"); table.Columns.Add("加權平均年組距count10Down"); table.Columns.Add("類1總分組距count10Down"); table.Columns.Add("類1平均組距count10Down"); table.Columns.Add("類1加權總分組距count10Down"); table.Columns.Add("類1加權平均組距count10Down"); table.Columns.Add("類2總分組距count10Down"); table.Columns.Add("類2平均組距count10Down"); table.Columns.Add("類2加權總分組距count10Down"); table.Columns.Add("類2加權平均組距count10Down");
            #endregion



            #endregion

            return table;
        }


        /// <summary>
        /// 取得有計算固定排名的類1 類2 相關 資訊
        /// </summary>
        /// <param name="schoolYear"></param>
        /// <param name="semester"></param>
        /// <param name="refExamId"></param>
        /// <param name="gradeYear"></param>
        /// <returns></returns>
        internal static Dictionary<string, List<string>> GetTagInfoFromRankMatrix(string schoolYear, string semester, string refExamId, List<ClassRecord> selectStudentRecords)
        {
            QueryHelper queryHelper = new QueryHelper();
            // 取得所選班級 的 年級
            List<string> gradeYears = selectStudentRecords.Select(x => x.GradeYear).ToList();

            Dictionary<string, List<string>> results = new Dictionary<string, List<string>>();

            string sql = @"		
 SELECT 
      school_year 
     , semester
	 , rank_name 
     , grade_year 
	 , rank_type
	 , ref_exam_id
	 , count(*)
FROM rank_matrix 
WHERE  
	school_year = {0} 
AND     semester = {1}
AND 	is_alive =true
AND     ref_exam_id  = {2}
AND     (rank_type  ='類別1排名' OR rank_type ='類別2排名')
AND grade_year IN ({3}) 
		    
GROUP BY  
	 school_year 
	 , semester
	 , rank_type
	 , rank_name
	 , grade_year
	 , ref_exam_id
	   ";

            sql = string.Format(sql, schoolYear, semester, refExamId, string.Join(",", gradeYears));
            DataTable dt = queryHelper.Select(sql);
            foreach (DataRow dr in dt.Rows)
            {
                string rankType = "" + dr["rank_type"];
                string rankName = "" + dr["rank_name"];

                if (!results.ContainsKey(rankType))
                {
                    results.Add(rankType, new List<string>());
                }
                results[rankType].Add(rankName);
            }

            return results;
        }

        /// <summary>
        ///　取得　特定　學年度　學期　試別　特定班級之 固定排名結算 類別1 類別2 下 的 分組(自然組、社會組 or 籃球社 )
        /// </summary>
        /// <param name="schoolYear"></param>
        /// <param name="semester"></param>
        /// <param name="refExamId"></param>
        /// <param name="selectStudentRecords"></param>
        /// <returns></returns>
        internal static Dictionary<string, Dictionary<string, List<string>>> GetTagInfoFromRankMatrixByClass(string schoolYear, string semester, string refExamId, List<ClassRecord> selectClassesRecords)
        {
            Dictionary<string, Dictionary<string, List<string>>> results = new Dictionary<string, Dictionary<string, List<string>>>();

            List<string> selectedClassesID = selectClassesRecords.Select(x => x.ClassID).ToList();

            QueryHelper queryHelper = new QueryHelper();

            string sql = @"		
WITH    rank_matrix_tag AS (
		SELECT
			  id
			, school_year 
			, semester
			, rank_name 
			, rank_type
			, ref_exam_id
			, item_name 
		FROM rank_matrix 
		WHERE 
			school_year = {0}
				AND     semester = {1}
				AND 	is_alive =true
				AND     ref_exam_id = {2}
				AND     (rank_type  ='類別1排名' OR rank_type ='類別2排名')
) SELECT
	student.ref_class_id 
   , rank_matrix_tag.rank_name
   , rank_matrix_tag.rank_type
FROM 
   rank_detail 
INNER  JOIN
    rank_matrix_tag
ON rank_matrix_tag.id =rank_detail.ref_matrix_id
	INNER JOIN  student 
ON student.id =rank_detail.ref_student_id 
	 WHERE ref_student_id IN ( SELECT id  FROM  student  WHERE ref_class_id  IN( {3}))   
GROUP BY 
 	student.ref_class_id 
   , rank_matrix_tag.rank_name
   , rank_matrix_tag.rank_type
   
	   ";

            sql = string.Format(sql, schoolYear, semester, refExamId, string.Join(",", selectedClassesID));
            DataTable dt = queryHelper.Select(sql);
            foreach (DataRow dr in dt.Rows)
            {
                string refClassID = "" + dr["ref_class_id"];
                string rankType = "" + dr["rank_type"];
                string rankName = "" + dr["rank_name"];


                if (!results.ContainsKey(refClassID))  //  1.如果還沒有此班級資訊
                {
                    results.Add(refClassID, new Dictionary<string, List<string>>());
                }
                else // 已經有此班級 
                {

                }

                if (!results[refClassID].ContainsKey(rankType)) // 2.是否已經有類別1排名 OR 類別1排名
                {
                    results[refClassID].Add(rankType, new List<string>());
                }
                else
                {

                }

                // 3.加入rankName
                if (!results[refClassID][rankType].Contains(rankName))
                {
                    results[refClassID][rankType].Add(rankName);
                }


            }
            return results;
        }


        /// <summary>
        /// 取得班級(普101) 下 (類別) 類別1排名 下 科目 (國文、英文)
        /// </summary>
        /// <returns></returns>
        internal static Dictionary<string, Dictionary<string, List<string>>> GetTagsSubjextFromRankMatrixByClass(string schoolYear, string semester, string refExamId, List<ClassRecord> selectClassesRecords, string tag1Ortag2)
        {
            Dictionary<string, Dictionary<string, List<string>>> result = new Dictionary<string, Dictionary<string, List<string>>>();

            List<string> selectedClassesID = selectClassesRecords.Select(x => x.ClassID).ToList();

            QueryHelper queryHelper = new QueryHelper();


            string sql = @"		WITH  target_student AS
			(
		     	SELECT *FROM student WHERE status =1 AND ref_class_id IN({4})
			),sp_rank_matrix AS (
				SELECT 
					*
				FROM 
					rank_matrix 
				WHERE 
					school_year ={0}
						AND     semester = {1}
						AND 	is_alive = true 
						AND     ref_exam_id = {2}
						AND    ( rank_type  ='類別1排名'  OR   rank_type  ='類別2排名')  
			) SELECT
				    ref_class_id
					,sp_rank_matrix.item_type
					,sp_rank_matrix.item_name
		            ,sp_rank_matrix.rank_type
				    ,sp_rank_matrix.rank_name
       
			  FROM 
					sp_rank_matrix 
			  INNER  JOIN   rank_detail 
				  ON sp_rank_matrix.id =rank_detail.ref_matrix_id 
			  INNER JOIN    target_student 
				 ON target_student.id = rank_detail.ref_student_id 
			WHERE item_type   ='定期評量/科目成績'	 
	          GROUP BY ref_class_id
						  	,sp_rank_matrix.item_type
						 , sp_rank_matrix.item_name
		               	 ,sp_rank_matrix.rank_type
						 ,sp_rank_matrix.rank_name";

            sql = string.Format(sql, schoolYear, semester, refExamId, tag1Ortag2, string.Join(",", selectedClassesID));

            DataTable dt = queryHelper.Select(sql);
            foreach (DataRow dr in dt.Rows)
            {
                string classID = "" + dr["ref_class_id"];
                string rankType = "" + dr["rank_type"];
                string itemName = "" + dr["item_name"];

                if (!result.ContainsKey(classID)) //如果沒有班級 先加班級
                {

                    result.Add(classID, new Dictionary<string, List<string>>());
                }

                if (!result[classID].ContainsKey(rankType)) // 如果沒有 類別
                {
                    result[classID].Add(rankType, new List<string>());
                }

                if (!result[classID][rankType].Contains(itemName))
                {
                    result[classID][rankType].Add(itemName);
                }
            }

            return result;

        }




        /// <summary>
        /// 回傳班級下(類別1,類別2 之學生結算時之類別)
        /// </summary>
        public static Dictionary<string, Dictionary<string, string>> GetStudentTagSByClass(string schoolYear, string semester, string refExamId, List<ClassRecord> selectClassesRecords, string tag1orTag2)
        {
            Dictionary<string, Dictionary<string, string>> result = new Dictionary<string, Dictionary<string, string>>();

            List<string> selectedClassesID = selectClassesRecords.Select(x => x.ClassID).ToList();

            QueryHelper queryHelper = new QueryHelper();

            string sql = @"	
		   
		WITH  target_student AS
			(
		     	SELECT *FROM student WHERE status =1 AND ref_class_id IN({3})
			),sp_rank_matrix AS (
				SELECT 
					*
				FROM 
					rank_matrix 
				WHERE 
					school_year = {0}
						AND     semester = {1}
						AND 	is_alive = true 
						AND     ref_exam_id = {2} 
						AND     rank_type  ='{4}' 
			) SELECT
				    ref_class_id
			        ,ref_student_id
		            ,sp_rank_matrix.rank_type
				    ,sp_rank_matrix.rank_name
	
			  FROM 
					sp_rank_matrix 
			  INNER  JOIN   rank_detail 
				  ON sp_rank_matrix.id =rank_detail.ref_matrix_id 
			  INNER JOIN    target_student 
				 ON target_student.id = rank_detail.ref_student_id 
	          GROUP BY ref_class_id
			              ,ref_student_id
		               	 ,sp_rank_matrix.rank_type
						 ,sp_rank_matrix.rank_name
";
            sql = string.Format(sql, schoolYear, semester, refExamId, string.Join(",", selectedClassesID), tag1orTag2); ;


            DataTable dt = queryHelper.Select(sql);
            foreach (DataRow dr in dt.Rows)
            {
                string classID = "" + dr["ref_class_id"];
                string studentID = "" + dr["ref_student_id"];
                string rankName = "" + dr["rank_name"];
                if (!result.ContainsKey(classID))
                {
                    result.Add(classID, new Dictionary<string, string>());

                }
                if (!result[classID].ContainsKey(studentID))
                {

                    result[classID].Add(studentID, rankName);
                };
            }


            return result;
        }

        /// <summary>
        /// 取得類別1 類別2 下分組與 五標 組距 之排列組合之list 
        /// </summary>
        /// <returns></returns>
        private static List<string> GetTagItemIntervalList()
        {
            List<string> columnsHeader = new List<string>();
            // 母群統計值

            Dictionary<string, string> items = new Dictionary<string, string> {
                        { "頂標" ,"頂標" }
                        ,{"高標" ,"高標"   }
                        ,{"均標" ,"均標"   }
                        ,{"低標" ,"低標"   }
                        ,{"底標" ,"底標"   }
                        ,{"新頂標" ,"新頂標"   }
                        ,{"新前標" ,"新前標"   }
                        ,{"新均標" ,"新均標"   }
                        ,{"新後標" ,"新後標"   }
                        ,{"新底標" ,"新底標"   }
                        ,{"標準差" ,"標準差"   }

                        ,{"100以上"      ,"組距count100Up"   }
                        ,{"90以上"       ,"組距count90Up"   }
                        ,{"90以上小於100","組距count90"   }
                        ,{"小於90"       ,"組距count90Down"   }
                        ,{"80以上"       ,"組距count80Up"   }
                        ,{"80以上小於90" ,"組距count80"   }
                        ,{"小於80"      ,"組距count80Down"   }
                        ,{"70以上"      ,"組距count70Up"   }
                        ,{"70以上小於80" ,"組距count70"   }
                        ,{"小於70"       ,"組距count70Down"   }
                        ,{"60以上"       ,"組距count60Up"   }
                        ,{"60以上小於70" ,"組距count60"   }
                        ,{"小於60"       ,"組距count60Down"   }
                        ,{"50以上"       ,"組距count50Up"   }
                        ,{"50以上小於60" ,"組距count50"   }
                        ,{"小於50"       ,"組距count50Down"   }
                        ,{"40以上"       ,"組距count40Up"   }
                        ,{"40以上小於50" ,"組距count40"   }
                        ,{"小於40"       ,"組距count40Down"   }
                        ,{"30以上"       ,"組距count30Up"   }
                        ,{"30以上小於40"  ,"組距count30"   }
                        ,{"小於30"       ,"組距count30Down"   }
                        ,{"20以上"       ,"組距count20Up"   }
                        ,{"20以上小於30" ,"組距count20"   }
                        ,{"小於20"       ,"組距count20Down"   }
                        ,{"10以上"       ,"組距count10Up"   }
                        ,{"10以上小於20" ,"組距count10"   }
                        ,{"小於10"       ,"組距count10Down" }
            };


            foreach (string tagName in new string[] { "類別1", "類別2" })
            {
                string[] TotalScoreTypes = new string[] { "總分", "平均", "加權總分", "加權平均" };  // 總計成績 有4種

                foreach (string totalScoreType in TotalScoreTypes) // (總分 、 平均 、加權總分 、 加權平均 ....)
                {
                    for (int i = 1; i < 11; i++) // 產生10個 
                    {
                        foreach (string itemKey in items.Keys) // (頂標 、前標 、均標 ....)
                        {
                            columnsHeader.Add($"{totalScoreType}{tagName}_分組{i}{items[itemKey]}");
                        }
                    }
                }
            }

            return columnsHeader;
        }


        /// <summary>
        /// 取的 類別1 類別2 下分組名稱 ex :自然組 社會組
        /// </summary>
        /// <returns></returns>
        public static List<string> GetTagItemNames()
        {
            List<string> TageItemNames = new List<string>();

            foreach (string tagName in new string[] { "類別1", "類別2" })
            {
                for (int i = 1; i < 11; i++) // 產生10個 
                {
                    TageItemNames.Add($"{tagName}_分組{i}名稱");
                }
            }
            return TageItemNames;
        }


        /// <summary>
        /// 取得 類別1 類別2 下排列 組合 
        /// </summary>
        /// <returns></returns>
        public static List<string> GetTagItems()
        {
            List<string> TageItemNames = new List<string>();

            foreach (string tagName in new string[] { "類別1", "類別2" })
            {
                for (int i = 1; i < 11; i++) // 產生10個 
                {
                    TageItemNames.Add($"{tagName}_分組{i}");
                }
            }
            return TageItemNames;
        }

        // 取得學生課程規畫表內9D科目名稱
        public static Dictionary<string, List<string>> GetStudent9DSubjectNameByID(List<string> studentIDs)
        {
            Dictionary<string, List<string>> value = new Dictionary<string, List<string>>();

            if (studentIDs.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = string.Format(@"
                WITH student_data AS(
                    SELECT
                        student.id AS student_id,
                        COALESCE(
                            student.ref_graduation_plan_id,
                            class.ref_graduation_plan_id
                        ) AS graduation_plan_id
                    FROM
                        student
                        LEFT JOIN class ON student.ref_class_id = class.id
                    WHERE
                        student.id IN({0})
                ),
                graduation_plan_expand AS(
                    SELECT
                        graduation_plan_id,
                        array_to_string(xpath('//Subject/@SubjectName', subject_ele), '') :: TEXT AS subject_name,
                        array_to_string(xpath('//Subject/@課程代碼', subject_ele), '') :: TEXT AS 課程代碼
                    FROM
                        (
                            SELECT
                                graduation_plan_id,
                                unnest(
                                    xpath(
                                        '//GraduationPlan/Subject',
                                        xmlparse(content graduation_plan.content)
                                    )
                                ) AS subject_ele
                            FROM
                                (
                                    SELECT
                                        DISTINCT graduation_plan_id
                                    FROM
                                        student_data
                                ) AS target_graduation_plan
                                INNER JOIN graduation_plan ON graduation_plan.id = target_graduation_plan.graduation_plan_id
                        ) AS graduation_plan_expand
                )
                SELECT
                    DISTINCT student_data.student_id,
                    graduation_plan_expand.subject_name                    
                FROM
                    student_data
                    LEFT JOIN graduation_plan_expand ON student_data.graduation_plan_id = graduation_plan_expand.graduation_plan_id
                WHERE
                    SUBSTRING(graduation_plan_expand.課程代碼, 17, 1) = '9'
                    AND SUBSTRING(graduation_plan_expand.課程代碼, 19, 1) = 'D'
                ORDER BY
                    subject_name
                ", string.Join(",", studentIDs.ToArray()));

                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr["student_id"] + "";
                    string SubjectName = dr["subject_name"] + "";
                    if (!value.ContainsKey(sid))
                        value.Add(sid, new List<string>());

                    if (!value[sid].Contains(SubjectName))
                        value[sid].Add(SubjectName);

                }
            }
            return value;
        }
    }
}
