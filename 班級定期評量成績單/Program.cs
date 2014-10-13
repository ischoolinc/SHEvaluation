using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.IO;
using System.Threading;
using SmartSchool.Customization.Data;
using SmartSchool.Customization.Data.StudentExtension;
using SmartSchool;

namespace 班級定期評量成績單
{
     public class Program
     {
          [FISCA.MainMethod]
          public static void Main()
          {
               var btn = K12.Presentation.NLDPanels.Class.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["班級定期評量成績單(測試版)"];
               btn.Enable = false;
               K12.Presentation.NLDPanels.Class.SelectedSourceChanged += delegate { btn.Enable = K12.Presentation.NLDPanels.Class.SelectedSource.Count > 0; };
               btn.Click += new EventHandler(Program_Click);
          }

          private static string GetNumber(decimal? p)
          {
               if (p == null) return "";
               string levelNumber;
               switch (((int)p.Value))
               {
                    #region 對應levelNumber
                    case 0:
                         levelNumber = "";
                         break;
                    case 1:
                         levelNumber = "Ⅰ";
                         break;
                    case 2:
                         levelNumber = "Ⅱ";
                         break;
                    case 3:
                         levelNumber = "Ⅲ";
                         break;
                    case 4:
                         levelNumber = "Ⅳ";
                         break;
                    case 5:
                         levelNumber = "Ⅴ";
                         break;
                    case 6:
                         levelNumber = "Ⅵ";
                         break;
                    case 7:
                         levelNumber = "Ⅶ";
                         break;
                    case 8:
                         levelNumber = "Ⅷ";
                         break;
                    case 9:
                         levelNumber = "Ⅸ";
                         break;
                    case 10:
                         levelNumber = "Ⅹ";
                         break;
                    default:
                         levelNumber = "" + (p);
                         break;
                    #endregion
               }
               return levelNumber;
          }

          private static void Program_Click(object sender_, EventArgs e_)
          {

               ConfigForm form = new ConfigForm();
               if (form.ShowDialog() == DialogResult.OK)
               {
                    AccessHelper accessHelper = new AccessHelper();
                    List<ClassRecord> overflowRecords = new List<ClassRecord>();
                    Exception exc = null;
                    //取得列印設定
                    Configure conf = form.Configure;
                    //建立測試的選取學生
                    List<ClassRecord> selectedClasses = accessHelper.ClassHelper.GetSelectedClass();
                    List<StudentRecord> selectedStudents = new List<StudentRecord>();
                    foreach (ClassRecord classRec in selectedClasses)
                    {
                         foreach (StudentRecord stuRec in classRec.Students)
                         {
                              if (!selectedStudents.Contains(stuRec))
                                   selectedStudents.Add(stuRec);
                         }
                    }
                    //建立合併欄位總表
                    DataTable table = new DataTable();
                    #region 所有的合併欄位
                    table.Columns.Add("學校名稱");
                    table.Columns.Add("學校地址");
                    table.Columns.Add("學校電話");
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
                    //«通訊地址»«通訊地址郵遞區號»«通訊地址內容»
                    //«戶籍地址»«戶籍地址郵遞區號»«戶籍地址內容»
                    //«監護人»«父親»«母親»«科別名稱»
                    for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
                    {
                         table.Columns.Add("科目名稱" + subjectIndex);
                         table.Columns.Add("學分數" + subjectIndex);
                    }
                    for (int i = 1; i <= conf.StudentLimit; i++)
                    {
                         table.Columns.Add("座號" + i);
                         table.Columns.Add("學號" + i);
                         table.Columns.Add("姓名" + i);
                    }

                    for (int Num = 1; Num <= conf.StudentLimit; Num++)
                    {
                         for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
                         {
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
                              table.Columns.Add("全校排名" + Num + "-" + subjectIndex);
                              table.Columns.Add("全校排名母數" + Num + "-" + subjectIndex);
                         }
                         table.Columns.Add("總分" + Num);
                         table.Columns.Add("總分班排名" + Num);
                         table.Columns.Add("總分班排名母數" + Num);
                         table.Columns.Add("總分科排名" + Num);
                         table.Columns.Add("總分科排名母數" + Num);
                         table.Columns.Add("總分全校排名" + Num);
                         table.Columns.Add("總分全校排名母數" + Num);
                         table.Columns.Add("平均" + Num);
                         table.Columns.Add("平均班排名" + Num);
                         table.Columns.Add("平均班排名母數" + Num);
                         table.Columns.Add("平均科排名" + Num);
                         table.Columns.Add("平均科排名母數" + Num);
                         table.Columns.Add("平均全校排名" + Num);
                         table.Columns.Add("平均全校排名母數" + Num);
                         table.Columns.Add("加權總分" + Num);
                         table.Columns.Add("加權總分班排名" + Num);
                         table.Columns.Add("加權總分班排名母數" + Num);
                         table.Columns.Add("加權總分科排名" + Num);
                         table.Columns.Add("加權總分科排名母數" + Num);
                         table.Columns.Add("加權總分全校排名" + Num);
                         table.Columns.Add("加權總分全校排名母數" + Num);
                         table.Columns.Add("加權平均" + Num);
                         table.Columns.Add("加權平均班排名" + Num);
                         table.Columns.Add("加權平均班排名母數" + Num);
                         table.Columns.Add("加權平均科排名" + Num);
                         table.Columns.Add("加權平均科排名母數" + Num);
                         table.Columns.Add("加權平均全校排名" + Num);
                         table.Columns.Add("加權平均全校排名母數" + Num);

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
                    #region 瘋狂的組距及分析
                    #region 各科目組距及分析
                    for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
                    {
                         table.Columns.Add("班高標" + subjectIndex); table.Columns.Add("科高標" + subjectIndex); table.Columns.Add("校高標" + subjectIndex); table.Columns.Add("類1高標" + subjectIndex); table.Columns.Add("類2高標" + subjectIndex);
                         table.Columns.Add("班均標" + subjectIndex); table.Columns.Add("科均標" + subjectIndex); table.Columns.Add("校均標" + subjectIndex); table.Columns.Add("類1均標" + subjectIndex); table.Columns.Add("類2均標" + subjectIndex);
                         table.Columns.Add("班低標" + subjectIndex); table.Columns.Add("科低標" + subjectIndex); table.Columns.Add("校低標" + subjectIndex); table.Columns.Add("類1低標" + subjectIndex); table.Columns.Add("類2低標" + subjectIndex);
                         table.Columns.Add("班標準差" + subjectIndex); table.Columns.Add("科標準差" + subjectIndex); table.Columns.Add("校標準差" + subjectIndex); table.Columns.Add("類1標準差" + subjectIndex); table.Columns.Add("類2標準差" + subjectIndex);
                         table.Columns.Add("班組距" + subjectIndex + "count90"); table.Columns.Add("科組距" + subjectIndex + "count90"); table.Columns.Add("校組距" + subjectIndex + "count90"); table.Columns.Add("類1組距" + subjectIndex + "count90"); table.Columns.Add("類2組距" + subjectIndex + "count90");
                         table.Columns.Add("班組距" + subjectIndex + "count80"); table.Columns.Add("科組距" + subjectIndex + "count80"); table.Columns.Add("校組距" + subjectIndex + "count80"); table.Columns.Add("類1組距" + subjectIndex + "count80"); table.Columns.Add("類2組距" + subjectIndex + "count80");
                         table.Columns.Add("班組距" + subjectIndex + "count70"); table.Columns.Add("科組距" + subjectIndex + "count70"); table.Columns.Add("校組距" + subjectIndex + "count70"); table.Columns.Add("類1組距" + subjectIndex + "count70"); table.Columns.Add("類2組距" + subjectIndex + "count70");
                         table.Columns.Add("班組距" + subjectIndex + "count60"); table.Columns.Add("科組距" + subjectIndex + "count60"); table.Columns.Add("校組距" + subjectIndex + "count60"); table.Columns.Add("類1組距" + subjectIndex + "count60"); table.Columns.Add("類2組距" + subjectIndex + "count60");
                         table.Columns.Add("班組距" + subjectIndex + "count50"); table.Columns.Add("科組距" + subjectIndex + "count50"); table.Columns.Add("校組距" + subjectIndex + "count50"); table.Columns.Add("類1組距" + subjectIndex + "count50"); table.Columns.Add("類2組距" + subjectIndex + "count50");
                         table.Columns.Add("班組距" + subjectIndex + "count40"); table.Columns.Add("科組距" + subjectIndex + "count40"); table.Columns.Add("校組距" + subjectIndex + "count40"); table.Columns.Add("類1組距" + subjectIndex + "count40"); table.Columns.Add("類2組距" + subjectIndex + "count40");
                         table.Columns.Add("班組距" + subjectIndex + "count30"); table.Columns.Add("科組距" + subjectIndex + "count30"); table.Columns.Add("校組距" + subjectIndex + "count30"); table.Columns.Add("類1組距" + subjectIndex + "count30"); table.Columns.Add("類2組距" + subjectIndex + "count30");
                         table.Columns.Add("班組距" + subjectIndex + "count20"); table.Columns.Add("科組距" + subjectIndex + "count20"); table.Columns.Add("校組距" + subjectIndex + "count20"); table.Columns.Add("類1組距" + subjectIndex + "count20"); table.Columns.Add("類2組距" + subjectIndex + "count20");
                         table.Columns.Add("班組距" + subjectIndex + "count10"); table.Columns.Add("科組距" + subjectIndex + "count10"); table.Columns.Add("校組距" + subjectIndex + "count10"); table.Columns.Add("類1組距" + subjectIndex + "count10"); table.Columns.Add("類2組距" + subjectIndex + "count10");
                         table.Columns.Add("班組距" + subjectIndex + "count100Up"); table.Columns.Add("科組距" + subjectIndex + "count100Up"); table.Columns.Add("校組距" + subjectIndex + "count100Up"); table.Columns.Add("類1組距" + subjectIndex + "count100Up"); table.Columns.Add("類2組距" + subjectIndex + "count100Up");
                         table.Columns.Add("班組距" + subjectIndex + "count90Up"); table.Columns.Add("科組距" + subjectIndex + "count90Up"); table.Columns.Add("校組距" + subjectIndex + "count90Up"); table.Columns.Add("類1組距" + subjectIndex + "count90Up"); table.Columns.Add("類2組距" + subjectIndex + "count90Up");
                         table.Columns.Add("班組距" + subjectIndex + "count80Up"); table.Columns.Add("科組距" + subjectIndex + "count80Up"); table.Columns.Add("校組距" + subjectIndex + "count80Up"); table.Columns.Add("類1組距" + subjectIndex + "count80Up"); table.Columns.Add("類2組距" + subjectIndex + "count80Up");
                         table.Columns.Add("班組距" + subjectIndex + "count70Up"); table.Columns.Add("科組距" + subjectIndex + "count70Up"); table.Columns.Add("校組距" + subjectIndex + "count70Up"); table.Columns.Add("類1組距" + subjectIndex + "count70Up"); table.Columns.Add("類2組距" + subjectIndex + "count70Up");
                         table.Columns.Add("班組距" + subjectIndex + "count60Up"); table.Columns.Add("科組距" + subjectIndex + "count60Up"); table.Columns.Add("校組距" + subjectIndex + "count60Up"); table.Columns.Add("類1組距" + subjectIndex + "count60Up"); table.Columns.Add("類2組距" + subjectIndex + "count60Up");
                         table.Columns.Add("班組距" + subjectIndex + "count50Up"); table.Columns.Add("科組距" + subjectIndex + "count50Up"); table.Columns.Add("校組距" + subjectIndex + "count50Up"); table.Columns.Add("類1組距" + subjectIndex + "count50Up"); table.Columns.Add("類2組距" + subjectIndex + "count50Up");
                         table.Columns.Add("班組距" + subjectIndex + "count40Up"); table.Columns.Add("科組距" + subjectIndex + "count40Up"); table.Columns.Add("校組距" + subjectIndex + "count40Up"); table.Columns.Add("類1組距" + subjectIndex + "count40Up"); table.Columns.Add("類2組距" + subjectIndex + "count40Up");
                         table.Columns.Add("班組距" + subjectIndex + "count30Up"); table.Columns.Add("科組距" + subjectIndex + "count30Up"); table.Columns.Add("校組距" + subjectIndex + "count30Up"); table.Columns.Add("類1組距" + subjectIndex + "count30Up"); table.Columns.Add("類2組距" + subjectIndex + "count30Up");
                         table.Columns.Add("班組距" + subjectIndex + "count20Up"); table.Columns.Add("科組距" + subjectIndex + "count20Up"); table.Columns.Add("校組距" + subjectIndex + "count20Up"); table.Columns.Add("類1組距" + subjectIndex + "count20Up"); table.Columns.Add("類2組距" + subjectIndex + "count20Up");
                         table.Columns.Add("班組距" + subjectIndex + "count10Up"); table.Columns.Add("科組距" + subjectIndex + "count10Up"); table.Columns.Add("校組距" + subjectIndex + "count10Up"); table.Columns.Add("類1組距" + subjectIndex + "count10Up"); table.Columns.Add("類2組距" + subjectIndex + "count10Up");
                         table.Columns.Add("班組距" + subjectIndex + "count90Down"); table.Columns.Add("科組距" + subjectIndex + "count90Down"); table.Columns.Add("校組距" + subjectIndex + "count90Down"); table.Columns.Add("類1組距" + subjectIndex + "count90Down"); table.Columns.Add("類2組距" + subjectIndex + "count90Down");
                         table.Columns.Add("班組距" + subjectIndex + "count80Down"); table.Columns.Add("科組距" + subjectIndex + "count80Down"); table.Columns.Add("校組距" + subjectIndex + "count80Down"); table.Columns.Add("類1組距" + subjectIndex + "count80Down"); table.Columns.Add("類2組距" + subjectIndex + "count80Down");
                         table.Columns.Add("班組距" + subjectIndex + "count70Down"); table.Columns.Add("科組距" + subjectIndex + "count70Down"); table.Columns.Add("校組距" + subjectIndex + "count70Down"); table.Columns.Add("類1組距" + subjectIndex + "count70Down"); table.Columns.Add("類2組距" + subjectIndex + "count70Down");
                         table.Columns.Add("班組距" + subjectIndex + "count60Down"); table.Columns.Add("科組距" + subjectIndex + "count60Down"); table.Columns.Add("校組距" + subjectIndex + "count60Down"); table.Columns.Add("類1組距" + subjectIndex + "count60Down"); table.Columns.Add("類2組距" + subjectIndex + "count60Down");
                         table.Columns.Add("班組距" + subjectIndex + "count50Down"); table.Columns.Add("科組距" + subjectIndex + "count50Down"); table.Columns.Add("校組距" + subjectIndex + "count50Down"); table.Columns.Add("類1組距" + subjectIndex + "count50Down"); table.Columns.Add("類2組距" + subjectIndex + "count50Down");
                         table.Columns.Add("班組距" + subjectIndex + "count40Down"); table.Columns.Add("科組距" + subjectIndex + "count40Down"); table.Columns.Add("校組距" + subjectIndex + "count40Down"); table.Columns.Add("類1組距" + subjectIndex + "count40Down"); table.Columns.Add("類2組距" + subjectIndex + "count40Down");
                         table.Columns.Add("班組距" + subjectIndex + "count30Down"); table.Columns.Add("科組距" + subjectIndex + "count30Down"); table.Columns.Add("校組距" + subjectIndex + "count30Down"); table.Columns.Add("類1組距" + subjectIndex + "count30Down"); table.Columns.Add("類2組距" + subjectIndex + "count30Down");
                         table.Columns.Add("班組距" + subjectIndex + "count20Down"); table.Columns.Add("科組距" + subjectIndex + "count20Down"); table.Columns.Add("校組距" + subjectIndex + "count20Down"); table.Columns.Add("類1組距" + subjectIndex + "count20Down"); table.Columns.Add("類2組距" + subjectIndex + "count20Down");
                         table.Columns.Add("班組距" + subjectIndex + "count10Down"); table.Columns.Add("科組距" + subjectIndex + "count10Down"); table.Columns.Add("校組距" + subjectIndex + "count10Down"); table.Columns.Add("類1組距" + subjectIndex + "count10Down"); table.Columns.Add("類2組距" + subjectIndex + "count10Down");
                    }
                    #endregion
                    table.Columns.Add("總分班高標"); table.Columns.Add("總分科高標"); table.Columns.Add("總分校高標"); table.Columns.Add("平均班高標"); table.Columns.Add("平均科高標"); table.Columns.Add("平均校高標"); table.Columns.Add("加權總分班高標"); table.Columns.Add("加權總分科高標"); table.Columns.Add("加權總分校高標"); table.Columns.Add("加權平均班高標"); table.Columns.Add("加權平均科高標"); table.Columns.Add("加權平均校高標"); table.Columns.Add("類1總分高標"); table.Columns.Add("類1平均高標"); table.Columns.Add("類1加權總分高標"); table.Columns.Add("類1加權平均高標"); table.Columns.Add("類2總分高標"); table.Columns.Add("類2平均高標"); table.Columns.Add("類2加權總分高標"); table.Columns.Add("類2加權平均高標");
                    table.Columns.Add("總分班均標"); table.Columns.Add("總分科均標"); table.Columns.Add("總分校均標"); table.Columns.Add("平均班均標"); table.Columns.Add("平均科均標"); table.Columns.Add("平均校均標"); table.Columns.Add("加權總分班均標"); table.Columns.Add("加權總分科均標"); table.Columns.Add("加權總分校均標"); table.Columns.Add("加權平均班均標"); table.Columns.Add("加權平均科均標"); table.Columns.Add("加權平均校均標"); table.Columns.Add("類1總分均標"); table.Columns.Add("類1平均均標"); table.Columns.Add("類1加權總分均標"); table.Columns.Add("類1加權平均均標"); table.Columns.Add("類2總分均標"); table.Columns.Add("類2平均均標"); table.Columns.Add("類2加權總分均標"); table.Columns.Add("類2加權平均均標");
                    table.Columns.Add("總分班低標"); table.Columns.Add("總分科低標"); table.Columns.Add("總分校低標"); table.Columns.Add("平均班低標"); table.Columns.Add("平均科低標"); table.Columns.Add("平均校低標"); table.Columns.Add("加權總分班低標"); table.Columns.Add("加權總分科低標"); table.Columns.Add("加權總分校低標"); table.Columns.Add("加權平均班低標"); table.Columns.Add("加權平均科低標"); table.Columns.Add("加權平均校低標"); table.Columns.Add("類1總分低標"); table.Columns.Add("類1平均低標"); table.Columns.Add("類1加權總分低標"); table.Columns.Add("類1加權平均低標"); table.Columns.Add("類2總分低標"); table.Columns.Add("類2平均低標"); table.Columns.Add("類2加權總分低標"); table.Columns.Add("類2加權平均低標");
                    table.Columns.Add("總分班標準差"); table.Columns.Add("總分科標準差"); table.Columns.Add("總分校標準差"); table.Columns.Add("平均班標準差"); table.Columns.Add("平均科標準差"); table.Columns.Add("平均校標準差"); table.Columns.Add("加權總分班標準差"); table.Columns.Add("加權總分科標準差"); table.Columns.Add("加權總分校標準差"); table.Columns.Add("加權平均班標準差"); table.Columns.Add("加權平均科標準差"); table.Columns.Add("加權平均校標準差"); table.Columns.Add("類1總分標準差"); table.Columns.Add("類1平均標準差"); table.Columns.Add("類1加權總分標準差"); table.Columns.Add("類1加權平均標準差"); table.Columns.Add("類2總分標準差"); table.Columns.Add("類2平均標準差"); table.Columns.Add("類2加權總分標準差"); table.Columns.Add("類2加權平均標準差");
                    table.Columns.Add("總分班組距count90"); table.Columns.Add("總分科組距count90"); table.Columns.Add("總分校組距count90"); table.Columns.Add("平均班組距count90"); table.Columns.Add("平均科組距count90"); table.Columns.Add("平均校組距count90"); table.Columns.Add("加權總分班組距count90"); table.Columns.Add("加權總分科組距count90"); table.Columns.Add("加權總分校組距count90"); table.Columns.Add("加權平均班組距count90"); table.Columns.Add("加權平均科組距count90"); table.Columns.Add("加權平均校組距count90"); table.Columns.Add("類1總分組距count90"); table.Columns.Add("類1平均組距count90"); table.Columns.Add("類1加權總分組距count90"); table.Columns.Add("類1加權平均組距count90"); table.Columns.Add("類2總分組距count90"); table.Columns.Add("類2平均組距count90"); table.Columns.Add("類2加權總分組距count90"); table.Columns.Add("類2加權平均組距count90");
                    table.Columns.Add("總分班組距count80"); table.Columns.Add("總分科組距count80"); table.Columns.Add("總分校組距count80"); table.Columns.Add("平均班組距count80"); table.Columns.Add("平均科組距count80"); table.Columns.Add("平均校組距count80"); table.Columns.Add("加權總分班組距count80"); table.Columns.Add("加權總分科組距count80"); table.Columns.Add("加權總分校組距count80"); table.Columns.Add("加權平均班組距count80"); table.Columns.Add("加權平均科組距count80"); table.Columns.Add("加權平均校組距count80"); table.Columns.Add("類1總分組距count80"); table.Columns.Add("類1平均組距count80"); table.Columns.Add("類1加權總分組距count80"); table.Columns.Add("類1加權平均組距count80"); table.Columns.Add("類2總分組距count80"); table.Columns.Add("類2平均組距count80"); table.Columns.Add("類2加權總分組距count80"); table.Columns.Add("類2加權平均組距count80");
                    table.Columns.Add("總分班組距count70"); table.Columns.Add("總分科組距count70"); table.Columns.Add("總分校組距count70"); table.Columns.Add("平均班組距count70"); table.Columns.Add("平均科組距count70"); table.Columns.Add("平均校組距count70"); table.Columns.Add("加權總分班組距count70"); table.Columns.Add("加權總分科組距count70"); table.Columns.Add("加權總分校組距count70"); table.Columns.Add("加權平均班組距count70"); table.Columns.Add("加權平均科組距count70"); table.Columns.Add("加權平均校組距count70"); table.Columns.Add("類1總分組距count70"); table.Columns.Add("類1平均組距count70"); table.Columns.Add("類1加權總分組距count70"); table.Columns.Add("類1加權平均組距count70"); table.Columns.Add("類2總分組距count70"); table.Columns.Add("類2平均組距count70"); table.Columns.Add("類2加權總分組距count70"); table.Columns.Add("類2加權平均組距count70");
                    table.Columns.Add("總分班組距count60"); table.Columns.Add("總分科組距count60"); table.Columns.Add("總分校組距count60"); table.Columns.Add("平均班組距count60"); table.Columns.Add("平均科組距count60"); table.Columns.Add("平均校組距count60"); table.Columns.Add("加權總分班組距count60"); table.Columns.Add("加權總分科組距count60"); table.Columns.Add("加權總分校組距count60"); table.Columns.Add("加權平均班組距count60"); table.Columns.Add("加權平均科組距count60"); table.Columns.Add("加權平均校組距count60"); table.Columns.Add("類1總分組距count60"); table.Columns.Add("類1平均組距count60"); table.Columns.Add("類1加權總分組距count60"); table.Columns.Add("類1加權平均組距count60"); table.Columns.Add("類2總分組距count60"); table.Columns.Add("類2平均組距count60"); table.Columns.Add("類2加權總分組距count60"); table.Columns.Add("類2加權平均組距count60");
                    table.Columns.Add("總分班組距count50"); table.Columns.Add("總分科組距count50"); table.Columns.Add("總分校組距count50"); table.Columns.Add("平均班組距count50"); table.Columns.Add("平均科組距count50"); table.Columns.Add("平均校組距count50"); table.Columns.Add("加權總分班組距count50"); table.Columns.Add("加權總分科組距count50"); table.Columns.Add("加權總分校組距count50"); table.Columns.Add("加權平均班組距count50"); table.Columns.Add("加權平均科組距count50"); table.Columns.Add("加權平均校組距count50"); table.Columns.Add("類1總分組距count50"); table.Columns.Add("類1平均組距count50"); table.Columns.Add("類1加權總分組距count50"); table.Columns.Add("類1加權平均組距count50"); table.Columns.Add("類2總分組距count50"); table.Columns.Add("類2平均組距count50"); table.Columns.Add("類2加權總分組距count50"); table.Columns.Add("類2加權平均組距count50");
                    table.Columns.Add("總分班組距count40"); table.Columns.Add("總分科組距count40"); table.Columns.Add("總分校組距count40"); table.Columns.Add("平均班組距count40"); table.Columns.Add("平均科組距count40"); table.Columns.Add("平均校組距count40"); table.Columns.Add("加權總分班組距count40"); table.Columns.Add("加權總分科組距count40"); table.Columns.Add("加權總分校組距count40"); table.Columns.Add("加權平均班組距count40"); table.Columns.Add("加權平均科組距count40"); table.Columns.Add("加權平均校組距count40"); table.Columns.Add("類1總分組距count40"); table.Columns.Add("類1平均組距count40"); table.Columns.Add("類1加權總分組距count40"); table.Columns.Add("類1加權平均組距count40"); table.Columns.Add("類2總分組距count40"); table.Columns.Add("類2平均組距count40"); table.Columns.Add("類2加權總分組距count40"); table.Columns.Add("類2加權平均組距count40");
                    table.Columns.Add("總分班組距count30"); table.Columns.Add("總分科組距count30"); table.Columns.Add("總分校組距count30"); table.Columns.Add("平均班組距count30"); table.Columns.Add("平均科組距count30"); table.Columns.Add("平均校組距count30"); table.Columns.Add("加權總分班組距count30"); table.Columns.Add("加權總分科組距count30"); table.Columns.Add("加權總分校組距count30"); table.Columns.Add("加權平均班組距count30"); table.Columns.Add("加權平均科組距count30"); table.Columns.Add("加權平均校組距count30"); table.Columns.Add("類1總分組距count30"); table.Columns.Add("類1平均組距count30"); table.Columns.Add("類1加權總分組距count30"); table.Columns.Add("類1加權平均組距count30"); table.Columns.Add("類2總分組距count30"); table.Columns.Add("類2平均組距count30"); table.Columns.Add("類2加權總分組距count30"); table.Columns.Add("類2加權平均組距count30");
                    table.Columns.Add("總分班組距count20"); table.Columns.Add("總分科組距count20"); table.Columns.Add("總分校組距count20"); table.Columns.Add("平均班組距count20"); table.Columns.Add("平均科組距count20"); table.Columns.Add("平均校組距count20"); table.Columns.Add("加權總分班組距count20"); table.Columns.Add("加權總分科組距count20"); table.Columns.Add("加權總分校組距count20"); table.Columns.Add("加權平均班組距count20"); table.Columns.Add("加權平均科組距count20"); table.Columns.Add("加權平均校組距count20"); table.Columns.Add("類1總分組距count20"); table.Columns.Add("類1平均組距count20"); table.Columns.Add("類1加權總分組距count20"); table.Columns.Add("類1加權平均組距count20"); table.Columns.Add("類2總分組距count20"); table.Columns.Add("類2平均組距count20"); table.Columns.Add("類2加權總分組距count20"); table.Columns.Add("類2加權平均組距count20");
                    table.Columns.Add("總分班組距count10"); table.Columns.Add("總分科組距count10"); table.Columns.Add("總分校組距count10"); table.Columns.Add("平均班組距count10"); table.Columns.Add("平均科組距count10"); table.Columns.Add("平均校組距count10"); table.Columns.Add("加權總分班組距count10"); table.Columns.Add("加權總分科組距count10"); table.Columns.Add("加權總分校組距count10"); table.Columns.Add("加權平均班組距count10"); table.Columns.Add("加權平均科組距count10"); table.Columns.Add("加權平均校組距count10"); table.Columns.Add("類1總分組距count10"); table.Columns.Add("類1平均組距count10"); table.Columns.Add("類1加權總分組距count10"); table.Columns.Add("類1加權平均組距count10"); table.Columns.Add("類2總分組距count10"); table.Columns.Add("類2平均組距count10"); table.Columns.Add("類2加權總分組距count10"); table.Columns.Add("類2加權平均組距count10");
                    table.Columns.Add("總分班組距count100Up"); table.Columns.Add("總分科組距count100Up"); table.Columns.Add("總分校組距count100Up"); table.Columns.Add("平均班組距count100Up"); table.Columns.Add("平均科組距count100Up"); table.Columns.Add("平均校組距count100Up"); table.Columns.Add("加權總分班組距count100Up"); table.Columns.Add("加權總分科組距count100Up"); table.Columns.Add("加權總分校組距count100Up"); table.Columns.Add("加權平均班組距count100Up"); table.Columns.Add("加權平均科組距count100Up"); table.Columns.Add("加權平均校組距count100Up"); table.Columns.Add("類1總分組距count100Up"); table.Columns.Add("類1平均組距count100Up"); table.Columns.Add("類1加權總分組距count100Up"); table.Columns.Add("類1加權平均組距count100Up"); table.Columns.Add("類2總分組距count100Up"); table.Columns.Add("類2平均組距count100Up"); table.Columns.Add("類2加權總分組距count100Up"); table.Columns.Add("類2加權平均組距count100Up");
                    table.Columns.Add("總分班組距count90Up"); table.Columns.Add("總分科組距count90Up"); table.Columns.Add("總分校組距count90Up"); table.Columns.Add("平均班組距count90Up"); table.Columns.Add("平均科組距count90Up"); table.Columns.Add("平均校組距count90Up"); table.Columns.Add("加權總分班組距count90Up"); table.Columns.Add("加權總分科組距count90Up"); table.Columns.Add("加權總分校組距count90Up"); table.Columns.Add("加權平均班組距count90Up"); table.Columns.Add("加權平均科組距count90Up"); table.Columns.Add("加權平均校組距count90Up"); table.Columns.Add("類1總分組距count90Up"); table.Columns.Add("類1平均組距count90Up"); table.Columns.Add("類1加權總分組距count90Up"); table.Columns.Add("類1加權平均組距count90Up"); table.Columns.Add("類2總分組距count90Up"); table.Columns.Add("類2平均組距count90Up"); table.Columns.Add("類2加權總分組距count90Up"); table.Columns.Add("類2加權平均組距count90Up");
                    table.Columns.Add("總分班組距count80Up"); table.Columns.Add("總分科組距count80Up"); table.Columns.Add("總分校組距count80Up"); table.Columns.Add("平均班組距count80Up"); table.Columns.Add("平均科組距count80Up"); table.Columns.Add("平均校組距count80Up"); table.Columns.Add("加權總分班組距count80Up"); table.Columns.Add("加權總分科組距count80Up"); table.Columns.Add("加權總分校組距count80Up"); table.Columns.Add("加權平均班組距count80Up"); table.Columns.Add("加權平均科組距count80Up"); table.Columns.Add("加權平均校組距count80Up"); table.Columns.Add("類1總分組距count80Up"); table.Columns.Add("類1平均組距count80Up"); table.Columns.Add("類1加權總分組距count80Up"); table.Columns.Add("類1加權平均組距count80Up"); table.Columns.Add("類2總分組距count80Up"); table.Columns.Add("類2平均組距count80Up"); table.Columns.Add("類2加權總分組距count80Up"); table.Columns.Add("類2加權平均組距count80Up");
                    table.Columns.Add("總分班組距count70Up"); table.Columns.Add("總分科組距count70Up"); table.Columns.Add("總分校組距count70Up"); table.Columns.Add("平均班組距count70Up"); table.Columns.Add("平均科組距count70Up"); table.Columns.Add("平均校組距count70Up"); table.Columns.Add("加權總分班組距count70Up"); table.Columns.Add("加權總分科組距count70Up"); table.Columns.Add("加權總分校組距count70Up"); table.Columns.Add("加權平均班組距count70Up"); table.Columns.Add("加權平均科組距count70Up"); table.Columns.Add("加權平均校組距count70Up"); table.Columns.Add("類1總分組距count70Up"); table.Columns.Add("類1平均組距count70Up"); table.Columns.Add("類1加權總分組距count70Up"); table.Columns.Add("類1加權平均組距count70Up"); table.Columns.Add("類2總分組距count70Up"); table.Columns.Add("類2平均組距count70Up"); table.Columns.Add("類2加權總分組距count70Up"); table.Columns.Add("類2加權平均組距count70Up");
                    table.Columns.Add("總分班組距count60Up"); table.Columns.Add("總分科組距count60Up"); table.Columns.Add("總分校組距count60Up"); table.Columns.Add("平均班組距count60Up"); table.Columns.Add("平均科組距count60Up"); table.Columns.Add("平均校組距count60Up"); table.Columns.Add("加權總分班組距count60Up"); table.Columns.Add("加權總分科組距count60Up"); table.Columns.Add("加權總分校組距count60Up"); table.Columns.Add("加權平均班組距count60Up"); table.Columns.Add("加權平均科組距count60Up"); table.Columns.Add("加權平均校組距count60Up"); table.Columns.Add("類1總分組距count60Up"); table.Columns.Add("類1平均組距count60Up"); table.Columns.Add("類1加權總分組距count60Up"); table.Columns.Add("類1加權平均組距count60Up"); table.Columns.Add("類2總分組距count60Up"); table.Columns.Add("類2平均組距count60Up"); table.Columns.Add("類2加權總分組距count60Up"); table.Columns.Add("類2加權平均組距count60Up");
                    table.Columns.Add("總分班組距count50Up"); table.Columns.Add("總分科組距count50Up"); table.Columns.Add("總分校組距count50Up"); table.Columns.Add("平均班組距count50Up"); table.Columns.Add("平均科組距count50Up"); table.Columns.Add("平均校組距count50Up"); table.Columns.Add("加權總分班組距count50Up"); table.Columns.Add("加權總分科組距count50Up"); table.Columns.Add("加權總分校組距count50Up"); table.Columns.Add("加權平均班組距count50Up"); table.Columns.Add("加權平均科組距count50Up"); table.Columns.Add("加權平均校組距count50Up"); table.Columns.Add("類1總分組距count50Up"); table.Columns.Add("類1平均組距count50Up"); table.Columns.Add("類1加權總分組距count50Up"); table.Columns.Add("類1加權平均組距count50Up"); table.Columns.Add("類2總分組距count50Up"); table.Columns.Add("類2平均組距count50Up"); table.Columns.Add("類2加權總分組距count50Up"); table.Columns.Add("類2加權平均組距count50Up");
                    table.Columns.Add("總分班組距count40Up"); table.Columns.Add("總分科組距count40Up"); table.Columns.Add("總分校組距count40Up"); table.Columns.Add("平均班組距count40Up"); table.Columns.Add("平均科組距count40Up"); table.Columns.Add("平均校組距count40Up"); table.Columns.Add("加權總分班組距count40Up"); table.Columns.Add("加權總分科組距count40Up"); table.Columns.Add("加權總分校組距count40Up"); table.Columns.Add("加權平均班組距count40Up"); table.Columns.Add("加權平均科組距count40Up"); table.Columns.Add("加權平均校組距count40Up"); table.Columns.Add("類1總分組距count40Up"); table.Columns.Add("類1平均組距count40Up"); table.Columns.Add("類1加權總分組距count40Up"); table.Columns.Add("類1加權平均組距count40Up"); table.Columns.Add("類2總分組距count40Up"); table.Columns.Add("類2平均組距count40Up"); table.Columns.Add("類2加權總分組距count40Up"); table.Columns.Add("類2加權平均組距count40Up");
                    table.Columns.Add("總分班組距count30Up"); table.Columns.Add("總分科組距count30Up"); table.Columns.Add("總分校組距count30Up"); table.Columns.Add("平均班組距count30Up"); table.Columns.Add("平均科組距count30Up"); table.Columns.Add("平均校組距count30Up"); table.Columns.Add("加權總分班組距count30Up"); table.Columns.Add("加權總分科組距count30Up"); table.Columns.Add("加權總分校組距count30Up"); table.Columns.Add("加權平均班組距count30Up"); table.Columns.Add("加權平均科組距count30Up"); table.Columns.Add("加權平均校組距count30Up"); table.Columns.Add("類1總分組距count30Up"); table.Columns.Add("類1平均組距count30Up"); table.Columns.Add("類1加權總分組距count30Up"); table.Columns.Add("類1加權平均組距count30Up"); table.Columns.Add("類2總分組距count30Up"); table.Columns.Add("類2平均組距count30Up"); table.Columns.Add("類2加權總分組距count30Up"); table.Columns.Add("類2加權平均組距count30Up");
                    table.Columns.Add("總分班組距count20Up"); table.Columns.Add("總分科組距count20Up"); table.Columns.Add("總分校組距count20Up"); table.Columns.Add("平均班組距count20Up"); table.Columns.Add("平均科組距count20Up"); table.Columns.Add("平均校組距count20Up"); table.Columns.Add("加權總分班組距count20Up"); table.Columns.Add("加權總分科組距count20Up"); table.Columns.Add("加權總分校組距count20Up"); table.Columns.Add("加權平均班組距count20Up"); table.Columns.Add("加權平均科組距count20Up"); table.Columns.Add("加權平均校組距count20Up"); table.Columns.Add("類1總分組距count20Up"); table.Columns.Add("類1平均組距count20Up"); table.Columns.Add("類1加權總分組距count20Up"); table.Columns.Add("類1加權平均組距count20Up"); table.Columns.Add("類2總分組距count20Up"); table.Columns.Add("類2平均組距count20Up"); table.Columns.Add("類2加權總分組距count20Up"); table.Columns.Add("類2加權平均組距count20Up");
                    table.Columns.Add("總分班組距count10Up"); table.Columns.Add("總分科組距count10Up"); table.Columns.Add("總分校組距count10Up"); table.Columns.Add("平均班組距count10Up"); table.Columns.Add("平均科組距count10Up"); table.Columns.Add("平均校組距count10Up"); table.Columns.Add("加權總分班組距count10Up"); table.Columns.Add("加權總分科組距count10Up"); table.Columns.Add("加權總分校組距count10Up"); table.Columns.Add("加權平均班組距count10Up"); table.Columns.Add("加權平均科組距count10Up"); table.Columns.Add("加權平均校組距count10Up"); table.Columns.Add("類1總分組距count10Up"); table.Columns.Add("類1平均組距count10Up"); table.Columns.Add("類1加權總分組距count10Up"); table.Columns.Add("類1加權平均組距count10Up"); table.Columns.Add("類2總分組距count10Up"); table.Columns.Add("類2平均組距count10Up"); table.Columns.Add("類2加權總分組距count10Up"); table.Columns.Add("類2加權平均組距count10Up");
                    table.Columns.Add("總分班組距count90Down"); table.Columns.Add("總分科組距count90Down"); table.Columns.Add("總分校組距count90Down"); table.Columns.Add("平均班組距count90Down"); table.Columns.Add("平均科組距count90Down"); table.Columns.Add("平均校組距count90Down"); table.Columns.Add("加權總分班組距count90Down"); table.Columns.Add("加權總分科組距count90Down"); table.Columns.Add("加權總分校組距count90Down"); table.Columns.Add("加權平均班組距count90Down"); table.Columns.Add("加權平均科組距count90Down"); table.Columns.Add("加權平均校組距count90Down"); table.Columns.Add("類1總分組距count90Down"); table.Columns.Add("類1平均組距count90Down"); table.Columns.Add("類1加權總分組距count90Down"); table.Columns.Add("類1加權平均組距count90Down"); table.Columns.Add("類2總分組距count90Down"); table.Columns.Add("類2平均組距count90Down"); table.Columns.Add("類2加權總分組距count90Down"); table.Columns.Add("類2加權平均組距count90Down");
                    table.Columns.Add("總分班組距count80Down"); table.Columns.Add("總分科組距count80Down"); table.Columns.Add("總分校組距count80Down"); table.Columns.Add("平均班組距count80Down"); table.Columns.Add("平均科組距count80Down"); table.Columns.Add("平均校組距count80Down"); table.Columns.Add("加權總分班組距count80Down"); table.Columns.Add("加權總分科組距count80Down"); table.Columns.Add("加權總分校組距count80Down"); table.Columns.Add("加權平均班組距count80Down"); table.Columns.Add("加權平均科組距count80Down"); table.Columns.Add("加權平均校組距count80Down"); table.Columns.Add("類1總分組距count80Down"); table.Columns.Add("類1平均組距count80Down"); table.Columns.Add("類1加權總分組距count80Down"); table.Columns.Add("類1加權平均組距count80Down"); table.Columns.Add("類2總分組距count80Down"); table.Columns.Add("類2平均組距count80Down"); table.Columns.Add("類2加權總分組距count80Down"); table.Columns.Add("類2加權平均組距count80Down");
                    table.Columns.Add("總分班組距count70Down"); table.Columns.Add("總分科組距count70Down"); table.Columns.Add("總分校組距count70Down"); table.Columns.Add("平均班組距count70Down"); table.Columns.Add("平均科組距count70Down"); table.Columns.Add("平均校組距count70Down"); table.Columns.Add("加權總分班組距count70Down"); table.Columns.Add("加權總分科組距count70Down"); table.Columns.Add("加權總分校組距count70Down"); table.Columns.Add("加權平均班組距count70Down"); table.Columns.Add("加權平均科組距count70Down"); table.Columns.Add("加權平均校組距count70Down"); table.Columns.Add("類1總分組距count70Down"); table.Columns.Add("類1平均組距count70Down"); table.Columns.Add("類1加權總分組距count70Down"); table.Columns.Add("類1加權平均組距count70Down"); table.Columns.Add("類2總分組距count70Down"); table.Columns.Add("類2平均組距count70Down"); table.Columns.Add("類2加權總分組距count70Down"); table.Columns.Add("類2加權平均組距count70Down");
                    table.Columns.Add("總分班組距count60Down"); table.Columns.Add("總分科組距count60Down"); table.Columns.Add("總分校組距count60Down"); table.Columns.Add("平均班組距count60Down"); table.Columns.Add("平均科組距count60Down"); table.Columns.Add("平均校組距count60Down"); table.Columns.Add("加權總分班組距count60Down"); table.Columns.Add("加權總分科組距count60Down"); table.Columns.Add("加權總分校組距count60Down"); table.Columns.Add("加權平均班組距count60Down"); table.Columns.Add("加權平均科組距count60Down"); table.Columns.Add("加權平均校組距count60Down"); table.Columns.Add("類1總分組距count60Down"); table.Columns.Add("類1平均組距count60Down"); table.Columns.Add("類1加權總分組距count60Down"); table.Columns.Add("類1加權平均組距count60Down"); table.Columns.Add("類2總分組距count60Down"); table.Columns.Add("類2平均組距count60Down"); table.Columns.Add("類2加權總分組距count60Down"); table.Columns.Add("類2加權平均組距count60Down");
                    table.Columns.Add("總分班組距count50Down"); table.Columns.Add("總分科組距count50Down"); table.Columns.Add("總分校組距count50Down"); table.Columns.Add("平均班組距count50Down"); table.Columns.Add("平均科組距count50Down"); table.Columns.Add("平均校組距count50Down"); table.Columns.Add("加權總分班組距count50Down"); table.Columns.Add("加權總分科組距count50Down"); table.Columns.Add("加權總分校組距count50Down"); table.Columns.Add("加權平均班組距count50Down"); table.Columns.Add("加權平均科組距count50Down"); table.Columns.Add("加權平均校組距count50Down"); table.Columns.Add("類1總分組距count50Down"); table.Columns.Add("類1平均組距count50Down"); table.Columns.Add("類1加權總分組距count50Down"); table.Columns.Add("類1加權平均組距count50Down"); table.Columns.Add("類2總分組距count50Down"); table.Columns.Add("類2平均組距count50Down"); table.Columns.Add("類2加權總分組距count50Down"); table.Columns.Add("類2加權平均組距count50Down");
                    table.Columns.Add("總分班組距count40Down"); table.Columns.Add("總分科組距count40Down"); table.Columns.Add("總分校組距count40Down"); table.Columns.Add("平均班組距count40Down"); table.Columns.Add("平均科組距count40Down"); table.Columns.Add("平均校組距count40Down"); table.Columns.Add("加權總分班組距count40Down"); table.Columns.Add("加權總分科組距count40Down"); table.Columns.Add("加權總分校組距count40Down"); table.Columns.Add("加權平均班組距count40Down"); table.Columns.Add("加權平均科組距count40Down"); table.Columns.Add("加權平均校組距count40Down"); table.Columns.Add("類1總分組距count40Down"); table.Columns.Add("類1平均組距count40Down"); table.Columns.Add("類1加權總分組距count40Down"); table.Columns.Add("類1加權平均組距count40Down"); table.Columns.Add("類2總分組距count40Down"); table.Columns.Add("類2平均組距count40Down"); table.Columns.Add("類2加權總分組距count40Down"); table.Columns.Add("類2加權平均組距count40Down");
                    table.Columns.Add("總分班組距count30Down"); table.Columns.Add("總分科組距count30Down"); table.Columns.Add("總分校組距count30Down"); table.Columns.Add("平均班組距count30Down"); table.Columns.Add("平均科組距count30Down"); table.Columns.Add("平均校組距count30Down"); table.Columns.Add("加權總分班組距count30Down"); table.Columns.Add("加權總分科組距count30Down"); table.Columns.Add("加權總分校組距count30Down"); table.Columns.Add("加權平均班組距count30Down"); table.Columns.Add("加權平均科組距count30Down"); table.Columns.Add("加權平均校組距count30Down"); table.Columns.Add("類1總分組距count30Down"); table.Columns.Add("類1平均組距count30Down"); table.Columns.Add("類1加權總分組距count30Down"); table.Columns.Add("類1加權平均組距count30Down"); table.Columns.Add("類2總分組距count30Down"); table.Columns.Add("類2平均組距count30Down"); table.Columns.Add("類2加權總分組距count30Down"); table.Columns.Add("類2加權平均組距count30Down");
                    table.Columns.Add("總分班組距count20Down"); table.Columns.Add("總分科組距count20Down"); table.Columns.Add("總分校組距count20Down"); table.Columns.Add("平均班組距count20Down"); table.Columns.Add("平均科組距count20Down"); table.Columns.Add("平均校組距count20Down"); table.Columns.Add("加權總分班組距count20Down"); table.Columns.Add("加權總分科組距count20Down"); table.Columns.Add("加權總分校組距count20Down"); table.Columns.Add("加權平均班組距count20Down"); table.Columns.Add("加權平均科組距count20Down"); table.Columns.Add("加權平均校組距count20Down"); table.Columns.Add("類1總分組距count20Down"); table.Columns.Add("類1平均組距count20Down"); table.Columns.Add("類1加權總分組距count20Down"); table.Columns.Add("類1加權平均組距count20Down"); table.Columns.Add("類2總分組距count20Down"); table.Columns.Add("類2平均組距count20Down"); table.Columns.Add("類2加權總分組距count20Down"); table.Columns.Add("類2加權平均組距count20Down");
                    table.Columns.Add("總分班組距count10Down"); table.Columns.Add("總分科組距count10Down"); table.Columns.Add("總分校組距count10Down"); table.Columns.Add("平均班組距count10Down"); table.Columns.Add("平均科組距count10Down"); table.Columns.Add("平均校組距count10Down"); table.Columns.Add("加權總分班組距count10Down"); table.Columns.Add("加權總分科組距count10Down"); table.Columns.Add("加權總分校組距count10Down"); table.Columns.Add("加權平均班組距count10Down"); table.Columns.Add("加權平均科組距count10Down"); table.Columns.Add("加權平均校組距count10Down"); table.Columns.Add("類1總分組距count10Down"); table.Columns.Add("類1平均組距count10Down"); table.Columns.Add("類1加權總分組距count10Down"); table.Columns.Add("類1加權平均組距count10Down"); table.Columns.Add("類2總分組距count10Down"); table.Columns.Add("類2平均組距count10Down"); table.Columns.Add("類2加權總分組距count10Down"); table.Columns.Add("類2加權平均組距count10Down");
                    #endregion
                    #endregion
                    //宣告產生的報表
                    Aspose.Words.Document document = new Aspose.Words.Document();
                    //用一個BackgroundWorker包起來
                    System.ComponentModel.BackgroundWorker bkw = new System.ComponentModel.BackgroundWorker();
                    bkw.WorkerReportsProgress = true;
                    System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 班級評量成績單產生 S");
                    bkw.ProgressChanged += delegate(object sender, System.ComponentModel.ProgressChangedEventArgs e)
                    {
                         FISCA.Presentation.MotherForm.SetStatusBarMessage(" 班級評量成績單產生中", e.ProgressPercentage);
                         System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 班級評量成績單產生 " + e.ProgressPercentage);
                    };
                    bkw.RunWorkerCompleted += delegate
                    {
                         System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 班級評量成績單產生 E");
                         string err = "下列班級因成績項目超過樣板支援上限，\n或者班級學生數超過樣板支援上限，\n超出部分科目成績無法印出，建議調整樣板內容。";
                         if (overflowRecords.Count > 0)
                         {
                              foreach (ClassRecord classRec in overflowRecords)
                              {
                                   err += "\n" + classRec.ClassName;
                              }
                         }
                         if (exc != null)
                         {
                              throw new Exception("產生班級評量成績單發生錯誤", exc);
                         }
                         #region 儲存檔案
                         string inputReportName = "班級評量成績單";
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
                              document.Save(path, Aspose.Words.SaveFormat.Doc);
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
                                        document.Save(sd.FileName, Aspose.Words.SaveFormat.Doc);

                                   }
                                   catch
                                   {
                                        FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                                        return;
                                   }
                              }
                         }
                         #endregion
                         FISCA.Presentation.MotherForm.SetStatusBarMessage("班級評量成績單產生完成。", 100);
                         if (overflowRecords.Count > 0)
                              MessageBox.Show(err);
                    };
                    bkw.DoWork += delegate(object sender, System.ComponentModel.DoWorkEventArgs e)
                    {


                         #region 偷跑取得考試成績
                         Dictionary<string, Dictionary<string, Dictionary<string, ExamScoreInfo>>> studentExamSores = new Dictionary<string, Dictionary<string, Dictionary<string, ExamScoreInfo>>>();
                         Dictionary<string, Dictionary<string, ExamScoreInfo>> studentRefExamSores = new Dictionary<string, Dictionary<string, ExamScoreInfo>>();
                         ManualResetEvent scoreReady = new ManualResetEvent(false);
                         ManualResetEvent elseReady = new ManualResetEvent(false);

                         bkw.ReportProgress(3);
                         new Thread(new ThreadStart(delegate
                         {
                              // 取得學生學期科目成績
                              int sSchoolYear, sSemester;
                              int.TryParse(conf.SchoolYear, out sSchoolYear);
                              int.TryParse(conf.Semester, out sSemester);
                              #region 整理學生定期評量成績
                              #region 篩選課程學年度、學期、科目取得有可能有需要的資料
                              List<CourseRecord> targetCourseList = new List<CourseRecord>();
                              try
                              {
                                   foreach (var courseRecord in accessHelper.CourseHelper.GetAllCourse(sSchoolYear, sSemester))
                                   {
                                        //用科目濾出可能有用到的課程
                                        if (conf.PrintSubjectList.Contains(courseRecord.Subject)
                                            || conf.TagRank1SubjectList.Contains(courseRecord.Subject)
                                            || conf.TagRank2SubjectList.Contains(courseRecord.Subject))
                                             targetCourseList.Add(courseRecord);
                                   }
                              }
                              catch (Exception exception)
                              {
                                   exc = exception;
                              }
                              #endregion
                              try
                              {
                                   if (conf.ExamRecord != null || conf.RefenceExamRecord != null)
                                   {
                                        accessHelper.CourseHelper.FillExam(targetCourseList);
                                        var tcList = new List<CourseRecord>();
                                        var totalList = new List<CourseRecord>();
                                        foreach (var courseRec in targetCourseList)
                                        {
                                             if (conf.ExamRecord != null && courseRec.ExamList.Contains(conf.ExamRecord.Name))
                                             {
                                                  tcList.Add(courseRec);
                                                  totalList.Add(courseRec);
                                             }
                                             if (tcList.Count == 180)
                                             {
                                                  accessHelper.CourseHelper.FillStudentAttend(tcList);
                                                  accessHelper.CourseHelper.FillExamScore(tcList);
                                                  tcList.Clear();
                                             }
                                        }
                                        accessHelper.CourseHelper.FillStudentAttend(tcList);
                                        accessHelper.CourseHelper.FillExamScore(tcList);
                                        foreach (var courseRecord in totalList)
                                        {
                                             #region 整理本次定期評量成績
                                             if (conf.ExamRecord != null && courseRecord.ExamList.Contains(conf.ExamRecord.Name))
                                             {
                                                  foreach (var attendStudent in courseRecord.StudentAttendList)
                                                  {
                                                       if (!studentExamSores.ContainsKey(attendStudent.StudentID)) studentExamSores.Add(attendStudent.StudentID, new Dictionary<string, Dictionary<string, ExamScoreInfo>>());
                                                       if (!studentExamSores[attendStudent.StudentID].ContainsKey(courseRecord.Subject)) studentExamSores[attendStudent.StudentID].Add(courseRecord.Subject, new Dictionary<string, ExamScoreInfo>());
                                                       studentExamSores[attendStudent.StudentID][courseRecord.Subject].Add("" + attendStudent.CourseID, null);
                                                  }
                                                  foreach (var examScoreRec in courseRecord.ExamScoreList)
                                                  {
                                                       if (examScoreRec.ExamName == conf.ExamRecord.Name)
                                                       {
                                                            studentExamSores[examScoreRec.StudentID][courseRecord.Subject]["" + examScoreRec.CourseID] = examScoreRec;
                                                       }
                                                  }
                                             }
                                             #endregion
                                             #region 整理前次定期評量成績
                                             if (conf.RefenceExamRecord != null && courseRecord.ExamList.Contains(conf.RefenceExamRecord.Name))
                                             {
                                                  foreach (var examScoreRec in courseRecord.ExamScoreList)
                                                  {
                                                       if (examScoreRec.ExamName == conf.RefenceExamRecord.Name)
                                                       {
                                                            if (!studentRefExamSores.ContainsKey(examScoreRec.StudentID))
                                                                 studentRefExamSores.Add(examScoreRec.StudentID, new Dictionary<string, ExamScoreInfo>());
                                                            studentRefExamSores[examScoreRec.StudentID].Add("" + examScoreRec.CourseID, examScoreRec);
                                                       }
                                                  }
                                             }
                                             #endregion
                                        }
                                   }
                              }
                              catch (Exception exception)
                              {
                                   exc = exception;
                              }
                              finally
                              {
                                   scoreReady.Set();
                              }
                              #endregion
                              #region 整理學生學期、學年成績
                              try
                              {
                                   accessHelper.StudentHelper.FillAttendance(selectedStudents);
                                   accessHelper.StudentHelper.FillReward(selectedStudents);
                              }
                              catch (Exception exception)
                              {
                                   exc = exception;
                              }
                              finally
                              {
                                   elseReady.Set();
                              }
                              #endregion
                         })).Start();
                         #endregion
                         try
                         {
                              string key = "";
                              bkw.ReportProgress(10);
                              #region 整理同年級學生
                              //整理選取學生的年級
                              Dictionary<string, List<StudentRecord>> gradeyearStudents = new Dictionary<string, List<StudentRecord>>();
                              foreach (ClassRecord classRec in selectedClasses)
                              {
                                   string grade = "";
                                   grade = "" + classRec.GradeYear;
                                   if (!gradeyearStudents.ContainsKey(grade))
                                        gradeyearStudents.Add(grade, new List<StudentRecord>());
                                   foreach (var studentRec in classRec.Students)
                                   {
                                        gradeyearStudents[grade].Add(studentRec);
                                   }
                              }
                              foreach (var classRec in accessHelper.ClassHelper.GetAllClass())
                              {
                                   if (!selectedClasses.Contains(classRec) && gradeyearStudents.ContainsKey("" + classRec.GradeYear))
                                   {
                                        //用班級去取出可能有相關的學生
                                        foreach (var studentRec in classRec.Students)
                                        {
                                             string grade = "";
                                             if (studentRec.RefClass != null)
                                                  grade = "" + studentRec.RefClass.GradeYear;
                                             if (!gradeyearStudents[grade].Contains(studentRec))
                                                  gradeyearStudents[grade].Add(studentRec);
                                        }
                                   }
                              }
                              #endregion
                              bkw.ReportProgress(15);
                              #region 取得學生類別
                              Dictionary<string, List<K12.Data.StudentTagRecord>> studentTags = new Dictionary<string, List<K12.Data.StudentTagRecord>>();
                              List<string> list = new List<string>();
                              foreach (var sRecs in gradeyearStudents.Values)
                              {
                                   foreach (var stuRec in sRecs)
                                   {
                                        list.Add(stuRec.StudentID);
                                   }
                              }
                              foreach (var tag in K12.Data.StudentTag.SelectByStudentIDs(list))
                              {
                                   if (!studentTags.ContainsKey(tag.RefStudentID))
                                        studentTags.Add(tag.RefStudentID, new List<K12.Data.StudentTagRecord>());
                                   studentTags[tag.RefStudentID].Add(tag);
                              }
                              #endregion
                              bkw.ReportProgress(20);
                              //等到成績載完
                              scoreReady.WaitOne();
                              bkw.ReportProgress(35);
                              int progressCount = 0;

                              #region 計算總分及各項目排名
                              Dictionary<string, string> studentTag1Group = new Dictionary<string, string>();
                              Dictionary<string, string> studentTag2Group = new Dictionary<string, string>();
                              Dictionary<string, List<decimal>> ranks = new Dictionary<string, List<decimal>>();
                              Dictionary<string, List<string>> rankStudents = new Dictionary<string, List<string>>();
                              Dictionary<string, decimal> studentPrintSubjectSum = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentTag1SubjectSum = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentTag2SubjectSum = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentPrintSubjectAvg = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentTag1SubjectAvg = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentTag2SubjectAvg = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentPrintSubjectSumW = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentTag1SubjectSumW = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentTag2SubjectSumW = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentPrintSubjectAvgW = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentTag1SubjectAvgW = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> studentTag2SubjectAvgW = new Dictionary<string, decimal>();
                              Dictionary<string, decimal> analytics = new Dictionary<string, decimal>();
                              int total = 0;
                              foreach (var gss in gradeyearStudents.Values)
                              {
                                   total += gss.Count;
                              }
                              bkw.ReportProgress(40);
                              foreach (string gradeyear in gradeyearStudents.Keys)
                              {
                                   //找出全年級學生
                                   foreach (var studentRec in gradeyearStudents[gradeyear])
                                   {
                                        string studentID = studentRec.StudentID;
                                        bool rank = true;
                                        string tag1ID = "";
                                        string tag2ID = "";
                                        #region 分析學生所屬類別
                                        if (studentTags.ContainsKey(studentID))
                                        {
                                             foreach (var tag in studentTags[studentID])
                                             {
                                                  #region 判斷學生是否屬於不排名類別
                                                  if (conf.RankFilterTagList.Contains(tag.RefTagID))
                                                  {
                                                       rank = false;
                                                  }
                                                  #endregion
                                                  #region 判斷學生在類別排名1中所屬的類別
                                                  if (tag1ID == "" && conf.TagRank1TagList.Contains(tag.RefTagID))
                                                  {
                                                       tag1ID = tag.RefTagID;
                                                       studentTag1Group.Add(studentID, tag1ID);
                                                  }
                                                  #endregion
                                                  #region 判斷學生在類別排名2中所屬的類別
                                                  if (tag2ID == "" && conf.TagRank2TagList.Contains(tag.RefTagID))
                                                  {
                                                       tag2ID = tag.RefTagID;
                                                       studentTag2Group.Add(studentID, tag2ID);
                                                  }
                                                  #endregion
                                             }
                                        }
                                        #endregion
                                        bool summaryRank = true;
                                        bool tag1SummaryRank = true;
                                        bool tag2SummaryRank = true;
                                        if (studentExamSores.ContainsKey(studentID))
                                        {
                                             decimal printSubjectSum = 0;
                                             int printSubjectCount = 0;
                                             decimal tag1SubjectSum = 0;
                                             int tag1SubjectCount = 0;
                                             decimal tag2SubjectSum = 0;
                                             int tag2SubjectCount = 0;
                                             decimal printSubjectSumW = 0;
                                             decimal printSubjectCreditSum = 0;
                                             decimal tag1SubjectSumW = 0;
                                             decimal tag1SubjectCreditSum = 0;
                                             decimal tag2SubjectSumW = 0;
                                             decimal tag2SubjectCreditSum = 0;
                                             foreach (var subjectName in studentExamSores[studentID].Keys)
                                             {
                                                  if (conf.PrintSubjectList.Contains(subjectName))
                                                  {
                                                       #region 是列印科目
                                                       foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                                       {
                                                            if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                            {
                                                                 printSubjectSum += sceTakeRecord.ExamScore;//計算總分
                                                                 printSubjectCount++;
                                                                 //計算加權總分
                                                                 printSubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                                 printSubjectCreditSum += sceTakeRecord.CreditDec();
                                                                 if (rank && sceTakeRecord.Status == "一般")//不在過濾名單且為一般生才做排名
                                                                 {
                                                                      if (sceTakeRecord.RefClass != null)
                                                                      {
                                                                           //各科目班排名
                                                                           key = "班排名" + sceTakeRecord.RefClass.ClassID + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                           if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                                           if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                                           ranks[key].Add(sceTakeRecord.ExamScore);
                                                                           rankStudents[key].Add(studentID);
                                                                      }
                                                                      if (sceTakeRecord.Department != "")
                                                                      {
                                                                           //各科目科排名
                                                                           key = "科排名" + sceTakeRecord.Department + "^^^" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                           if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                                           if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                                           ranks[key].Add(sceTakeRecord.ExamScore);
                                                                           rankStudents[key].Add(studentID);
                                                                      }
                                                                      //各科目全校排名
                                                                      key = "全校排名" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                      if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                                      if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                                      ranks[key].Add(sceTakeRecord.ExamScore);
                                                                      rankStudents[key].Add(studentID);
                                                                 }
                                                            }
                                                            else
                                                            {
                                                                 summaryRank = false;
                                                            }
                                                       }
                                                       #endregion
                                                  }
                                                  if (tag1ID != "" && conf.TagRank1SubjectList.Contains(subjectName))
                                                  {
                                                       #region 有Tag1且是排名科目
                                                       foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                                       {
                                                            if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                            {
                                                                 tag1SubjectSum += sceTakeRecord.ExamScore;//計算總分
                                                                 tag1SubjectCount++;
                                                                 //計算加權總分
                                                                 tag1SubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                                 tag1SubjectCreditSum += sceTakeRecord.CreditDec();
                                                                 //各科目類別1排名
                                                                 if (rank && sceTakeRecord.Status == "一般")//不在過濾名單且為一般生才做排名
                                                                 {
                                                                      if (conf.PrintSubjectList.Contains(subjectName))//是列印科目才算科目排名                                                
                                                                      {
                                                                           key = "類別1排名" + tag1ID + "^^^" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                           if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                                           if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                                           ranks[key].Add(sceTakeRecord.ExamScore);
                                                                           rankStudents[key].Add(studentID);
                                                                      }
                                                                 }
                                                            }
                                                            else
                                                            {
                                                                 tag1SummaryRank = false;
                                                            }
                                                       }
                                                       #endregion
                                                  }
                                                  if (tag2ID != "" && conf.TagRank2SubjectList.Contains(subjectName))
                                                  {
                                                       #region 有Tag2且是排名科目
                                                       foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                                       {
                                                            if (sceTakeRecord != null && sceTakeRecord.SpecialCase == "")
                                                            {
                                                                 tag2SubjectSum += sceTakeRecord.ExamScore;//計算總分
                                                                 tag2SubjectCount++;
                                                                 //計算加權總分
                                                                 tag2SubjectSumW += sceTakeRecord.ExamScore * sceTakeRecord.CreditDec();
                                                                 tag2SubjectCreditSum += sceTakeRecord.CreditDec();
                                                                 //各科目類別2排名
                                                                 if (rank && sceTakeRecord.Status == "一般")//不在過濾名單且為一般生才做排名
                                                                 {
                                                                      if (conf.PrintSubjectList.Contains(subjectName))//是列印科目才算科目排名                                                
                                                                      {
                                                                           key = "類別2排名" + tag2ID + "^^^" + gradeyear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                           if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                                           if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                                           ranks[key].Add(sceTakeRecord.ExamScore);
                                                                           rankStudents[key].Add(studentID);
                                                                      }
                                                                 }
                                                            }
                                                            else
                                                            {
                                                                 tag2SummaryRank = false;
                                                            }
                                                       }
                                                       #endregion
                                                  }

                                             }
                                             if (printSubjectCount > 0)
                                             {
                                                  #region 有列印科目處理加總成績
                                                  //總分
                                                  studentPrintSubjectSum.Add(studentID, printSubjectSum);
                                                  //平均四捨五入至小數點第二位
                                                  studentPrintSubjectAvg.Add(studentID, Math.Round(printSubjectSum / printSubjectCount, 2, MidpointRounding.AwayFromZero));
                                                  if (rank && studentRec.Status == "一般" && summaryRank == true)//不在過濾名單且沒有特殊成績狀況且為一般生才做排名
                                                  {
                                                       //總分班排名
                                                       key = "總分班排名" + studentRec.RefClass.ClassID;
                                                       if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                       if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                       ranks[key].Add(printSubjectSum);
                                                       rankStudents[key].Add(studentID);
                                                       //總分科排名
                                                       key = "總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                                       if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                       if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                       ranks[key].Add(printSubjectSum);
                                                       rankStudents[key].Add(studentID);
                                                       //總分全校排名
                                                       key = "總分全校排名" + gradeyear;
                                                       if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                       if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                       ranks[key].Add(printSubjectSum);
                                                       rankStudents[key].Add(studentID);
                                                       //平均班排名
                                                       key = "平均班排名" + studentRec.RefClass.ClassID;
                                                       if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                       if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                       ranks[key].Add(Math.Round(printSubjectSum / printSubjectCount, 2, MidpointRounding.AwayFromZero));
                                                       rankStudents[key].Add(studentID);
                                                       //平均科排名
                                                       key = "平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                                       if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                       if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                       ranks[key].Add(Math.Round(printSubjectSum / printSubjectCount, 2, MidpointRounding.AwayFromZero));
                                                       rankStudents[key].Add(studentID);
                                                       //平均全校排名
                                                       key = "平均全校排名" + gradeyear;
                                                       if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                       if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                       ranks[key].Add(Math.Round(printSubjectSum / printSubjectCount, 2, MidpointRounding.AwayFromZero));
                                                       rankStudents[key].Add(studentID);
                                                  }
                                                  #endregion
                                                  if (printSubjectCreditSum > 0)
                                                  {
                                                       #region 有總學分數處理加總
                                                       //加權總分
                                                       studentPrintSubjectSumW.Add(studentID, printSubjectSumW);
                                                       //加權平均四捨五入至小數點第二位
                                                       studentPrintSubjectAvgW.Add(studentID, Math.Round(printSubjectSumW / printSubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                       if (rank && studentRec.Status == "一般" && summaryRank == true)//不在過濾名單且為一般生才做排名
                                                       {
                                                            //加權總分班排名
                                                            key = "加權總分班排名" + studentRec.RefClass.ClassID;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(printSubjectSumW);
                                                            rankStudents[key].Add(studentID);
                                                            //加權總分科排名
                                                            key = "加權總分科排名" + studentRec.Department + "^^^" + gradeyear;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(printSubjectSumW);
                                                            rankStudents[key].Add(studentID);
                                                            //加權總分全校排名
                                                            key = "加權總分全校排名" + gradeyear;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(printSubjectSumW);
                                                            rankStudents[key].Add(studentID);
                                                            //加權平均班排名
                                                            key = "加權平均班排名" + studentRec.RefClass.ClassID;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(Math.Round(printSubjectSumW / printSubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                            rankStudents[key].Add(studentID);
                                                            //加權平均科排名
                                                            key = "加權平均科排名" + studentRec.Department + "^^^" + gradeyear;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(Math.Round(printSubjectSumW / printSubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                            rankStudents[key].Add(studentID);
                                                            //加權平均全校排名
                                                            key = "加權平均全校排名" + gradeyear;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(Math.Round(printSubjectSumW / printSubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                            rankStudents[key].Add(studentID);
                                                       }
                                                       #endregion
                                                  }
                                             }
                                             //類別1總分平均排名
                                             if (tag1SubjectCount > 0)
                                             {
                                                  //總分
                                                  studentTag1SubjectSum.Add(studentID, tag1SubjectSum);
                                                  //平均四捨五入至小數點第二位
                                                  studentTag1SubjectAvg.Add(studentID, Math.Round(tag1SubjectSum / tag1SubjectCount, 2, MidpointRounding.AwayFromZero));
                                                  if (rank && studentRec.Status == "一般" && tag1SummaryRank == true)//不在過濾名單且為一般生才做排名
                                                  {
                                                       key = "類別1總分排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                                       if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                       if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                       ranks[key].Add(tag1SubjectSum);
                                                       rankStudents[key].Add(studentID);

                                                       key = "類別1平均排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                                       if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                       if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                       ranks[key].Add(Math.Round(tag1SubjectSum / tag1SubjectCount, 2, MidpointRounding.AwayFromZero));
                                                       rankStudents[key].Add(studentID);
                                                  }
                                                  //類別1加權總分平均排名
                                                  if (tag1SubjectCreditSum > 0)
                                                  {
                                                       studentTag1SubjectSumW.Add(studentID, tag1SubjectSumW);
                                                       studentTag1SubjectAvgW.Add(studentID, Math.Round(tag1SubjectSumW / tag1SubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                       if (rank && studentRec.Status == "一般" && tag1SummaryRank == true)//不在過濾名單且為一般生才做排名
                                                       {
                                                            key = "類別1加權總分排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(tag1SubjectSumW);
                                                            rankStudents[key].Add(studentID);

                                                            key = "類別1加權平均排名" + "^^^" + gradeyear + "^^^" + tag1ID;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(Math.Round(tag1SubjectSumW / tag1SubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                            rankStudents[key].Add(studentID);
                                                       }
                                                  }
                                             }
                                             //類別2總分平均排名
                                             if (tag2SubjectCount > 0)
                                             {
                                                  //總分
                                                  studentTag2SubjectSum.Add(studentID, tag2SubjectSum);
                                                  //平均四捨五入至小數點第二位
                                                  studentTag2SubjectAvg.Add(studentID, Math.Round(tag2SubjectSum / tag2SubjectCount, 2, MidpointRounding.AwayFromZero));
                                                  if (rank && studentRec.Status == "一般" && tag2SummaryRank == true)//不在過濾名單且為一般生才做排名
                                                  {
                                                       key = "類別2總分排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                                       if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                       if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                       ranks[key].Add(tag2SubjectSum);
                                                       rankStudents[key].Add(studentID);
                                                       key = "類別2平均排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                                       if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                       if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                       ranks[key].Add(Math.Round(tag2SubjectSum / tag2SubjectCount, 2, MidpointRounding.AwayFromZero));
                                                       rankStudents[key].Add(studentID);
                                                  }
                                                  //類別2加權總分平均排名
                                                  if (tag2SubjectCreditSum > 0)
                                                  {
                                                       studentTag2SubjectSumW.Add(studentID, tag2SubjectSumW);
                                                       studentTag2SubjectAvgW.Add(studentID, Math.Round(tag2SubjectSumW / tag2SubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                       if (rank && studentRec.Status == "一般" && tag2SummaryRank == true)//不在過濾名單且為一般生才做排名
                                                       {
                                                            key = "類別2加權總分排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(tag2SubjectSumW);
                                                            rankStudents[key].Add(studentID);

                                                            key = "類別2加權平均排名" + "^^^" + gradeyear + "^^^" + tag2ID;
                                                            if (!ranks.ContainsKey(key)) ranks.Add(key, new List<decimal>());
                                                            if (!rankStudents.ContainsKey(key)) rankStudents.Add(key, new List<string>());
                                                            ranks[key].Add(Math.Round(tag2SubjectSumW / tag2SubjectCreditSum, 2, MidpointRounding.AwayFromZero));
                                                            rankStudents[key].Add(studentID);
                                                       }
                                                  }
                                             }
                                        }
                                        progressCount++;
                                        bkw.ReportProgress(40 + progressCount * 30 / total);
                                   }
                              }
                              foreach (var k in ranks.Keys)
                              {
                                   var rankscores = ranks[k];
                                   //排序
                                   rankscores.Sort();
                                   rankscores.Reverse();
                                   //高均標、組距
                                   if (rankscores.Count > 0)
                                   {
                                        #region 算高標的中點
                                        int middleIndex = 0;
                                        int count = 1;
                                        var score = rankscores[0];
                                        while (rankscores.Count > middleIndex)
                                        {
                                             if (score != rankscores[middleIndex])
                                             {
                                                  if (count * 2 >= rankscores.Count) break;
                                                  score = rankscores[middleIndex];
                                             }
                                             middleIndex++;
                                             count++;
                                        }
                                        if (rankscores.Count == middleIndex)
                                        {
                                             middleIndex--;
                                             count--;
                                        }
                                        #endregion
                                        analytics.Add(k + "^^^高標", Math.Round(rankscores.GetRange(0, count).Average(), 2, MidpointRounding.AwayFromZero));
                                        analytics.Add(k + "^^^均標", Math.Round(rankscores.Average(), 2, MidpointRounding.AwayFromZero));
                                        #region 算低標的中點
                                        middleIndex = rankscores.Count - 1;
                                        count = 1;
                                        score = rankscores[middleIndex];
                                        while (middleIndex >= 0)
                                        {
                                             if (score != rankscores[middleIndex])
                                             {
                                                  if (count * 2 >= rankscores.Count) break;
                                                  score = rankscores[middleIndex];
                                             }
                                             middleIndex--;
                                             count++;
                                        }
                                        if (middleIndex < 0)
                                        {
                                             middleIndex++;
                                             count--;
                                        }
                                        #endregion
                                        analytics.Add(k + "^^^低標", Math.Round(rankscores.GetRange(middleIndex, count).Average(), 2, MidpointRounding.AwayFromZero));
                                        //Compute the Average      
                                        var avg = (double)rankscores.Average();
                                        //Perform the Sum of (value-avg)_2_2      
                                        var sum = (double)rankscores.Sum(d => Math.Pow((double)d - avg, 2));
                                        //Put it all together      
                                        analytics.Add(k + "^^^標準差", Math.Round((decimal)Math.Sqrt((sum) / rankscores.Count()), 2, MidpointRounding.AwayFromZero));
                                   }
                                   #region 計算級距
                                   int count90 = 0, count80 = 0, count70 = 0, count60 = 0, count50 = 0, count40 = 0, count30 = 0, count20 = 0, count10 = 0;
                                   int count100Up = 0, count90Up = 0, count80Up = 0, count70Up = 0, count60Up = 0, count50Up = 0, count40Up = 0, count30Up = 0, count20Up = 0, count10Up = 0;
                                   int count90Down = 0, count80Down = 0, count70Down = 0, count60Down = 0, count50Down = 0, count40Down = 0, count30Down = 0, count20Down = 0, count10Down = 0;
                                   foreach (var score in rankscores)
                                   {
                                        if (score >= 100)
                                             count100Up++;
                                        else if (score >= 90)
                                             count90++;
                                        else if (score >= 80)
                                             count80++;
                                        else if (score >= 70)
                                             count70++;
                                        else if (score >= 60)
                                             count60++;
                                        else if (score >= 50)
                                             count50++;
                                        else if (score >= 40)
                                             count40++;
                                        else if (score >= 30)
                                             count30++;
                                        else if (score >= 20)
                                             count20++;
                                        else if (score >= 10)
                                             count10++;
                                        else
                                             count10Down++;
                                   }
                                   count90Up = count100Up + count90;
                                   count80Up = count90Up + count80;
                                   count70Up = count80Up + count70;
                                   count60Up = count70Up + count60;
                                   count50Up = count60Up + count50;
                                   count40Up = count50Up + count40;
                                   count30Up = count40Up + count30;
                                   count20Up = count30Up + count20;
                                   count10Up = count20Up + count10;

                                   count20Down = count10Down + count10;
                                   count30Down = count20Down + count20;
                                   count40Down = count30Down + count30;
                                   count50Down = count40Down + count40;
                                   count60Down = count50Down + count50;
                                   count70Down = count60Down + count60;
                                   count80Down = count70Down + count70;
                                   count90Down = count80Down + count80;

                                   analytics.Add(k + "^^^count90", count90);
                                   analytics.Add(k + "^^^count80", count80);
                                   analytics.Add(k + "^^^count70", count70);
                                   analytics.Add(k + "^^^count60", count60);
                                   analytics.Add(k + "^^^count50", count50);
                                   analytics.Add(k + "^^^count40", count40);
                                   analytics.Add(k + "^^^count30", count30);
                                   analytics.Add(k + "^^^count20", count20);
                                   analytics.Add(k + "^^^count10", count10);
                                   analytics.Add(k + "^^^count100Up", count100Up);
                                   analytics.Add(k + "^^^count90Up", count90Up);
                                   analytics.Add(k + "^^^count80Up", count80Up);
                                   analytics.Add(k + "^^^count70Up", count70Up);
                                   analytics.Add(k + "^^^count60Up", count60Up);
                                   analytics.Add(k + "^^^count50Up", count50Up);
                                   analytics.Add(k + "^^^count40Up", count40Up);
                                   analytics.Add(k + "^^^count30Up", count30Up);
                                   analytics.Add(k + "^^^count20Up", count20Up);
                                   analytics.Add(k + "^^^count10Up", count10Up);
                                   analytics.Add(k + "^^^count90Down", count90Down);
                                   analytics.Add(k + "^^^count80Down", count80Down);
                                   analytics.Add(k + "^^^count70Down", count70Down);
                                   analytics.Add(k + "^^^count60Down", count60Down);
                                   analytics.Add(k + "^^^count50Down", count50Down);
                                   analytics.Add(k + "^^^count40Down", count40Down);
                                   analytics.Add(k + "^^^count30Down", count30Down);
                                   analytics.Add(k + "^^^count20Down", count20Down);
                                   analytics.Add(k + "^^^count10Down", count10Down);
                                   #endregion
                              }
                              #endregion
                              //參考成績載完
                              elseReady.WaitOne();
                              bkw.ReportProgress(70);
                              progressCount = 0;
                              #region 填入資料表
                              foreach (ClassRecord classRec in selectedClasses)
                              {
                                   DataRow row = table.NewRow();
                                   List<string> tag1List = new List<string>();
                                   List<string> tag2List = new List<string>();
                                   Dictionary<string, Dictionary<string, List<string>>> classSubjects = new Dictionary<string, Dictionary<string, List<string>>>();
                                   string gradeYear = classRec.GradeYear;
                                   #region 基本資料
                                   row["學校名稱"] = SmartSchool.Customization.Data.SystemInformation.SchoolChineseName;
                                   row["學校地址"] = SmartSchool.Customization.Data.SystemInformation.Address;
                                   row["學校電話"] = SmartSchool.Customization.Data.SystemInformation.Telephone;
                                   row["科別名稱"] = classRec.Department;
                                   row["試別"] = conf.ExamRecord.Name;

                                   row["學年度"] = conf.SchoolYear;
                                   row["學期"] = conf.Semester;
                                   row["班級科別名稱"] = classRec.Department;
                                   row["班級"] = classRec.ClassName;
                                   row["定期評量"] = conf.ExamRecord.Name;
                                   row["班導師"] = classRec.RefTeacher == null ? "" : classRec.RefTeacher.TeacherName;
                                   int ClassStuNum = 0;
                                   foreach (StudentRecord stuRec in classRec.Students)
                                   {
                                        string studentID = stuRec.StudentID;
                                        ClassStuNum++;
                                        if (ClassStuNum > conf.StudentLimit)
                                        {
                                             if (!overflowRecords.Contains(classRec))
                                                  overflowRecords.Add(classRec);
                                             break;
                                        }
                                        row["座號" + ClassStuNum] = stuRec.SeatNo;
                                        row["學號" + ClassStuNum] = stuRec.StudentNumber;
                                        row["姓名" + ClassStuNum] = stuRec.StudentName;
                                        //整理這個班的學生所屬類別
                                        if (studentTag1Group.ContainsKey(studentID) && !tag1List.Contains(studentTag1Group[studentID]))
                                        {
                                             tag1List.Add(studentTag1Group[studentID]);
                                             foreach (var tag in studentTags[studentID])
                                             {
                                                  if (tag.RefTagID == studentTag1Group[studentID])
                                                  {
                                                       row["類別排名1"] += (("" + row["類別排名1"]) == "" ? "" : "/") + tag.Name;
                                                  }
                                             }
                                        }
                                        if (studentTag2Group.ContainsKey(studentID) && !tag2List.Contains(studentTag2Group[studentID]))
                                        {
                                             tag2List.Add(studentTag2Group[studentID]);
                                             foreach (var tag in studentTags[studentID])
                                             {
                                                  if (tag.RefTagID == studentTag2Group[studentID])
                                                  {
                                                       row["類別排名2"] += (("" + row["類別排名2"]) == "" ? "" : "/") + tag.Name;
                                                  }
                                             }
                                        }

                                        #region 整理班級中列印科目
                                        if (studentExamSores.ContainsKey(stuRec.StudentID))
                                        {
                                             foreach (var subjectName in studentExamSores[studentID].Keys)
                                             {
                                                  foreach (var courseID in studentExamSores[studentID][subjectName].Keys)
                                                  {
                                                       string subjectLevel = accessHelper.CourseHelper.GetCourse(courseID)[0].SubjectLevel;
                                                       string credit = "" + accessHelper.CourseHelper.GetCourse(courseID)[0].CreditDec();
                                                       if (conf.PrintSubjectList.Contains(subjectName))
                                                       {
                                                            if (!classSubjects.ContainsKey(subjectName))
                                                                 classSubjects.Add(subjectName, new Dictionary<string, List<string>>());

                                                            if (!classSubjects[subjectName].ContainsKey(subjectLevel))
                                                                 classSubjects[subjectName].Add(subjectLevel, new List<string>());

                                                            if (!classSubjects[subjectName][subjectLevel].Contains(credit))
                                                                 classSubjects[subjectName][subjectLevel].Add(credit);

                                                       }
                                                  }
                                             }
                                        }
                                        #endregion
                                   }
                                   #endregion
                                   #region 各科成績資料
                                   #region 整理列印順序
                                   List<string> subjectNameList = new List<string>(classSubjects.Keys);
                                   subjectNameList.Sort(new StringComparer("國文"
                                                   , "英文"
                                                   , "數學"
                                                   , "理化"
                                                   , "生物"
                                                   , "社會"
                                                   , "物理"
                                                   , "化學"
                                                   , "歷史"
                                                   , "地理"
                                                   , "公民"));
                                   #endregion
                                   int subjectIndex = 1;
                                   // 學期科目與定期評量
                                   foreach (string subjectName in subjectNameList)
                                   {
                                        foreach (string subjectLevel in classSubjects[subjectName].Keys)
                                        {
                                             decimal level;
                                             decimal? subjectNumber = null;
                                             subjectNumber = decimal.TryParse(subjectLevel, out level) ? (decimal?)level : null;
                                             if (subjectIndex <= conf.SubjectLimit)
                                             {
                                                  row["科目名稱" + subjectIndex] = subjectName + GetNumber(subjectNumber);
                                                  row["學分數" + subjectIndex] = "";
                                                  foreach (string credit in classSubjects[subjectName][subjectLevel])
                                                  {
                                                       row["學分數" + subjectIndex] += (("" + row["學分數" + subjectIndex]) == "" ? "" : ",") + credit;
                                                  }
                                                  // 檢查畫面上定期評量列印科目
                                                  if (conf.PrintSubjectList.Contains(subjectName))
                                                  {
                                                       ClassStuNum = 0;
                                                       foreach (StudentRecord stuRec in classRec.Students)
                                                       {
                                                            ClassStuNum++;
                                                            if (ClassStuNum > conf.StudentLimit)
                                                                 break;
                                                            string studentID = stuRec.StudentID;
                                                            if (studentExamSores.ContainsKey(studentID))
                                                            {
                                                                 if (studentExamSores[studentID].ContainsKey(subjectName))
                                                                 {
                                                                      foreach (var courseID in studentExamSores[studentID][subjectName].Keys)
                                                                      {
                                                                           if (accessHelper.CourseHelper.GetCourse(courseID)[0].SubjectLevel == subjectLevel)//studentExamSores[studentID][subjectName][courseID]可能是null
                                                                           {
                                                                                #region 評量成績
                                                                                var sceTakeRecord = studentExamSores[studentID][subjectName][courseID];
                                                                                if (sceTakeRecord != null)
                                                                                {//有輸入
                                                                                     row["科目成績" + ClassStuNum + "-" + subjectIndex] = sceTakeRecord.SpecialCase == "" ? ("" + sceTakeRecord.ExamScore) : sceTakeRecord.SpecialCase;
                                                                                     #region 班排名及落點分析
                                                                                     if (stuRec.RefClass != null)
                                                                                     {
                                                                                          key = "班排名" + stuRec.RefClass.ClassID + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                                          if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                                                                          {
                                                                                               row["班排名" + ClassStuNum + "-" + subjectIndex] = ranks[key].IndexOf(sceTakeRecord.ExamScore) + 1;
                                                                                               row["班排名母數" + ClassStuNum + "-" + subjectIndex] = ranks[key].Count;
                                                                                          }
                                                                                     }
                                                                                     #endregion
                                                                                     #region 科排名及落點分析
                                                                                     if (stuRec.Department != "")
                                                                                     {
                                                                                          key = "科排名" + stuRec.Department + "^^^" + gradeYear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                                          if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                                                                          {
                                                                                               row["科排名" + ClassStuNum + "-" + subjectIndex] = ranks[key].IndexOf(sceTakeRecord.ExamScore) + 1;
                                                                                               row["科排名母數" + ClassStuNum + "-" + subjectIndex] = ranks[key].Count;
                                                                                          }
                                                                                     }
                                                                                     #endregion
                                                                                     #region 全校排名及落點分析
                                                                                     key = "全校排名" + gradeYear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                                     if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                                                                     {
                                                                                          row["全校排名" + ClassStuNum + "-" + subjectIndex] = ranks[key].IndexOf(sceTakeRecord.ExamScore) + 1;
                                                                                          row["全校排名母數" + ClassStuNum + "-" + subjectIndex] = ranks[key].Count;
                                                                                     }
                                                                                     #endregion
                                                                                     #region 類別1排名及落點分析
                                                                                     if (studentTag1Group.ContainsKey(studentID) && conf.TagRank1SubjectList.Contains(subjectName))
                                                                                     {
                                                                                          key = "類別1排名" + studentTag1Group[studentID] + "^^^" + gradeYear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                                          if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                                                                          {
                                                                                               row["類別1排名" + ClassStuNum + "-" + subjectIndex] = ranks[key].IndexOf(sceTakeRecord.ExamScore) + 1;
                                                                                               row["類別1排名母數" + ClassStuNum + "-" + subjectIndex] = ranks[key].Count;
                                                                                          }
                                                                                     }
                                                                                     #endregion
                                                                                     #region 類別2排名及落點分析
                                                                                     if (studentTag2Group.ContainsKey(studentID) && conf.TagRank2SubjectList.Contains(subjectName))
                                                                                     {
                                                                                          key = "類別2排名" + studentTag2Group[studentID] + "^^^" + gradeYear + "^^^" + sceTakeRecord.Subject + "^^^" + sceTakeRecord.SubjectLevel;
                                                                                          if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                                                                          {
                                                                                               row["類別2排名" + ClassStuNum + "-" + subjectIndex] = ranks[key].IndexOf(sceTakeRecord.ExamScore) + 1;
                                                                                               row["類別2排名母數" + ClassStuNum + "-" + subjectIndex] = ranks[key].Count;
                                                                                          }
                                                                                     }
                                                                                     #endregion
                                                                                }
                                                                                else
                                                                                {//修課有該考試但沒有成績資料
                                                                                     row["科目成績" + ClassStuNum + "-" + subjectIndex] = "未輸入";
                                                                                }
                                                                                #endregion
                                                                                #region 參考成績
                                                                                if (studentRefExamSores.ContainsKey(studentID) && studentRefExamSores[studentID].ContainsKey(courseID))
                                                                                {
                                                                                     row["前次成績" + ClassStuNum + "-" + subjectIndex] =
                                                                                         studentRefExamSores[studentID][courseID].SpecialCase == ""
                                                                                         ? ("" + studentRefExamSores[studentID][courseID].ExamScore)
                                                                                         : studentRefExamSores[studentID][courseID].SpecialCase;
                                                                                }
                                                                                #endregion
                                                                                break;
                                                                           }
                                                                      }
                                                                 }
                                                            }
                                                       }
                                                       #region 班組距及高低標
                                                       key = "班排名" + classRec.ClassID + "^^^" + subjectName + "^^^" + subjectLevel;
                                                       if (rankStudents.ContainsKey(key))
                                                       {
                                                            row["班高標" + subjectIndex] = analytics[key + "^^^高標"];
                                                            row["班均標" + subjectIndex] = analytics[key + "^^^均標"];
                                                            row["班低標" + subjectIndex] = analytics[key + "^^^低標"];
                                                            row["班標準差" + subjectIndex] = analytics[key + "^^^標準差"];
                                                            row["班組距" + subjectIndex + "count90"] = analytics[key + "^^^count90"];
                                                            row["班組距" + subjectIndex + "count80"] = analytics[key + "^^^count80"];
                                                            row["班組距" + subjectIndex + "count70"] = analytics[key + "^^^count70"];
                                                            row["班組距" + subjectIndex + "count60"] = analytics[key + "^^^count60"];
                                                            row["班組距" + subjectIndex + "count50"] = analytics[key + "^^^count50"];
                                                            row["班組距" + subjectIndex + "count40"] = analytics[key + "^^^count40"];
                                                            row["班組距" + subjectIndex + "count30"] = analytics[key + "^^^count30"];
                                                            row["班組距" + subjectIndex + "count20"] = analytics[key + "^^^count20"];
                                                            row["班組距" + subjectIndex + "count10"] = analytics[key + "^^^count10"];
                                                            row["班組距" + subjectIndex + "count100Up"] = analytics[key + "^^^count100Up"];
                                                            row["班組距" + subjectIndex + "count90Up"] = analytics[key + "^^^count90Up"];
                                                            row["班組距" + subjectIndex + "count80Up"] = analytics[key + "^^^count80Up"];
                                                            row["班組距" + subjectIndex + "count70Up"] = analytics[key + "^^^count70Up"];
                                                            row["班組距" + subjectIndex + "count60Up"] = analytics[key + "^^^count60Up"];
                                                            row["班組距" + subjectIndex + "count50Up"] = analytics[key + "^^^count50Up"];
                                                            row["班組距" + subjectIndex + "count40Up"] = analytics[key + "^^^count40Up"];
                                                            row["班組距" + subjectIndex + "count30Up"] = analytics[key + "^^^count30Up"];
                                                            row["班組距" + subjectIndex + "count20Up"] = analytics[key + "^^^count20Up"];
                                                            row["班組距" + subjectIndex + "count10Up"] = analytics[key + "^^^count10Up"];
                                                            row["班組距" + subjectIndex + "count90Down"] = analytics[key + "^^^count90Down"];
                                                            row["班組距" + subjectIndex + "count80Down"] = analytics[key + "^^^count80Down"];
                                                            row["班組距" + subjectIndex + "count70Down"] = analytics[key + "^^^count70Down"];
                                                            row["班組距" + subjectIndex + "count60Down"] = analytics[key + "^^^count60Down"];
                                                            row["班組距" + subjectIndex + "count50Down"] = analytics[key + "^^^count50Down"];
                                                            row["班組距" + subjectIndex + "count40Down"] = analytics[key + "^^^count40Down"];
                                                            row["班組距" + subjectIndex + "count30Down"] = analytics[key + "^^^count30Down"];
                                                            row["班組距" + subjectIndex + "count20Down"] = analytics[key + "^^^count20Down"];
                                                            row["班組距" + subjectIndex + "count10Down"] = analytics[key + "^^^count10Down"];
                                                       }
                                                       #endregion
                                                       #region 科組距及高低標
                                                       key = "科排名" + classRec.Department + "^^^" + gradeYear + "^^^" + subjectName + "^^^" + subjectLevel;
                                                       if (rankStudents.ContainsKey(key))
                                                       {
                                                            row["科高標" + subjectIndex] = analytics[key + "^^^高標"];
                                                            row["科均標" + subjectIndex] = analytics[key + "^^^均標"];
                                                            row["科低標" + subjectIndex] = analytics[key + "^^^低標"];
                                                            row["科標準差" + subjectIndex] = analytics[key + "^^^標準差"];
                                                            row["科組距" + subjectIndex + "count90"] = analytics[key + "^^^count90"];
                                                            row["科組距" + subjectIndex + "count80"] = analytics[key + "^^^count80"];
                                                            row["科組距" + subjectIndex + "count70"] = analytics[key + "^^^count70"];
                                                            row["科組距" + subjectIndex + "count60"] = analytics[key + "^^^count60"];
                                                            row["科組距" + subjectIndex + "count50"] = analytics[key + "^^^count50"];
                                                            row["科組距" + subjectIndex + "count40"] = analytics[key + "^^^count40"];
                                                            row["科組距" + subjectIndex + "count30"] = analytics[key + "^^^count30"];
                                                            row["科組距" + subjectIndex + "count20"] = analytics[key + "^^^count20"];
                                                            row["科組距" + subjectIndex + "count10"] = analytics[key + "^^^count10"];
                                                            row["科組距" + subjectIndex + "count100Up"] = analytics[key + "^^^count100Up"];
                                                            row["科組距" + subjectIndex + "count90Up"] = analytics[key + "^^^count90Up"];
                                                            row["科組距" + subjectIndex + "count80Up"] = analytics[key + "^^^count80Up"];
                                                            row["科組距" + subjectIndex + "count70Up"] = analytics[key + "^^^count70Up"];
                                                            row["科組距" + subjectIndex + "count60Up"] = analytics[key + "^^^count60Up"];
                                                            row["科組距" + subjectIndex + "count50Up"] = analytics[key + "^^^count50Up"];
                                                            row["科組距" + subjectIndex + "count40Up"] = analytics[key + "^^^count40Up"];
                                                            row["科組距" + subjectIndex + "count30Up"] = analytics[key + "^^^count30Up"];
                                                            row["科組距" + subjectIndex + "count20Up"] = analytics[key + "^^^count20Up"];
                                                            row["科組距" + subjectIndex + "count10Up"] = analytics[key + "^^^count10Up"];
                                                            row["科組距" + subjectIndex + "count90Down"] = analytics[key + "^^^count90Down"];
                                                            row["科組距" + subjectIndex + "count80Down"] = analytics[key + "^^^count80Down"];
                                                            row["科組距" + subjectIndex + "count70Down"] = analytics[key + "^^^count70Down"];
                                                            row["科組距" + subjectIndex + "count60Down"] = analytics[key + "^^^count60Down"];
                                                            row["科組距" + subjectIndex + "count50Down"] = analytics[key + "^^^count50Down"];
                                                            row["科組距" + subjectIndex + "count40Down"] = analytics[key + "^^^count40Down"];
                                                            row["科組距" + subjectIndex + "count30Down"] = analytics[key + "^^^count30Down"];
                                                            row["科組距" + subjectIndex + "count20Down"] = analytics[key + "^^^count20Down"];
                                                            row["科組距" + subjectIndex + "count10Down"] = analytics[key + "^^^count10Down"];
                                                       }
                                                       #endregion
                                                       #region 校組距及高低標
                                                       key = "全校排名" + gradeYear + "^^^" + subjectName + "^^^" + subjectLevel;
                                                       if (rankStudents.ContainsKey(key))
                                                       {
                                                            row["校高標" + subjectIndex] = analytics[key + "^^^高標"];
                                                            row["校均標" + subjectIndex] = analytics[key + "^^^均標"];
                                                            row["校低標" + subjectIndex] = analytics[key + "^^^低標"];
                                                            row["校標準差" + subjectIndex] = analytics[key + "^^^標準差"];
                                                            row["校組距" + subjectIndex + "count90"] = analytics[key + "^^^count90"];
                                                            row["校組距" + subjectIndex + "count80"] = analytics[key + "^^^count80"];
                                                            row["校組距" + subjectIndex + "count70"] = analytics[key + "^^^count70"];
                                                            row["校組距" + subjectIndex + "count60"] = analytics[key + "^^^count60"];
                                                            row["校組距" + subjectIndex + "count50"] = analytics[key + "^^^count50"];
                                                            row["校組距" + subjectIndex + "count40"] = analytics[key + "^^^count40"];
                                                            row["校組距" + subjectIndex + "count30"] = analytics[key + "^^^count30"];
                                                            row["校組距" + subjectIndex + "count20"] = analytics[key + "^^^count20"];
                                                            row["校組距" + subjectIndex + "count10"] = analytics[key + "^^^count10"];
                                                            row["校組距" + subjectIndex + "count100Up"] = analytics[key + "^^^count100Up"];
                                                            row["校組距" + subjectIndex + "count90Up"] = analytics[key + "^^^count90Up"];
                                                            row["校組距" + subjectIndex + "count80Up"] = analytics[key + "^^^count80Up"];
                                                            row["校組距" + subjectIndex + "count70Up"] = analytics[key + "^^^count70Up"];
                                                            row["校組距" + subjectIndex + "count60Up"] = analytics[key + "^^^count60Up"];
                                                            row["校組距" + subjectIndex + "count50Up"] = analytics[key + "^^^count50Up"];
                                                            row["校組距" + subjectIndex + "count40Up"] = analytics[key + "^^^count40Up"];
                                                            row["校組距" + subjectIndex + "count30Up"] = analytics[key + "^^^count30Up"];
                                                            row["校組距" + subjectIndex + "count20Up"] = analytics[key + "^^^count20Up"];
                                                            row["校組距" + subjectIndex + "count10Up"] = analytics[key + "^^^count10Up"];
                                                            row["校組距" + subjectIndex + "count90Down"] = analytics[key + "^^^count90Down"];
                                                            row["校組距" + subjectIndex + "count80Down"] = analytics[key + "^^^count80Down"];
                                                            row["校組距" + subjectIndex + "count70Down"] = analytics[key + "^^^count70Down"];
                                                            row["校組距" + subjectIndex + "count60Down"] = analytics[key + "^^^count60Down"];
                                                            row["校組距" + subjectIndex + "count50Down"] = analytics[key + "^^^count50Down"];
                                                            row["校組距" + subjectIndex + "count40Down"] = analytics[key + "^^^count40Down"];
                                                            row["校組距" + subjectIndex + "count30Down"] = analytics[key + "^^^count30Down"];
                                                            row["校組距" + subjectIndex + "count20Down"] = analytics[key + "^^^count20Down"];
                                                            row["校組距" + subjectIndex + "count10Down"] = analytics[key + "^^^count10Down"];
                                                       }
                                                       #endregion
                                                       #region 類別1組距及高低標
                                                       for (int i = 0; i < tag1List.Count; i++)
                                                       {
                                                            string tag = tag1List[i];
                                                            key = "類別1排名" + tag + "^^^" + gradeYear + "^^^" + subjectName + "^^^" + subjectLevel;
                                                            if (rankStudents.ContainsKey(key))
                                                            {
                                                                 if (i == 0)
                                                                 {
                                                                      row["類1高標" + subjectIndex] = analytics[key + "^^^高標"];
                                                                      row["類1均標" + subjectIndex] = analytics[key + "^^^均標"];
                                                                      row["類1低標" + subjectIndex] = analytics[key + "^^^低標"];
                                                                      row["類1標準差" + subjectIndex] = analytics[key + "^^^標準差"];
                                                                      row["類1組距" + subjectIndex + "count90"] = analytics[key + "^^^count90"];
                                                                      row["類1組距" + subjectIndex + "count80"] = analytics[key + "^^^count80"];
                                                                      row["類1組距" + subjectIndex + "count70"] = analytics[key + "^^^count70"];
                                                                      row["類1組距" + subjectIndex + "count60"] = analytics[key + "^^^count60"];
                                                                      row["類1組距" + subjectIndex + "count50"] = analytics[key + "^^^count50"];
                                                                      row["類1組距" + subjectIndex + "count40"] = analytics[key + "^^^count40"];
                                                                      row["類1組距" + subjectIndex + "count30"] = analytics[key + "^^^count30"];
                                                                      row["類1組距" + subjectIndex + "count20"] = analytics[key + "^^^count20"];
                                                                      row["類1組距" + subjectIndex + "count10"] = analytics[key + "^^^count10"];
                                                                      row["類1組距" + subjectIndex + "count100Up"] = analytics[key + "^^^count100Up"];
                                                                      row["類1組距" + subjectIndex + "count90Up"] = analytics[key + "^^^count90Up"];
                                                                      row["類1組距" + subjectIndex + "count80Up"] = analytics[key + "^^^count80Up"];
                                                                      row["類1組距" + subjectIndex + "count70Up"] = analytics[key + "^^^count70Up"];
                                                                      row["類1組距" + subjectIndex + "count60Up"] = analytics[key + "^^^count60Up"];
                                                                      row["類1組距" + subjectIndex + "count50Up"] = analytics[key + "^^^count50Up"];
                                                                      row["類1組距" + subjectIndex + "count40Up"] = analytics[key + "^^^count40Up"];
                                                                      row["類1組距" + subjectIndex + "count30Up"] = analytics[key + "^^^count30Up"];
                                                                      row["類1組距" + subjectIndex + "count20Up"] = analytics[key + "^^^count20Up"];
                                                                      row["類1組距" + subjectIndex + "count10Up"] = analytics[key + "^^^count10Up"];
                                                                      row["類1組距" + subjectIndex + "count90Down"] = analytics[key + "^^^count90Down"];
                                                                      row["類1組距" + subjectIndex + "count80Down"] = analytics[key + "^^^count80Down"];
                                                                      row["類1組距" + subjectIndex + "count70Down"] = analytics[key + "^^^count70Down"];
                                                                      row["類1組距" + subjectIndex + "count60Down"] = analytics[key + "^^^count60Down"];
                                                                      row["類1組距" + subjectIndex + "count50Down"] = analytics[key + "^^^count50Down"];
                                                                      row["類1組距" + subjectIndex + "count40Down"] = analytics[key + "^^^count40Down"];
                                                                      row["類1組距" + subjectIndex + "count30Down"] = analytics[key + "^^^count30Down"];
                                                                      row["類1組距" + subjectIndex + "count20Down"] = analytics[key + "^^^count20Down"];
                                                                      row["類1組距" + subjectIndex + "count10Down"] = analytics[key + "^^^count10Down"];
                                                                 }
                                                                 else
                                                                 {
                                                                      row["類1高標" + subjectIndex] += "/" + analytics[key + "^^^高標"];
                                                                      row["類1均標" + subjectIndex] += "/" + analytics[key + "^^^均標"];
                                                                      row["類1低標" + subjectIndex] += "/" + analytics[key + "^^^低標"];
                                                                      row["類1標準差" + subjectIndex] += "/" + analytics[key + "^^^標準差"];
                                                                      row["類1組距" + subjectIndex + "count90"] += "/" + analytics[key + "^^^count90"];
                                                                      row["類1組距" + subjectIndex + "count80"] += "/" + analytics[key + "^^^count80"];
                                                                      row["類1組距" + subjectIndex + "count70"] += "/" + analytics[key + "^^^count70"];
                                                                      row["類1組距" + subjectIndex + "count60"] += "/" + analytics[key + "^^^count60"];
                                                                      row["類1組距" + subjectIndex + "count50"] += "/" + analytics[key + "^^^count50"];
                                                                      row["類1組距" + subjectIndex + "count40"] += "/" + analytics[key + "^^^count40"];
                                                                      row["類1組距" + subjectIndex + "count30"] += "/" + analytics[key + "^^^count30"];
                                                                      row["類1組距" + subjectIndex + "count20"] += "/" + analytics[key + "^^^count20"];
                                                                      row["類1組距" + subjectIndex + "count10"] += "/" + analytics[key + "^^^count10"];
                                                                      row["類1組距" + subjectIndex + "count100Up"] += "/" + analytics[key + "^^^count100Up"];
                                                                      row["類1組距" + subjectIndex + "count90Up"] += "/" + analytics[key + "^^^count90Up"];
                                                                      row["類1組距" + subjectIndex + "count80Up"] += "/" + analytics[key + "^^^count80Up"];
                                                                      row["類1組距" + subjectIndex + "count70Up"] += "/" + analytics[key + "^^^count70Up"];
                                                                      row["類1組距" + subjectIndex + "count60Up"] += "/" + analytics[key + "^^^count60Up"];
                                                                      row["類1組距" + subjectIndex + "count50Up"] += "/" + analytics[key + "^^^count50Up"];
                                                                      row["類1組距" + subjectIndex + "count40Up"] += "/" + analytics[key + "^^^count40Up"];
                                                                      row["類1組距" + subjectIndex + "count30Up"] += "/" + analytics[key + "^^^count30Up"];
                                                                      row["類1組距" + subjectIndex + "count20Up"] += "/" + analytics[key + "^^^count20Up"];
                                                                      row["類1組距" + subjectIndex + "count10Up"] += "/" + analytics[key + "^^^count10Up"];
                                                                      row["類1組距" + subjectIndex + "count90Down"] += "/" + analytics[key + "^^^count90Down"];
                                                                      row["類1組距" + subjectIndex + "count80Down"] += "/" + analytics[key + "^^^count80Down"];
                                                                      row["類1組距" + subjectIndex + "count70Down"] += "/" + analytics[key + "^^^count70Down"];
                                                                      row["類1組距" + subjectIndex + "count60Down"] += "/" + analytics[key + "^^^count60Down"];
                                                                      row["類1組距" + subjectIndex + "count50Down"] += "/" + analytics[key + "^^^count50Down"];
                                                                      row["類1組距" + subjectIndex + "count40Down"] += "/" + analytics[key + "^^^count40Down"];
                                                                      row["類1組距" + subjectIndex + "count30Down"] += "/" + analytics[key + "^^^count30Down"];
                                                                      row["類1組距" + subjectIndex + "count20Down"] += "/" + analytics[key + "^^^count20Down"];
                                                                      row["類1組距" + subjectIndex + "count10Down"] += "/" + analytics[key + "^^^count10Down"];
                                                                 }
                                                            }
                                                       }
                                                       #endregion
                                                       #region 類別2組距及高低標
                                                       for (int i = 0; i < tag2List.Count; i++)
                                                       {
                                                            string tag = tag2List[i];
                                                            key = "類別2排名" + tag + "^^^" + gradeYear + "^^^" + subjectName + "^^^" + subjectLevel;
                                                            if (rankStudents.ContainsKey(key))
                                                            {
                                                                 if (i == 0)
                                                                 {
                                                                      row["類2高標" + subjectIndex] = analytics[key + "^^^高標"];
                                                                      row["類2均標" + subjectIndex] = analytics[key + "^^^均標"];
                                                                      row["類2低標" + subjectIndex] = analytics[key + "^^^低標"];
                                                                      row["類2標準差" + subjectIndex] = analytics[key + "^^^標準差"];
                                                                      row["類2組距" + subjectIndex + "count90"] = analytics[key + "^^^count90"];
                                                                      row["類2組距" + subjectIndex + "count80"] = analytics[key + "^^^count80"];
                                                                      row["類2組距" + subjectIndex + "count70"] = analytics[key + "^^^count70"];
                                                                      row["類2組距" + subjectIndex + "count60"] = analytics[key + "^^^count60"];
                                                                      row["類2組距" + subjectIndex + "count50"] = analytics[key + "^^^count50"];
                                                                      row["類2組距" + subjectIndex + "count40"] = analytics[key + "^^^count40"];
                                                                      row["類2組距" + subjectIndex + "count30"] = analytics[key + "^^^count30"];
                                                                      row["類2組距" + subjectIndex + "count20"] = analytics[key + "^^^count20"];
                                                                      row["類2組距" + subjectIndex + "count10"] = analytics[key + "^^^count10"];
                                                                      row["類2組距" + subjectIndex + "count100Up"] = analytics[key + "^^^count100Up"];
                                                                      row["類2組距" + subjectIndex + "count90Up"] = analytics[key + "^^^count90Up"];
                                                                      row["類2組距" + subjectIndex + "count80Up"] = analytics[key + "^^^count80Up"];
                                                                      row["類2組距" + subjectIndex + "count70Up"] = analytics[key + "^^^count70Up"];
                                                                      row["類2組距" + subjectIndex + "count60Up"] = analytics[key + "^^^count60Up"];
                                                                      row["類2組距" + subjectIndex + "count50Up"] = analytics[key + "^^^count50Up"];
                                                                      row["類2組距" + subjectIndex + "count40Up"] = analytics[key + "^^^count40Up"];
                                                                      row["類2組距" + subjectIndex + "count30Up"] = analytics[key + "^^^count30Up"];
                                                                      row["類2組距" + subjectIndex + "count20Up"] = analytics[key + "^^^count20Up"];
                                                                      row["類2組距" + subjectIndex + "count10Up"] = analytics[key + "^^^count10Up"];
                                                                      row["類2組距" + subjectIndex + "count90Down"] = analytics[key + "^^^count90Down"];
                                                                      row["類2組距" + subjectIndex + "count80Down"] = analytics[key + "^^^count80Down"];
                                                                      row["類2組距" + subjectIndex + "count70Down"] = analytics[key + "^^^count70Down"];
                                                                      row["類2組距" + subjectIndex + "count60Down"] = analytics[key + "^^^count60Down"];
                                                                      row["類2組距" + subjectIndex + "count50Down"] = analytics[key + "^^^count50Down"];
                                                                      row["類2組距" + subjectIndex + "count40Down"] = analytics[key + "^^^count40Down"];
                                                                      row["類2組距" + subjectIndex + "count30Down"] = analytics[key + "^^^count30Down"];
                                                                      row["類2組距" + subjectIndex + "count20Down"] = analytics[key + "^^^count20Down"];
                                                                      row["類2組距" + subjectIndex + "count10Down"] = analytics[key + "^^^count10Down"];
                                                                 }
                                                                 else
                                                                 {
                                                                      row["類2高標" + subjectIndex] += "/" + analytics[key + "^^^高標"];
                                                                      row["類2均標" + subjectIndex] += "/" + analytics[key + "^^^均標"];
                                                                      row["類2低標" + subjectIndex] += "/" + analytics[key + "^^^低標"];
                                                                      row["類2標準差" + subjectIndex] += "/" + analytics[key + "^^^標準差"];
                                                                      row["類2組距" + subjectIndex + "count90"] += "/" + analytics[key + "^^^count90"];
                                                                      row["類2組距" + subjectIndex + "count80"] += "/" + analytics[key + "^^^count80"];
                                                                      row["類2組距" + subjectIndex + "count70"] += "/" + analytics[key + "^^^count70"];
                                                                      row["類2組距" + subjectIndex + "count60"] += "/" + analytics[key + "^^^count60"];
                                                                      row["類2組距" + subjectIndex + "count50"] += "/" + analytics[key + "^^^count50"];
                                                                      row["類2組距" + subjectIndex + "count40"] += "/" + analytics[key + "^^^count40"];
                                                                      row["類2組距" + subjectIndex + "count30"] += "/" + analytics[key + "^^^count30"];
                                                                      row["類2組距" + subjectIndex + "count20"] += "/" + analytics[key + "^^^count20"];
                                                                      row["類2組距" + subjectIndex + "count10"] += "/" + analytics[key + "^^^count10"];
                                                                      row["類2組距" + subjectIndex + "count100Up"] += "/" + analytics[key + "^^^count100Up"];
                                                                      row["類2組距" + subjectIndex + "count90Up"] += "/" + analytics[key + "^^^count90Up"];
                                                                      row["類2組距" + subjectIndex + "count80Up"] += "/" + analytics[key + "^^^count80Up"];
                                                                      row["類2組距" + subjectIndex + "count70Up"] += "/" + analytics[key + "^^^count70Up"];
                                                                      row["類2組距" + subjectIndex + "count60Up"] += "/" + analytics[key + "^^^count60Up"];
                                                                      row["類2組距" + subjectIndex + "count50Up"] += "/" + analytics[key + "^^^count50Up"];
                                                                      row["類2組距" + subjectIndex + "count40Up"] += "/" + analytics[key + "^^^count40Up"];
                                                                      row["類2組距" + subjectIndex + "count30Up"] += "/" + analytics[key + "^^^count30Up"];
                                                                      row["類2組距" + subjectIndex + "count20Up"] += "/" + analytics[key + "^^^count20Up"];
                                                                      row["類2組距" + subjectIndex + "count10Up"] += "/" + analytics[key + "^^^count10Up"];
                                                                      row["類2組距" + subjectIndex + "count90Down"] += "/" + analytics[key + "^^^count90Down"];
                                                                      row["類2組距" + subjectIndex + "count80Down"] += "/" + analytics[key + "^^^count80Down"];
                                                                      row["類2組距" + subjectIndex + "count70Down"] += "/" + analytics[key + "^^^count70Down"];
                                                                      row["類2組距" + subjectIndex + "count60Down"] += "/" + analytics[key + "^^^count60Down"];
                                                                      row["類2組距" + subjectIndex + "count50Down"] += "/" + analytics[key + "^^^count50Down"];
                                                                      row["類2組距" + subjectIndex + "count40Down"] += "/" + analytics[key + "^^^count40Down"];
                                                                      row["類2組距" + subjectIndex + "count30Down"] += "/" + analytics[key + "^^^count30Down"];
                                                                      row["類2組距" + subjectIndex + "count20Down"] += "/" + analytics[key + "^^^count20Down"];
                                                                      row["類2組距" + subjectIndex + "count10Down"] += "/" + analytics[key + "^^^count10Down"];
                                                                 }
                                                            }
                                                       }
                                                       #endregion
                                                  }
                                                  subjectIndex++;
                                             }
                                             else
                                             {
                                                  //重要!!發現資料在樣板中印不下時一定要記錄起來，否則使用者自己不會去發現的
                                                  if (!overflowRecords.Contains(classRec))
                                                       overflowRecords.Add(classRec);
                                             }
                                        }
                                   }
                                   #endregion
                                   #region 總分
                                   ClassStuNum = 0;
                                   foreach (StudentRecord stuRec in classRec.Students)
                                   {
                                        ClassStuNum++;
                                        if (ClassStuNum > conf.StudentLimit)
                                             break;
                                        string studentID = stuRec.StudentID;
                                        if (studentPrintSubjectSum.ContainsKey(studentID))
                                        {
                                             row["總分" + ClassStuNum] = studentPrintSubjectSum[studentID];
                                             //總分班排名                                
                                             key = "總分班排名" + classRec.ClassID;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["總分班排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectSum[studentID]) + 1;
                                                  row["總分班排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                             //總分科排名
                                             key = "總分科排名" + stuRec.Department + "^^^" + gradeYear;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["總分科排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectSum[studentID]) + 1;
                                                  row["總分科排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                             //總分全校排名
                                             key = "總分全校排名" + gradeYear;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["總分全校排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectSum[studentID]) + 1;
                                                  row["總分全校排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                        }
                                   }
                                   #region 總分班組距及高低標
                                   key = "總分班排名" + classRec.ClassID;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["總分班高標"] = analytics[key + "^^^高標"];
                                        row["總分班均標"] = analytics[key + "^^^均標"];
                                        row["總分班低標"] = analytics[key + "^^^低標"];
                                        row["總分班標準差"] = analytics[key + "^^^標準差"];
                                        row["總分班組距count90"] = analytics[key + "^^^count90"];
                                        row["總分班組距count80"] = analytics[key + "^^^count80"];
                                        row["總分班組距count70"] = analytics[key + "^^^count70"];
                                        row["總分班組距count60"] = analytics[key + "^^^count60"];
                                        row["總分班組距count50"] = analytics[key + "^^^count50"];
                                        row["總分班組距count40"] = analytics[key + "^^^count40"];
                                        row["總分班組距count30"] = analytics[key + "^^^count30"];
                                        row["總分班組距count20"] = analytics[key + "^^^count20"];
                                        row["總分班組距count10"] = analytics[key + "^^^count10"];
                                        row["總分班組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["總分班組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["總分班組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["總分班組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["總分班組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["總分班組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["總分班組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["總分班組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["總分班組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["總分班組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["總分班組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["總分班組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["總分班組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["總分班組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["總分班組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["總分班組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["總分班組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["總分班組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["總分班組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #region 總分科組距及高低標
                                   key = "總分科排名" + classRec.Department + "^^^" + gradeYear;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["總分科高標"] = analytics[key + "^^^高標"];
                                        row["總分科均標"] = analytics[key + "^^^均標"];
                                        row["總分科低標"] = analytics[key + "^^^低標"];
                                        row["總分科標準差"] = analytics[key + "^^^標準差"];
                                        row["總分科組距count90"] = analytics[key + "^^^count90"];
                                        row["總分科組距count80"] = analytics[key + "^^^count80"];
                                        row["總分科組距count70"] = analytics[key + "^^^count70"];
                                        row["總分科組距count60"] = analytics[key + "^^^count60"];
                                        row["總分科組距count50"] = analytics[key + "^^^count50"];
                                        row["總分科組距count40"] = analytics[key + "^^^count40"];
                                        row["總分科組距count30"] = analytics[key + "^^^count30"];
                                        row["總分科組距count20"] = analytics[key + "^^^count20"];
                                        row["總分科組距count10"] = analytics[key + "^^^count10"];
                                        row["總分科組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["總分科組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["總分科組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["總分科組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["總分科組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["總分科組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["總分科組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["總分科組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["總分科組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["總分科組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["總分科組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["總分科組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["總分科組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["總分科組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["總分科組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["總分科組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["總分科組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["總分科組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["總分科組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #region 總分校組距及高低標
                                   key = "總分全校排名" + gradeYear;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["總分校高標"] = analytics[key + "^^^高標"];
                                        row["總分校均標"] = analytics[key + "^^^均標"];
                                        row["總分校低標"] = analytics[key + "^^^低標"];
                                        row["總分校標準差"] = analytics[key + "^^^標準差"];
                                        row["總分校組距count90"] = analytics[key + "^^^count90"];
                                        row["總分校組距count80"] = analytics[key + "^^^count80"];
                                        row["總分校組距count70"] = analytics[key + "^^^count70"];
                                        row["總分校組距count60"] = analytics[key + "^^^count60"];
                                        row["總分校組距count50"] = analytics[key + "^^^count50"];
                                        row["總分校組距count40"] = analytics[key + "^^^count40"];
                                        row["總分校組距count30"] = analytics[key + "^^^count30"];
                                        row["總分校組距count20"] = analytics[key + "^^^count20"];
                                        row["總分校組距count10"] = analytics[key + "^^^count10"];
                                        row["總分校組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["總分校組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["總分校組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["總分校組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["總分校組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["總分校組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["總分校組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["總分校組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["總分校組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["總分校組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["總分校組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["總分校組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["總分校組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["總分校組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["總分校組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["總分校組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["總分校組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["總分校組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["總分校組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #endregion
                                   #region 平均
                                   ClassStuNum = 0;
                                   foreach (StudentRecord stuRec in classRec.Students)
                                   {
                                        ClassStuNum++;
                                        if (ClassStuNum > conf.StudentLimit)
                                             break;
                                        string studentID = stuRec.StudentID;
                                        if (studentPrintSubjectAvg.ContainsKey(studentID))
                                        {
                                             row["平均" + ClassStuNum] = studentPrintSubjectAvg[studentID];
                                             key = "平均班排名" + classRec.ClassID;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["平均班排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) + 1;
                                                  row["平均班排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                             key = "平均科排名" + stuRec.Department + "^^^" + gradeYear;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["平均科排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) + 1;
                                                  row["平均科排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                             key = "平均全校排名" + gradeYear;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["平均全校排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectAvg[studentID]) + 1;
                                                  row["平均全校排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                        }
                                   }
                                   #region 平均班排名及高低標
                                   key = "平均班排名" + classRec.ClassID;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["平均班高標"] = analytics[key + "^^^高標"];
                                        row["平均班均標"] = analytics[key + "^^^均標"];
                                        row["平均班低標"] = analytics[key + "^^^低標"];
                                        row["平均班標準差"] = analytics[key + "^^^標準差"];
                                        row["平均班組距count90"] = analytics[key + "^^^count90"];
                                        row["平均班組距count80"] = analytics[key + "^^^count80"];
                                        row["平均班組距count70"] = analytics[key + "^^^count70"];
                                        row["平均班組距count60"] = analytics[key + "^^^count60"];
                                        row["平均班組距count50"] = analytics[key + "^^^count50"];
                                        row["平均班組距count40"] = analytics[key + "^^^count40"];
                                        row["平均班組距count30"] = analytics[key + "^^^count30"];
                                        row["平均班組距count20"] = analytics[key + "^^^count20"];
                                        row["平均班組距count10"] = analytics[key + "^^^count10"];
                                        row["平均班組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["平均班組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["平均班組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["平均班組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["平均班組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["平均班組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["平均班組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["平均班組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["平均班組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["平均班組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["平均班組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["平均班組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["平均班組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["平均班組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["平均班組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["平均班組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["平均班組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["平均班組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["平均班組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #region 平均科排名及高低標
                                   key = "平均科排名" + classRec.Department + "^^^" + gradeYear;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["平均科高標"] = analytics[key + "^^^高標"];
                                        row["平均科均標"] = analytics[key + "^^^均標"];
                                        row["平均科低標"] = analytics[key + "^^^低標"];
                                        row["平均科標準差"] = analytics[key + "^^^標準差"];
                                        row["平均科組距count90"] = analytics[key + "^^^count90"];
                                        row["平均科組距count80"] = analytics[key + "^^^count80"];
                                        row["平均科組距count70"] = analytics[key + "^^^count70"];
                                        row["平均科組距count60"] = analytics[key + "^^^count60"];
                                        row["平均科組距count50"] = analytics[key + "^^^count50"];
                                        row["平均科組距count40"] = analytics[key + "^^^count40"];
                                        row["平均科組距count30"] = analytics[key + "^^^count30"];
                                        row["平均科組距count20"] = analytics[key + "^^^count20"];
                                        row["平均科組距count10"] = analytics[key + "^^^count10"];
                                        row["平均科組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["平均科組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["平均科組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["平均科組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["平均科組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["平均科組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["平均科組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["平均科組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["平均科組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["平均科組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["平均科組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["平均科組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["平均科組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["平均科組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["平均科組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["平均科組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["平均科組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["平均科組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["平均科組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #region 平均校排名及高低標
                                   key = "平均全校排名" + gradeYear;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["平均校高標"] = analytics[key + "^^^高標"];
                                        row["平均校均標"] = analytics[key + "^^^均標"];
                                        row["平均校低標"] = analytics[key + "^^^低標"];
                                        row["平均校標準差"] = analytics[key + "^^^標準差"];
                                        row["平均校組距count90"] = analytics[key + "^^^count90"];
                                        row["平均校組距count80"] = analytics[key + "^^^count80"];
                                        row["平均校組距count70"] = analytics[key + "^^^count70"];
                                        row["平均校組距count60"] = analytics[key + "^^^count60"];
                                        row["平均校組距count50"] = analytics[key + "^^^count50"];
                                        row["平均校組距count40"] = analytics[key + "^^^count40"];
                                        row["平均校組距count30"] = analytics[key + "^^^count30"];
                                        row["平均校組距count20"] = analytics[key + "^^^count20"];
                                        row["平均校組距count10"] = analytics[key + "^^^count10"];
                                        row["平均校組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["平均校組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["平均校組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["平均校組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["平均校組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["平均校組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["平均校組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["平均校組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["平均校組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["平均校組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["平均校組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["平均校組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["平均校組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["平均校組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["平均校組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["平均校組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["平均校組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["平均校組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["平均校組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #endregion
                                   #region 加權總分
                                   ClassStuNum = 0;
                                   foreach (StudentRecord stuRec in classRec.Students)
                                   {
                                        ClassStuNum++;
                                        if (ClassStuNum > conf.StudentLimit)
                                             break;
                                        string studentID = stuRec.StudentID;
                                        if (studentPrintSubjectSumW.ContainsKey(studentID))
                                        {
                                             row["加權總分" + ClassStuNum] = studentPrintSubjectSumW[studentID];
                                             key = "加權總分班排名" + classRec.ClassID;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["加權總分班排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) + 1;
                                                  row["加權總分班排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                             key = "加權總分科排名" + stuRec.Department + "^^^" + gradeYear;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["加權總分科排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) + 1;
                                                  row["加權總分科排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                             key = "加權總分全校排名" + gradeYear;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["加權總分全校排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectSumW[studentID]) + 1;
                                                  row["加權總分全校排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                        }
                                   }
                                   #region 加權總分班組距及高低標
                                   key = "加權總分班排名" + classRec.ClassID;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["加權總分班高標"] = analytics[key + "^^^高標"];
                                        row["加權總分班均標"] = analytics[key + "^^^均標"];
                                        row["加權總分班低標"] = analytics[key + "^^^低標"];
                                        row["加權總分班標準差"] = analytics[key + "^^^標準差"];
                                        row["加權總分班組距count90"] = analytics[key + "^^^count90"];
                                        row["加權總分班組距count80"] = analytics[key + "^^^count80"];
                                        row["加權總分班組距count70"] = analytics[key + "^^^count70"];
                                        row["加權總分班組距count60"] = analytics[key + "^^^count60"];
                                        row["加權總分班組距count50"] = analytics[key + "^^^count50"];
                                        row["加權總分班組距count40"] = analytics[key + "^^^count40"];
                                        row["加權總分班組距count30"] = analytics[key + "^^^count30"];
                                        row["加權總分班組距count20"] = analytics[key + "^^^count20"];
                                        row["加權總分班組距count10"] = analytics[key + "^^^count10"];
                                        row["加權總分班組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["加權總分班組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["加權總分班組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["加權總分班組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["加權總分班組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["加權總分班組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["加權總分班組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["加權總分班組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["加權總分班組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["加權總分班組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["加權總分班組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["加權總分班組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["加權總分班組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["加權總分班組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["加權總分班組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["加權總分班組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["加權總分班組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["加權總分班組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["加權總分班組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #region 加權總分科組距及高低標
                                   key = "加權總分科排名" + classRec.Department + "^^^" + gradeYear;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["加權總分科高標"] = analytics[key + "^^^高標"];
                                        row["加權總分科均標"] = analytics[key + "^^^均標"];
                                        row["加權總分科低標"] = analytics[key + "^^^低標"];
                                        row["加權總分科標準差"] = analytics[key + "^^^標準差"];
                                        row["加權總分科組距count90"] = analytics[key + "^^^count90"];
                                        row["加權總分科組距count80"] = analytics[key + "^^^count80"];
                                        row["加權總分科組距count70"] = analytics[key + "^^^count70"];
                                        row["加權總分科組距count60"] = analytics[key + "^^^count60"];
                                        row["加權總分科組距count50"] = analytics[key + "^^^count50"];
                                        row["加權總分科組距count40"] = analytics[key + "^^^count40"];
                                        row["加權總分科組距count30"] = analytics[key + "^^^count30"];
                                        row["加權總分科組距count20"] = analytics[key + "^^^count20"];
                                        row["加權總分科組距count10"] = analytics[key + "^^^count10"];
                                        row["加權總分科組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["加權總分科組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["加權總分科組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["加權總分科組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["加權總分科組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["加權總分科組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["加權總分科組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["加權總分科組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["加權總分科組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["加權總分科組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["加權總分科組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["加權總分科組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["加權總分科組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["加權總分科組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["加權總分科組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["加權總分科組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["加權總分科組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["加權總分科組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["加權總分科組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #region 加權總分全校組距及高低標
                                   key = "加權總分全校排名" + gradeYear;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["加權總分校高標"] = analytics[key + "^^^高標"];
                                        row["加權總分校均標"] = analytics[key + "^^^均標"];
                                        row["加權總分校低標"] = analytics[key + "^^^低標"];
                                        row["加權總分校標準差"] = analytics[key + "^^^標準差"];
                                        row["加權總分校組距count90"] = analytics[key + "^^^count90"];
                                        row["加權總分校組距count80"] = analytics[key + "^^^count80"];
                                        row["加權總分校組距count70"] = analytics[key + "^^^count70"];
                                        row["加權總分校組距count60"] = analytics[key + "^^^count60"];
                                        row["加權總分校組距count50"] = analytics[key + "^^^count50"];
                                        row["加權總分校組距count40"] = analytics[key + "^^^count40"];
                                        row["加權總分校組距count30"] = analytics[key + "^^^count30"];
                                        row["加權總分校組距count20"] = analytics[key + "^^^count20"];
                                        row["加權總分校組距count10"] = analytics[key + "^^^count10"];
                                        row["加權總分校組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["加權總分校組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["加權總分校組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["加權總分校組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["加權總分校組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["加權總分校組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["加權總分校組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["加權總分校組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["加權總分校組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["加權總分校組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["加權總分校組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["加權總分校組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["加權總分校組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["加權總分校組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["加權總分校組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["加權總分校組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["加權總分校組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["加權總分校組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["加權總分校組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #endregion
                                   #region 加權平均
                                   ClassStuNum = 0;
                                   foreach (StudentRecord stuRec in classRec.Students)
                                   {
                                        ClassStuNum++;
                                        if (ClassStuNum > conf.StudentLimit)
                                             break;
                                        string studentID = stuRec.StudentID;
                                        if (studentPrintSubjectAvgW.ContainsKey(studentID))
                                        {
                                             row["加權平均" + ClassStuNum] = studentPrintSubjectAvgW[studentID];
                                             key = "加權平均班排名" + stuRec.RefClass.ClassID;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["加權平均班排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) + 1;
                                                  row["加權平均班排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                             key = "加權平均科排名" + stuRec.Department + "^^^" + gradeYear;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["加權平均科排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) + 1;
                                                  row["加權平均科排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                             key = "加權平均全校排名" + gradeYear;
                                             if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))//明確判斷學生是否參與排名
                                             {
                                                  row["加權平均全校排名" + ClassStuNum] = ranks[key].IndexOf(studentPrintSubjectAvgW[studentID]) + 1;
                                                  row["加權平均全校排名母數" + ClassStuNum] = ranks[key].Count;
                                             }
                                        }
                                   }
                                   #region 加權平均班組距及高低標
                                   key = "加權平均班排名" + classRec.ClassID;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["加權平均班高標"] = analytics[key + "^^^高標"];
                                        row["加權平均班均標"] = analytics[key + "^^^均標"];
                                        row["加權平均班低標"] = analytics[key + "^^^低標"];
                                        row["加權平均班標準差"] = analytics[key + "^^^標準差"];
                                        row["加權平均班組距count90"] = analytics[key + "^^^count90"];
                                        row["加權平均班組距count80"] = analytics[key + "^^^count80"];
                                        row["加權平均班組距count70"] = analytics[key + "^^^count70"];
                                        row["加權平均班組距count60"] = analytics[key + "^^^count60"];
                                        row["加權平均班組距count50"] = analytics[key + "^^^count50"];
                                        row["加權平均班組距count40"] = analytics[key + "^^^count40"];
                                        row["加權平均班組距count30"] = analytics[key + "^^^count30"];
                                        row["加權平均班組距count20"] = analytics[key + "^^^count20"];
                                        row["加權平均班組距count10"] = analytics[key + "^^^count10"];
                                        row["加權平均班組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["加權平均班組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["加權平均班組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["加權平均班組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["加權平均班組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["加權平均班組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["加權平均班組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["加權平均班組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["加權平均班組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["加權平均班組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["加權平均班組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["加權平均班組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["加權平均班組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["加權平均班組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["加權平均班組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["加權平均班組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["加權平均班組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["加權平均班組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["加權平均班組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #region 加權平均科組距及高低標
                                   key = "加權平均科排名" + classRec.Department + "^^^" + gradeYear;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["加權平均科高標"] = analytics[key + "^^^高標"];
                                        row["加權平均科均標"] = analytics[key + "^^^均標"];
                                        row["加權平均科低標"] = analytics[key + "^^^低標"];
                                        row["加權平均科標準差"] = analytics[key + "^^^標準差"];
                                        row["加權平均科組距count90"] = analytics[key + "^^^count90"];
                                        row["加權平均科組距count80"] = analytics[key + "^^^count80"];
                                        row["加權平均科組距count70"] = analytics[key + "^^^count70"];
                                        row["加權平均科組距count60"] = analytics[key + "^^^count60"];
                                        row["加權平均科組距count50"] = analytics[key + "^^^count50"];
                                        row["加權平均科組距count40"] = analytics[key + "^^^count40"];
                                        row["加權平均科組距count30"] = analytics[key + "^^^count30"];
                                        row["加權平均科組距count20"] = analytics[key + "^^^count20"];
                                        row["加權平均科組距count10"] = analytics[key + "^^^count10"];
                                        row["加權平均科組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["加權平均科組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["加權平均科組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["加權平均科組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["加權平均科組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["加權平均科組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["加權平均科組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["加權平均科組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["加權平均科組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["加權平均科組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["加權平均科組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["加權平均科組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["加權平均科組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["加權平均科組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["加權平均科組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["加權平均科組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["加權平均科組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["加權平均科組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["加權平均科組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #region 加權平均全校組距及高低標
                                   key = "加權平均全校排名" + gradeYear;
                                   if (rankStudents.ContainsKey(key))
                                   {
                                        row["加權平均校高標"] = analytics[key + "^^^高標"];
                                        row["加權平均校均標"] = analytics[key + "^^^均標"];
                                        row["加權平均校低標"] = analytics[key + "^^^低標"];
                                        row["加權平均校標準差"] = analytics[key + "^^^標準差"];
                                        row["加權平均校組距count90"] = analytics[key + "^^^count90"];
                                        row["加權平均校組距count80"] = analytics[key + "^^^count80"];
                                        row["加權平均校組距count70"] = analytics[key + "^^^count70"];
                                        row["加權平均校組距count60"] = analytics[key + "^^^count60"];
                                        row["加權平均校組距count50"] = analytics[key + "^^^count50"];
                                        row["加權平均校組距count40"] = analytics[key + "^^^count40"];
                                        row["加權平均校組距count30"] = analytics[key + "^^^count30"];
                                        row["加權平均校組距count20"] = analytics[key + "^^^count20"];
                                        row["加權平均校組距count10"] = analytics[key + "^^^count10"];
                                        row["加權平均校組距count100Up"] = analytics[key + "^^^count100Up"];
                                        row["加權平均校組距count90Up"] = analytics[key + "^^^count90Up"];
                                        row["加權平均校組距count80Up"] = analytics[key + "^^^count80Up"];
                                        row["加權平均校組距count70Up"] = analytics[key + "^^^count70Up"];
                                        row["加權平均校組距count60Up"] = analytics[key + "^^^count60Up"];
                                        row["加權平均校組距count50Up"] = analytics[key + "^^^count50Up"];
                                        row["加權平均校組距count40Up"] = analytics[key + "^^^count40Up"];
                                        row["加權平均校組距count30Up"] = analytics[key + "^^^count30Up"];
                                        row["加權平均校組距count20Up"] = analytics[key + "^^^count20Up"];
                                        row["加權平均校組距count10Up"] = analytics[key + "^^^count10Up"];
                                        row["加權平均校組距count90Down"] = analytics[key + "^^^count90Down"];
                                        row["加權平均校組距count80Down"] = analytics[key + "^^^count80Down"];
                                        row["加權平均校組距count70Down"] = analytics[key + "^^^count70Down"];
                                        row["加權平均校組距count60Down"] = analytics[key + "^^^count60Down"];
                                        row["加權平均校組距count50Down"] = analytics[key + "^^^count50Down"];
                                        row["加權平均校組距count40Down"] = analytics[key + "^^^count40Down"];
                                        row["加權平均校組距count30Down"] = analytics[key + "^^^count30Down"];
                                        row["加權平均校組距count20Down"] = analytics[key + "^^^count20Down"];
                                        row["加權平均校組距count10Down"] = analytics[key + "^^^count10Down"];
                                   }
                                   #endregion
                                   #endregion
                                   #region 類別1綜合成績
                                   ClassStuNum = 0;
                                   foreach (StudentRecord stuRec in classRec.Students)
                                   {
                                        ClassStuNum++;
                                        if (ClassStuNum > conf.StudentLimit)
                                             break;
                                        string studentID = stuRec.StudentID;
                                        if (studentTag1Group.ContainsKey(studentID))
                                        {
                                             if (studentTag1SubjectSum.ContainsKey(studentID))
                                             {
                                                  row["類別1總分" + ClassStuNum] = studentTag1SubjectSum[studentID];
                                                  key = "類別1總分排名" + "^^^" + gradeYear + "^^^" + studentTag1Group[studentID];
                                                  if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                  {
                                                       row["類別1總分排名" + ClassStuNum] = ranks[key].IndexOf(studentTag1SubjectSum[studentID]) + 1;
                                                       row["類別1總分排名母數" + ClassStuNum] = ranks[key].Count;
                                                  }
                                             }
                                             if (studentTag1SubjectAvg.ContainsKey(studentID))
                                             {
                                                  row["類別1平均" + ClassStuNum] = studentTag1SubjectAvg[studentID];
                                                  key = "類別1平均排名" + "^^^" + gradeYear + "^^^" + studentTag1Group[studentID];
                                                  if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                  {
                                                       row["類別1平均排名" + ClassStuNum] = ranks[key].IndexOf(studentTag1SubjectAvg[studentID]) + 1; ;
                                                       row["類別1平均排名母數" + ClassStuNum] = ranks[key].Count;
                                                  }
                                             }
                                             if (studentTag1SubjectSumW.ContainsKey(studentID))
                                             {
                                                  row["類別1加權總分" + ClassStuNum] = studentTag1SubjectSumW[studentID];
                                                  if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                  {
                                                       row["類別1加權總分排名" + ClassStuNum] = ranks[key].IndexOf(studentTag1SubjectSumW[studentID]) + 1; ;
                                                       row["類別1加權總分排名母數" + ClassStuNum] = ranks[key].Count;
                                                  }
                                             }
                                             if (studentTag1SubjectAvgW.ContainsKey(studentID))
                                             {
                                                  row["類別1加權平均" + ClassStuNum] = studentTag1SubjectAvgW[studentID];
                                                  key = "類別1加權平均排名" + "^^^" + gradeYear + "^^^" + studentTag1Group[studentID];
                                                  if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                  {
                                                       row["類別1加權平均排名" + ClassStuNum] = ranks[key].IndexOf(studentTag1SubjectAvgW[studentID]) + 1; ;
                                                       row["類別1加權平均排名母數" + ClassStuNum] = ranks[key].Count;
                                                  }
                                             }
                                        }
                                   }
                                   for (int i = 0; i < tag1List.Count; i++)
                                   {
                                        string tag = tag1List[i];
                                        #region 類別1總分組距及高低標
                                        key = "類別1總分排名" + "^^^" + gradeYear + "^^^" + tag;
                                        if (rankStudents.ContainsKey(key))
                                        {
                                             if (i == 0)
                                             {
                                                  row["類1總分高標"] = analytics[key + "^^^高標"];
                                                  row["類1總分均標"] = analytics[key + "^^^均標"];
                                                  row["類1總分低標"] = analytics[key + "^^^低標"];
                                                  row["類1總分標準差"] = analytics[key + "^^^標準差"];
                                                  row["類1總分組距count90"] = analytics[key + "^^^count90"];
                                                  row["類1總分組距count80"] = analytics[key + "^^^count80"];
                                                  row["類1總分組距count70"] = analytics[key + "^^^count70"];
                                                  row["類1總分組距count60"] = analytics[key + "^^^count60"];
                                                  row["類1總分組距count50"] = analytics[key + "^^^count50"];
                                                  row["類1總分組距count40"] = analytics[key + "^^^count40"];
                                                  row["類1總分組距count30"] = analytics[key + "^^^count30"];
                                                  row["類1總分組距count20"] = analytics[key + "^^^count20"];
                                                  row["類1總分組距count10"] = analytics[key + "^^^count10"];
                                                  row["類1總分組距count100Up"] = analytics[key + "^^^count100Up"];
                                                  row["類1總分組距count90Up"] = analytics[key + "^^^count90Up"];
                                                  row["類1總分組距count80Up"] = analytics[key + "^^^count80Up"];
                                                  row["類1總分組距count70Up"] = analytics[key + "^^^count70Up"];
                                                  row["類1總分組距count60Up"] = analytics[key + "^^^count60Up"];
                                                  row["類1總分組距count50Up"] = analytics[key + "^^^count50Up"];
                                                  row["類1總分組距count40Up"] = analytics[key + "^^^count40Up"];
                                                  row["類1總分組距count30Up"] = analytics[key + "^^^count30Up"];
                                                  row["類1總分組距count20Up"] = analytics[key + "^^^count20Up"];
                                                  row["類1總分組距count10Up"] = analytics[key + "^^^count10Up"];
                                                  row["類1總分組距count90Down"] = analytics[key + "^^^count90Down"];
                                                  row["類1總分組距count80Down"] = analytics[key + "^^^count80Down"];
                                                  row["類1總分組距count70Down"] = analytics[key + "^^^count70Down"];
                                                  row["類1總分組距count60Down"] = analytics[key + "^^^count60Down"];
                                                  row["類1總分組距count50Down"] = analytics[key + "^^^count50Down"];
                                                  row["類1總分組距count40Down"] = analytics[key + "^^^count40Down"];
                                                  row["類1總分組距count30Down"] = analytics[key + "^^^count30Down"];
                                                  row["類1總分組距count20Down"] = analytics[key + "^^^count20Down"];
                                                  row["類1總分組距count10Down"] = analytics[key + "^^^count10Down"];
                                             }
                                             else
                                             {
                                                  row["類1總分高標"] += "/" + analytics[key + "^^^高標"];
                                                  row["類1總分均標"] += "/" + analytics[key + "^^^均標"];
                                                  row["類1總分低標"] += "/" + analytics[key + "^^^低標"];
                                                  row["類1總分標準差"] += "/" + analytics[key + "^^^標準差"];
                                                  row["類1總分組距count90"] += "/" + analytics[key + "^^^count90"];
                                                  row["類1總分組距count80"] += "/" + analytics[key + "^^^count80"];
                                                  row["類1總分組距count70"] += "/" + analytics[key + "^^^count70"];
                                                  row["類1總分組距count60"] += "/" + analytics[key + "^^^count60"];
                                                  row["類1總分組距count50"] += "/" + analytics[key + "^^^count50"];
                                                  row["類1總分組距count40"] += "/" + analytics[key + "^^^count40"];
                                                  row["類1總分組距count30"] += "/" + analytics[key + "^^^count30"];
                                                  row["類1總分組距count20"] += "/" + analytics[key + "^^^count20"];
                                                  row["類1總分組距count10"] += "/" + analytics[key + "^^^count10"];
                                                  row["類1總分組距count100Up"] += "/" + analytics[key + "^^^count100Up"];
                                                  row["類1總分組距count90Up"] += "/" + analytics[key + "^^^count90Up"];
                                                  row["類1總分組距count80Up"] += "/" + analytics[key + "^^^count80Up"];
                                                  row["類1總分組距count70Up"] += "/" + analytics[key + "^^^count70Up"];
                                                  row["類1總分組距count60Up"] += "/" + analytics[key + "^^^count60Up"];
                                                  row["類1總分組距count50Up"] += "/" + analytics[key + "^^^count50Up"];
                                                  row["類1總分組距count40Up"] += "/" + analytics[key + "^^^count40Up"];
                                                  row["類1總分組距count30Up"] += "/" + analytics[key + "^^^count30Up"];
                                                  row["類1總分組距count20Up"] += "/" + analytics[key + "^^^count20Up"];
                                                  row["類1總分組距count10Up"] += "/" + analytics[key + "^^^count10Up"];
                                                  row["類1總分組距count90Down"] += "/" + analytics[key + "^^^count90Down"];
                                                  row["類1總分組距count80Down"] += "/" + analytics[key + "^^^count80Down"];
                                                  row["類1總分組距count70Down"] += "/" + analytics[key + "^^^count70Down"];
                                                  row["類1總分組距count60Down"] += "/" + analytics[key + "^^^count60Down"];
                                                  row["類1總分組距count50Down"] += "/" + analytics[key + "^^^count50Down"];
                                                  row["類1總分組距count40Down"] += "/" + analytics[key + "^^^count40Down"];
                                                  row["類1總分組距count30Down"] += "/" + analytics[key + "^^^count30Down"];
                                                  row["類1總分組距count20Down"] += "/" + analytics[key + "^^^count20Down"];
                                                  row["類1總分組距count10Down"] += "/" + analytics[key + "^^^count10Down"];
                                             }
                                        }
                                        #endregion
                                        #region 類別1平均組距及高低標
                                        key = "類別1平均排名" + "^^^" + gradeYear + "^^^" + tag;
                                        if (rankStudents.ContainsKey(key))
                                        {
                                             if (i == 0)
                                             {
                                                  row["類1平均高標"] = analytics[key + "^^^高標"];
                                                  row["類1平均均標"] = analytics[key + "^^^均標"];
                                                  row["類1平均低標"] = analytics[key + "^^^低標"];
                                                  row["類1平均標準差"] = analytics[key + "^^^標準差"];
                                                  row["類1平均組距count90"] = analytics[key + "^^^count90"];
                                                  row["類1平均組距count80"] = analytics[key + "^^^count80"];
                                                  row["類1平均組距count70"] = analytics[key + "^^^count70"];
                                                  row["類1平均組距count60"] = analytics[key + "^^^count60"];
                                                  row["類1平均組距count50"] = analytics[key + "^^^count50"];
                                                  row["類1平均組距count40"] = analytics[key + "^^^count40"];
                                                  row["類1平均組距count30"] = analytics[key + "^^^count30"];
                                                  row["類1平均組距count20"] = analytics[key + "^^^count20"];
                                                  row["類1平均組距count10"] = analytics[key + "^^^count10"];
                                                  row["類1平均組距count100Up"] = analytics[key + "^^^count100Up"];
                                                  row["類1平均組距count90Up"] = analytics[key + "^^^count90Up"];
                                                  row["類1平均組距count80Up"] = analytics[key + "^^^count80Up"];
                                                  row["類1平均組距count70Up"] = analytics[key + "^^^count70Up"];
                                                  row["類1平均組距count60Up"] = analytics[key + "^^^count60Up"];
                                                  row["類1平均組距count50Up"] = analytics[key + "^^^count50Up"];
                                                  row["類1平均組距count40Up"] = analytics[key + "^^^count40Up"];
                                                  row["類1平均組距count30Up"] = analytics[key + "^^^count30Up"];
                                                  row["類1平均組距count20Up"] = analytics[key + "^^^count20Up"];
                                                  row["類1平均組距count10Up"] = analytics[key + "^^^count10Up"];
                                                  row["類1平均組距count90Down"] = analytics[key + "^^^count90Down"];
                                                  row["類1平均組距count80Down"] = analytics[key + "^^^count80Down"];
                                                  row["類1平均組距count70Down"] = analytics[key + "^^^count70Down"];
                                                  row["類1平均組距count60Down"] = analytics[key + "^^^count60Down"];
                                                  row["類1平均組距count50Down"] = analytics[key + "^^^count50Down"];
                                                  row["類1平均組距count40Down"] = analytics[key + "^^^count40Down"];
                                                  row["類1平均組距count30Down"] = analytics[key + "^^^count30Down"];
                                                  row["類1平均組距count20Down"] = analytics[key + "^^^count20Down"];
                                                  row["類1平均組距count10Down"] = analytics[key + "^^^count10Down"];
                                             }
                                             else
                                             {
                                                  row["類1平均高標"] += "/" + analytics[key + "^^^高標"];
                                                  row["類1平均均標"] += "/" + analytics[key + "^^^均標"];
                                                  row["類1平均低標"] += "/" + analytics[key + "^^^低標"];
                                                  row["類1平均標準差"] += "/" + analytics[key + "^^^標準差"];
                                                  row["類1平均組距count90"] += "/" + analytics[key + "^^^count90"];
                                                  row["類1平均組距count80"] += "/" + analytics[key + "^^^count80"];
                                                  row["類1平均組距count70"] += "/" + analytics[key + "^^^count70"];
                                                  row["類1平均組距count60"] += "/" + analytics[key + "^^^count60"];
                                                  row["類1平均組距count50"] += "/" + analytics[key + "^^^count50"];
                                                  row["類1平均組距count40"] += "/" + analytics[key + "^^^count40"];
                                                  row["類1平均組距count30"] += "/" + analytics[key + "^^^count30"];
                                                  row["類1平均組距count20"] += "/" + analytics[key + "^^^count20"];
                                                  row["類1平均組距count10"] += "/" + analytics[key + "^^^count10"];
                                                  row["類1平均組距count100Up"] += "/" + analytics[key + "^^^count100Up"];
                                                  row["類1平均組距count90Up"] += "/" + analytics[key + "^^^count90Up"];
                                                  row["類1平均組距count80Up"] += "/" + analytics[key + "^^^count80Up"];
                                                  row["類1平均組距count70Up"] += "/" + analytics[key + "^^^count70Up"];
                                                  row["類1平均組距count60Up"] += "/" + analytics[key + "^^^count60Up"];
                                                  row["類1平均組距count50Up"] += "/" + analytics[key + "^^^count50Up"];
                                                  row["類1平均組距count40Up"] += "/" + analytics[key + "^^^count40Up"];
                                                  row["類1平均組距count30Up"] += "/" + analytics[key + "^^^count30Up"];
                                                  row["類1平均組距count20Up"] += "/" + analytics[key + "^^^count20Up"];
                                                  row["類1平均組距count10Up"] += "/" + analytics[key + "^^^count10Up"];
                                                  row["類1平均組距count90Down"] += "/" + analytics[key + "^^^count90Down"];
                                                  row["類1平均組距count80Down"] += "/" + analytics[key + "^^^count80Down"];
                                                  row["類1平均組距count70Down"] += "/" + analytics[key + "^^^count70Down"];
                                                  row["類1平均組距count60Down"] += "/" + analytics[key + "^^^count60Down"];
                                                  row["類1平均組距count50Down"] += "/" + analytics[key + "^^^count50Down"];
                                                  row["類1平均組距count40Down"] += "/" + analytics[key + "^^^count40Down"];
                                                  row["類1平均組距count30Down"] += "/" + analytics[key + "^^^count30Down"];
                                                  row["類1平均組距count20Down"] += "/" + analytics[key + "^^^count20Down"];
                                                  row["類1平均組距count10Down"] += "/" + analytics[key + "^^^count10Down"];
                                             }
                                        }
                                        #endregion
                                        #region 類別1加權總分組距及高低標
                                        key = "類別1加權總分排名" + "^^^" + gradeYear + "^^^" + tag;
                                        if (rankStudents.ContainsKey(key))
                                        {
                                             if (i == 0)
                                             {
                                                  row["類1加權總分高標"] = analytics[key + "^^^高標"];
                                                  row["類1加權總分均標"] = analytics[key + "^^^均標"];
                                                  row["類1加權總分低標"] = analytics[key + "^^^低標"];
                                                  row["類1加權總分標準差"] = analytics[key + "^^^標準差"];
                                                  row["類1加權總分組距count90"] = analytics[key + "^^^count90"];
                                                  row["類1加權總分組距count80"] = analytics[key + "^^^count80"];
                                                  row["類1加權總分組距count70"] = analytics[key + "^^^count70"];
                                                  row["類1加權總分組距count60"] = analytics[key + "^^^count60"];
                                                  row["類1加權總分組距count50"] = analytics[key + "^^^count50"];
                                                  row["類1加權總分組距count40"] = analytics[key + "^^^count40"];
                                                  row["類1加權總分組距count30"] = analytics[key + "^^^count30"];
                                                  row["類1加權總分組距count20"] = analytics[key + "^^^count20"];
                                                  row["類1加權總分組距count10"] = analytics[key + "^^^count10"];
                                                  row["類1加權總分組距count100Up"] = analytics[key + "^^^count100Up"];
                                                  row["類1加權總分組距count90Up"] = analytics[key + "^^^count90Up"];
                                                  row["類1加權總分組距count80Up"] = analytics[key + "^^^count80Up"];
                                                  row["類1加權總分組距count70Up"] = analytics[key + "^^^count70Up"];
                                                  row["類1加權總分組距count60Up"] = analytics[key + "^^^count60Up"];
                                                  row["類1加權總分組距count50Up"] = analytics[key + "^^^count50Up"];
                                                  row["類1加權總分組距count40Up"] = analytics[key + "^^^count40Up"];
                                                  row["類1加權總分組距count30Up"] = analytics[key + "^^^count30Up"];
                                                  row["類1加權總分組距count20Up"] = analytics[key + "^^^count20Up"];
                                                  row["類1加權總分組距count10Up"] = analytics[key + "^^^count10Up"];
                                                  row["類1加權總分組距count90Down"] = analytics[key + "^^^count90Down"];
                                                  row["類1加權總分組距count80Down"] = analytics[key + "^^^count80Down"];
                                                  row["類1加權總分組距count70Down"] = analytics[key + "^^^count70Down"];
                                                  row["類1加權總分組距count60Down"] = analytics[key + "^^^count60Down"];
                                                  row["類1加權總分組距count50Down"] = analytics[key + "^^^count50Down"];
                                                  row["類1加權總分組距count40Down"] = analytics[key + "^^^count40Down"];
                                                  row["類1加權總分組距count30Down"] = analytics[key + "^^^count30Down"];
                                                  row["類1加權總分組距count20Down"] = analytics[key + "^^^count20Down"];
                                                  row["類1加權總分組距count10Down"] = analytics[key + "^^^count10Down"];
                                             }
                                             else
                                             {
                                                  row["類1加權總分高標"] += "/" + analytics[key + "^^^高標"];
                                                  row["類1加權總分均標"] += "/" + analytics[key + "^^^均標"];
                                                  row["類1加權總分低標"] += "/" + analytics[key + "^^^低標"];
                                                  row["類1加權總分標準差"] += "/" + analytics[key + "^^^標準差"];
                                                  row["類1加權總分組距count90"] += "/" + analytics[key + "^^^count90"];
                                                  row["類1加權總分組距count80"] += "/" + analytics[key + "^^^count80"];
                                                  row["類1加權總分組距count70"] += "/" + analytics[key + "^^^count70"];
                                                  row["類1加權總分組距count60"] += "/" + analytics[key + "^^^count60"];
                                                  row["類1加權總分組距count50"] += "/" + analytics[key + "^^^count50"];
                                                  row["類1加權總分組距count40"] += "/" + analytics[key + "^^^count40"];
                                                  row["類1加權總分組距count30"] += "/" + analytics[key + "^^^count30"];
                                                  row["類1加權總分組距count20"] += "/" + analytics[key + "^^^count20"];
                                                  row["類1加權總分組距count10"] += "/" + analytics[key + "^^^count10"];
                                                  row["類1加權總分組距count100Up"] += "/" + analytics[key + "^^^count100Up"];
                                                  row["類1加權總分組距count90Up"] += "/" + analytics[key + "^^^count90Up"];
                                                  row["類1加權總分組距count80Up"] += "/" + analytics[key + "^^^count80Up"];
                                                  row["類1加權總分組距count70Up"] += "/" + analytics[key + "^^^count70Up"];
                                                  row["類1加權總分組距count60Up"] += "/" + analytics[key + "^^^count60Up"];
                                                  row["類1加權總分組距count50Up"] += "/" + analytics[key + "^^^count50Up"];
                                                  row["類1加權總分組距count40Up"] += "/" + analytics[key + "^^^count40Up"];
                                                  row["類1加權總分組距count30Up"] += "/" + analytics[key + "^^^count30Up"];
                                                  row["類1加權總分組距count20Up"] += "/" + analytics[key + "^^^count20Up"];
                                                  row["類1加權總分組距count10Up"] += "/" + analytics[key + "^^^count10Up"];
                                                  row["類1加權總分組距count90Down"] += "/" + analytics[key + "^^^count90Down"];
                                                  row["類1加權總分組距count80Down"] += "/" + analytics[key + "^^^count80Down"];
                                                  row["類1加權總分組距count70Down"] += "/" + analytics[key + "^^^count70Down"];
                                                  row["類1加權總分組距count60Down"] += "/" + analytics[key + "^^^count60Down"];
                                                  row["類1加權總分組距count50Down"] += "/" + analytics[key + "^^^count50Down"];
                                                  row["類1加權總分組距count40Down"] += "/" + analytics[key + "^^^count40Down"];
                                                  row["類1加權總分組距count30Down"] += "/" + analytics[key + "^^^count30Down"];
                                                  row["類1加權總分組距count20Down"] += "/" + analytics[key + "^^^count20Down"];
                                                  row["類1加權總分組距count10Down"] += "/" + analytics[key + "^^^count10Down"];
                                             }
                                        }
                                        #endregion
                                        #region 類別1加權平均組距及高低標
                                        key = "類別1加權平均排名" + "^^^" + gradeYear + "^^^" + tag;
                                        if (rankStudents.ContainsKey(key))
                                        {
                                             if (i == 0)
                                             {
                                                  row["類1加權平均高標"] = analytics[key + "^^^高標"];
                                                  row["類1加權平均均標"] = analytics[key + "^^^均標"];
                                                  row["類1加權平均低標"] = analytics[key + "^^^低標"];
                                                  row["類1加權平均標準差"] = analytics[key + "^^^標準差"];
                                                  row["類1加權平均組距count90"] = analytics[key + "^^^count90"];
                                                  row["類1加權平均組距count80"] = analytics[key + "^^^count80"];
                                                  row["類1加權平均組距count70"] = analytics[key + "^^^count70"];
                                                  row["類1加權平均組距count60"] = analytics[key + "^^^count60"];
                                                  row["類1加權平均組距count50"] = analytics[key + "^^^count50"];
                                                  row["類1加權平均組距count40"] = analytics[key + "^^^count40"];
                                                  row["類1加權平均組距count30"] = analytics[key + "^^^count30"];
                                                  row["類1加權平均組距count20"] = analytics[key + "^^^count20"];
                                                  row["類1加權平均組距count10"] = analytics[key + "^^^count10"];
                                                  row["類1加權平均組距count100Up"] = analytics[key + "^^^count100Up"];
                                                  row["類1加權平均組距count90Up"] = analytics[key + "^^^count90Up"];
                                                  row["類1加權平均組距count80Up"] = analytics[key + "^^^count80Up"];
                                                  row["類1加權平均組距count70Up"] = analytics[key + "^^^count70Up"];
                                                  row["類1加權平均組距count60Up"] = analytics[key + "^^^count60Up"];
                                                  row["類1加權平均組距count50Up"] = analytics[key + "^^^count50Up"];
                                                  row["類1加權平均組距count40Up"] = analytics[key + "^^^count40Up"];
                                                  row["類1加權平均組距count30Up"] = analytics[key + "^^^count30Up"];
                                                  row["類1加權平均組距count20Up"] = analytics[key + "^^^count20Up"];
                                                  row["類1加權平均組距count10Up"] = analytics[key + "^^^count10Up"];
                                                  row["類1加權平均組距count90Down"] = analytics[key + "^^^count90Down"];
                                                  row["類1加權平均組距count80Down"] = analytics[key + "^^^count80Down"];
                                                  row["類1加權平均組距count70Down"] = analytics[key + "^^^count70Down"];
                                                  row["類1加權平均組距count60Down"] = analytics[key + "^^^count60Down"];
                                                  row["類1加權平均組距count50Down"] = analytics[key + "^^^count50Down"];
                                                  row["類1加權平均組距count40Down"] = analytics[key + "^^^count40Down"];
                                                  row["類1加權平均組距count30Down"] = analytics[key + "^^^count30Down"];
                                                  row["類1加權平均組距count20Down"] = analytics[key + "^^^count20Down"];
                                                  row["類1加權平均組距count10Down"] = analytics[key + "^^^count10Down"];
                                             }
                                             else
                                             {
                                                  row["類1加權平均高標"] += "/" + analytics[key + "^^^高標"];
                                                  row["類1加權平均均標"] += "/" + analytics[key + "^^^均標"];
                                                  row["類1加權平均低標"] += "/" + analytics[key + "^^^低標"];
                                                  row["類1加權平均標準差"] += "/" + analytics[key + "^^^標準差"];
                                                  row["類1加權平均組距count90"] += "/" + analytics[key + "^^^count90"];
                                                  row["類1加權平均組距count80"] += "/" + analytics[key + "^^^count80"];
                                                  row["類1加權平均組距count70"] += "/" + analytics[key + "^^^count70"];
                                                  row["類1加權平均組距count60"] += "/" + analytics[key + "^^^count60"];
                                                  row["類1加權平均組距count50"] += "/" + analytics[key + "^^^count50"];
                                                  row["類1加權平均組距count40"] += "/" + analytics[key + "^^^count40"];
                                                  row["類1加權平均組距count30"] += "/" + analytics[key + "^^^count30"];
                                                  row["類1加權平均組距count20"] += "/" + analytics[key + "^^^count20"];
                                                  row["類1加權平均組距count10"] += "/" + analytics[key + "^^^count10"];
                                                  row["類1加權平均組距count100Up"] += "/" + analytics[key + "^^^count100Up"];
                                                  row["類1加權平均組距count90Up"] += "/" + analytics[key + "^^^count90Up"];
                                                  row["類1加權平均組距count80Up"] += "/" + analytics[key + "^^^count80Up"];
                                                  row["類1加權平均組距count70Up"] += "/" + analytics[key + "^^^count70Up"];
                                                  row["類1加權平均組距count60Up"] += "/" + analytics[key + "^^^count60Up"];
                                                  row["類1加權平均組距count50Up"] += "/" + analytics[key + "^^^count50Up"];
                                                  row["類1加權平均組距count40Up"] += "/" + analytics[key + "^^^count40Up"];
                                                  row["類1加權平均組距count30Up"] += "/" + analytics[key + "^^^count30Up"];
                                                  row["類1加權平均組距count20Up"] += "/" + analytics[key + "^^^count20Up"];
                                                  row["類1加權平均組距count10Up"] += "/" + analytics[key + "^^^count10Up"];
                                                  row["類1加權平均組距count90Down"] += "/" + analytics[key + "^^^count90Down"];
                                                  row["類1加權平均組距count80Down"] += "/" + analytics[key + "^^^count80Down"];
                                                  row["類1加權平均組距count70Down"] += "/" + analytics[key + "^^^count70Down"];
                                                  row["類1加權平均組距count60Down"] += "/" + analytics[key + "^^^count60Down"];
                                                  row["類1加權平均組距count50Down"] += "/" + analytics[key + "^^^count50Down"];
                                                  row["類1加權平均組距count40Down"] += "/" + analytics[key + "^^^count40Down"];
                                                  row["類1加權平均組距count30Down"] += "/" + analytics[key + "^^^count30Down"];
                                                  row["類1加權平均組距count20Down"] += "/" + analytics[key + "^^^count20Down"];
                                                  row["類1加權平均組距count10Down"] += "/" + analytics[key + "^^^count10Down"];
                                             }
                                        }
                                        #endregion
                                   }
                                   #endregion
                                   #region 類別2綜合成績
                                   ClassStuNum = 0;
                                   foreach (StudentRecord stuRec in classRec.Students)
                                   {
                                        ClassStuNum++;
                                        if (ClassStuNum > conf.StudentLimit)
                                             break;
                                        string studentID = stuRec.StudentID;
                                        if (studentTag2Group.ContainsKey(studentID))
                                        {
                                             if (studentTag2SubjectSum.ContainsKey(studentID))
                                             {
                                                  row["類別2總分" + ClassStuNum] = studentTag2SubjectSum[studentID];
                                                  key = "類別2總分排名" + "^^^" + gradeYear + "^^^" + studentTag2Group[studentID];
                                                  if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                  {
                                                       row["類別2總分排名" + ClassStuNum] = ranks[key].IndexOf(studentTag2SubjectSum[studentID]) + 1;
                                                       row["類別2總分排名母數" + ClassStuNum] = ranks[key].Count;
                                                  }
                                             }
                                             if (studentTag2SubjectAvg.ContainsKey(studentID))
                                             {
                                                  row["類別2平均" + ClassStuNum] = studentTag2SubjectAvg[studentID];
                                                  key = "類別2平均排名" + "^^^" + gradeYear + "^^^" + studentTag2Group[studentID];
                                                  if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                  {
                                                       row["類別2平均排名" + ClassStuNum] = ranks[key].IndexOf(studentTag2SubjectAvg[studentID]) + 1; ;
                                                       row["類別2平均排名母數" + ClassStuNum] = ranks[key].Count;
                                                  }
                                             }
                                             if (studentTag2SubjectSumW.ContainsKey(studentID))
                                             {
                                                  row["類別2加權總分" + ClassStuNum] = studentTag2SubjectSumW[studentID];
                                                  key = "類別2加權總分排名" + "^^^" + gradeYear + "^^^" + studentTag2Group[studentID];
                                                  if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                  {
                                                       row["類別2加權總分排名" + ClassStuNum] = ranks[key].IndexOf(studentTag2SubjectSumW[studentID]) + 1; ;
                                                       row["類別2加權總分排名母數" + ClassStuNum] = ranks[key].Count;
                                                  }
                                             }
                                             if (studentTag2SubjectAvgW.ContainsKey(studentID))
                                             {
                                                  row["類別2加權平均" + ClassStuNum] = studentTag2SubjectAvgW[studentID];
                                                  key = "類別2加權平均排名" + "^^^" + gradeYear + "^^^" + studentTag2Group[studentID];
                                                  if (rankStudents.ContainsKey(key) && rankStudents[key].Contains(studentID))
                                                  {
                                                       row["類別2加權平均排名" + ClassStuNum] = ranks[key].IndexOf(studentTag2SubjectAvgW[studentID]) + 1; ;
                                                       row["類別2加權平均排名母數" + ClassStuNum] = ranks[key].Count;
                                                  }
                                             }
                                        }
                                   }
                                   for (int i = 0; i < tag2List.Count; i++)
                                   {
                                        string tag = tag2List[i];
                                        #region 類別2總分組距及高低標
                                        key = "類別2總分排名" + "^^^" + gradeYear + "^^^" + tag;
                                        if (rankStudents.ContainsKey(key))
                                        {
                                             if (i == 0)
                                             {
                                                  row["類2總分高標"] = analytics[key + "^^^高標"];
                                                  row["類2總分均標"] = analytics[key + "^^^均標"];
                                                  row["類2總分低標"] = analytics[key + "^^^低標"];
                                                  row["類2總分標準差"] = analytics[key + "^^^標準差"];
                                                  row["類2總分組距count90"] = analytics[key + "^^^count90"];
                                                  row["類2總分組距count80"] = analytics[key + "^^^count80"];
                                                  row["類2總分組距count70"] = analytics[key + "^^^count70"];
                                                  row["類2總分組距count60"] = analytics[key + "^^^count60"];
                                                  row["類2總分組距count50"] = analytics[key + "^^^count50"];
                                                  row["類2總分組距count40"] = analytics[key + "^^^count40"];
                                                  row["類2總分組距count30"] = analytics[key + "^^^count30"];
                                                  row["類2總分組距count20"] = analytics[key + "^^^count20"];
                                                  row["類2總分組距count10"] = analytics[key + "^^^count10"];
                                                  row["類2總分組距count100Up"] = analytics[key + "^^^count100Up"];
                                                  row["類2總分組距count90Up"] = analytics[key + "^^^count90Up"];
                                                  row["類2總分組距count80Up"] = analytics[key + "^^^count80Up"];
                                                  row["類2總分組距count70Up"] = analytics[key + "^^^count70Up"];
                                                  row["類2總分組距count60Up"] = analytics[key + "^^^count60Up"];
                                                  row["類2總分組距count50Up"] = analytics[key + "^^^count50Up"];
                                                  row["類2總分組距count40Up"] = analytics[key + "^^^count40Up"];
                                                  row["類2總分組距count30Up"] = analytics[key + "^^^count30Up"];
                                                  row["類2總分組距count20Up"] = analytics[key + "^^^count20Up"];
                                                  row["類2總分組距count10Up"] = analytics[key + "^^^count10Up"];
                                                  row["類2總分組距count90Down"] = analytics[key + "^^^count90Down"];
                                                  row["類2總分組距count80Down"] = analytics[key + "^^^count80Down"];
                                                  row["類2總分組距count70Down"] = analytics[key + "^^^count70Down"];
                                                  row["類2總分組距count60Down"] = analytics[key + "^^^count60Down"];
                                                  row["類2總分組距count50Down"] = analytics[key + "^^^count50Down"];
                                                  row["類2總分組距count40Down"] = analytics[key + "^^^count40Down"];
                                                  row["類2總分組距count30Down"] = analytics[key + "^^^count30Down"];
                                                  row["類2總分組距count20Down"] = analytics[key + "^^^count20Down"];
                                                  row["類2總分組距count10Down"] = analytics[key + "^^^count10Down"];
                                             }
                                             else
                                             {
                                                  row["類2總分高標"] += "/" + analytics[key + "^^^高標"];
                                                  row["類2總分均標"] += "/" + analytics[key + "^^^均標"];
                                                  row["類2總分低標"] += "/" + analytics[key + "^^^低標"];
                                                  row["類2總分標準差"] += "/" + analytics[key + "^^^標準差"];
                                                  row["類2總分組距count90"] += "/" + analytics[key + "^^^count90"];
                                                  row["類2總分組距count80"] += "/" + analytics[key + "^^^count80"];
                                                  row["類2總分組距count70"] += "/" + analytics[key + "^^^count70"];
                                                  row["類2總分組距count60"] += "/" + analytics[key + "^^^count60"];
                                                  row["類2總分組距count50"] += "/" + analytics[key + "^^^count50"];
                                                  row["類2總分組距count40"] += "/" + analytics[key + "^^^count40"];
                                                  row["類2總分組距count30"] += "/" + analytics[key + "^^^count30"];
                                                  row["類2總分組距count20"] += "/" + analytics[key + "^^^count20"];
                                                  row["類2總分組距count10"] += "/" + analytics[key + "^^^count10"];
                                                  row["類2總分組距count100Up"] += "/" + analytics[key + "^^^count100Up"];
                                                  row["類2總分組距count90Up"] += "/" + analytics[key + "^^^count90Up"];
                                                  row["類2總分組距count80Up"] += "/" + analytics[key + "^^^count80Up"];
                                                  row["類2總分組距count70Up"] += "/" + analytics[key + "^^^count70Up"];
                                                  row["類2總分組距count60Up"] += "/" + analytics[key + "^^^count60Up"];
                                                  row["類2總分組距count50Up"] += "/" + analytics[key + "^^^count50Up"];
                                                  row["類2總分組距count40Up"] += "/" + analytics[key + "^^^count40Up"];
                                                  row["類2總分組距count30Up"] += "/" + analytics[key + "^^^count30Up"];
                                                  row["類2總分組距count20Up"] += "/" + analytics[key + "^^^count20Up"];
                                                  row["類2總分組距count10Up"] += "/" + analytics[key + "^^^count10Up"];
                                                  row["類2總分組距count90Down"] += "/" + analytics[key + "^^^count90Down"];
                                                  row["類2總分組距count80Down"] += "/" + analytics[key + "^^^count80Down"];
                                                  row["類2總分組距count70Down"] += "/" + analytics[key + "^^^count70Down"];
                                                  row["類2總分組距count60Down"] += "/" + analytics[key + "^^^count60Down"];
                                                  row["類2總分組距count50Down"] += "/" + analytics[key + "^^^count50Down"];
                                                  row["類2總分組距count40Down"] += "/" + analytics[key + "^^^count40Down"];
                                                  row["類2總分組距count30Down"] += "/" + analytics[key + "^^^count30Down"];
                                                  row["類2總分組距count20Down"] += "/" + analytics[key + "^^^count20Down"];
                                                  row["類2總分組距count10Down"] += "/" + analytics[key + "^^^count10Down"];
                                             }
                                        }
                                        #endregion
                                        #region 類別2平均組距及高低標
                                        key = "類別2平均排名" + "^^^" + gradeYear + "^^^" + tag;
                                        if (rankStudents.ContainsKey(key))
                                        {
                                             if (i == 0)
                                             {
                                                  row["類2平均高標"] = analytics[key + "^^^高標"];
                                                  row["類2平均均標"] = analytics[key + "^^^均標"];
                                                  row["類2平均低標"] = analytics[key + "^^^低標"];
                                                  row["類2平均標準差"] = analytics[key + "^^^標準差"];
                                                  row["類2平均組距count90"] = analytics[key + "^^^count90"];
                                                  row["類2平均組距count80"] = analytics[key + "^^^count80"];
                                                  row["類2平均組距count70"] = analytics[key + "^^^count70"];
                                                  row["類2平均組距count60"] = analytics[key + "^^^count60"];
                                                  row["類2平均組距count50"] = analytics[key + "^^^count50"];
                                                  row["類2平均組距count40"] = analytics[key + "^^^count40"];
                                                  row["類2平均組距count30"] = analytics[key + "^^^count30"];
                                                  row["類2平均組距count20"] = analytics[key + "^^^count20"];
                                                  row["類2平均組距count10"] = analytics[key + "^^^count10"];
                                                  row["類2平均組距count100Up"] = analytics[key + "^^^count100Up"];
                                                  row["類2平均組距count90Up"] = analytics[key + "^^^count90Up"];
                                                  row["類2平均組距count80Up"] = analytics[key + "^^^count80Up"];
                                                  row["類2平均組距count70Up"] = analytics[key + "^^^count70Up"];
                                                  row["類2平均組距count60Up"] = analytics[key + "^^^count60Up"];
                                                  row["類2平均組距count50Up"] = analytics[key + "^^^count50Up"];
                                                  row["類2平均組距count40Up"] = analytics[key + "^^^count40Up"];
                                                  row["類2平均組距count30Up"] = analytics[key + "^^^count30Up"];
                                                  row["類2平均組距count20Up"] = analytics[key + "^^^count20Up"];
                                                  row["類2平均組距count10Up"] = analytics[key + "^^^count10Up"];
                                                  row["類2平均組距count90Down"] = analytics[key + "^^^count90Down"];
                                                  row["類2平均組距count80Down"] = analytics[key + "^^^count80Down"];
                                                  row["類2平均組距count70Down"] = analytics[key + "^^^count70Down"];
                                                  row["類2平均組距count60Down"] = analytics[key + "^^^count60Down"];
                                                  row["類2平均組距count50Down"] = analytics[key + "^^^count50Down"];
                                                  row["類2平均組距count40Down"] = analytics[key + "^^^count40Down"];
                                                  row["類2平均組距count30Down"] = analytics[key + "^^^count30Down"];
                                                  row["類2平均組距count20Down"] = analytics[key + "^^^count20Down"];
                                                  row["類2平均組距count10Down"] = analytics[key + "^^^count10Down"];
                                             }
                                             else
                                             {
                                                  row["類2平均高標"] += "/" + analytics[key + "^^^高標"];
                                                  row["類2平均均標"] += "/" + analytics[key + "^^^均標"];
                                                  row["類2平均低標"] += "/" + analytics[key + "^^^低標"];
                                                  row["類2平均標準差"] += "/" + analytics[key + "^^^標準差"];
                                                  row["類2平均組距count90"] += "/" + analytics[key + "^^^count90"];
                                                  row["類2平均組距count80"] += "/" + analytics[key + "^^^count80"];
                                                  row["類2平均組距count70"] += "/" + analytics[key + "^^^count70"];
                                                  row["類2平均組距count60"] += "/" + analytics[key + "^^^count60"];
                                                  row["類2平均組距count50"] += "/" + analytics[key + "^^^count50"];
                                                  row["類2平均組距count40"] += "/" + analytics[key + "^^^count40"];
                                                  row["類2平均組距count30"] += "/" + analytics[key + "^^^count30"];
                                                  row["類2平均組距count20"] += "/" + analytics[key + "^^^count20"];
                                                  row["類2平均組距count10"] += "/" + analytics[key + "^^^count10"];
                                                  row["類2平均組距count100Up"] += "/" + analytics[key + "^^^count100Up"];
                                                  row["類2平均組距count90Up"] += "/" + analytics[key + "^^^count90Up"];
                                                  row["類2平均組距count80Up"] += "/" + analytics[key + "^^^count80Up"];
                                                  row["類2平均組距count70Up"] += "/" + analytics[key + "^^^count70Up"];
                                                  row["類2平均組距count60Up"] += "/" + analytics[key + "^^^count60Up"];
                                                  row["類2平均組距count50Up"] += "/" + analytics[key + "^^^count50Up"];
                                                  row["類2平均組距count40Up"] += "/" + analytics[key + "^^^count40Up"];
                                                  row["類2平均組距count30Up"] += "/" + analytics[key + "^^^count30Up"];
                                                  row["類2平均組距count20Up"] += "/" + analytics[key + "^^^count20Up"];
                                                  row["類2平均組距count10Up"] += "/" + analytics[key + "^^^count10Up"];
                                                  row["類2平均組距count90Down"] += "/" + analytics[key + "^^^count90Down"];
                                                  row["類2平均組距count80Down"] += "/" + analytics[key + "^^^count80Down"];
                                                  row["類2平均組距count70Down"] += "/" + analytics[key + "^^^count70Down"];
                                                  row["類2平均組距count60Down"] += "/" + analytics[key + "^^^count60Down"];
                                                  row["類2平均組距count50Down"] += "/" + analytics[key + "^^^count50Down"];
                                                  row["類2平均組距count40Down"] += "/" + analytics[key + "^^^count40Down"];
                                                  row["類2平均組距count30Down"] += "/" + analytics[key + "^^^count30Down"];
                                                  row["類2平均組距count20Down"] += "/" + analytics[key + "^^^count20Down"];
                                                  row["類2平均組距count10Down"] += "/" + analytics[key + "^^^count10Down"];
                                             }
                                        }
                                        #endregion
                                        #region 類別2加權總分組距及高低標
                                        key = "類別2加權總分排名" + "^^^" + gradeYear + "^^^" + tag;
                                        if (rankStudents.ContainsKey(key))
                                        {
                                             if (i == 0)
                                             {
                                                  row["類2加權總分高標"] = analytics[key + "^^^高標"];
                                                  row["類2加權總分均標"] = analytics[key + "^^^均標"];
                                                  row["類2加權總分低標"] = analytics[key + "^^^低標"];
                                                  row["類2加權總分標準差"] = analytics[key + "^^^標準差"];
                                                  row["類2加權總分組距count90"] = analytics[key + "^^^count90"];
                                                  row["類2加權總分組距count80"] = analytics[key + "^^^count80"];
                                                  row["類2加權總分組距count70"] = analytics[key + "^^^count70"];
                                                  row["類2加權總分組距count60"] = analytics[key + "^^^count60"];
                                                  row["類2加權總分組距count50"] = analytics[key + "^^^count50"];
                                                  row["類2加權總分組距count40"] = analytics[key + "^^^count40"];
                                                  row["類2加權總分組距count30"] = analytics[key + "^^^count30"];
                                                  row["類2加權總分組距count20"] = analytics[key + "^^^count20"];
                                                  row["類2加權總分組距count10"] = analytics[key + "^^^count10"];
                                                  row["類2加權總分組距count100Up"] = analytics[key + "^^^count100Up"];
                                                  row["類2加權總分組距count90Up"] = analytics[key + "^^^count90Up"];
                                                  row["類2加權總分組距count80Up"] = analytics[key + "^^^count80Up"];
                                                  row["類2加權總分組距count70Up"] = analytics[key + "^^^count70Up"];
                                                  row["類2加權總分組距count60Up"] = analytics[key + "^^^count60Up"];
                                                  row["類2加權總分組距count50Up"] = analytics[key + "^^^count50Up"];
                                                  row["類2加權總分組距count40Up"] = analytics[key + "^^^count40Up"];
                                                  row["類2加權總分組距count30Up"] = analytics[key + "^^^count30Up"];
                                                  row["類2加權總分組距count20Up"] = analytics[key + "^^^count20Up"];
                                                  row["類2加權總分組距count10Up"] = analytics[key + "^^^count10Up"];
                                                  row["類2加權總分組距count90Down"] = analytics[key + "^^^count90Down"];
                                                  row["類2加權總分組距count80Down"] = analytics[key + "^^^count80Down"];
                                                  row["類2加權總分組距count70Down"] = analytics[key + "^^^count70Down"];
                                                  row["類2加權總分組距count60Down"] = analytics[key + "^^^count60Down"];
                                                  row["類2加權總分組距count50Down"] = analytics[key + "^^^count50Down"];
                                                  row["類2加權總分組距count40Down"] = analytics[key + "^^^count40Down"];
                                                  row["類2加權總分組距count30Down"] = analytics[key + "^^^count30Down"];
                                                  row["類2加權總分組距count20Down"] = analytics[key + "^^^count20Down"];
                                                  row["類2加權總分組距count10Down"] = analytics[key + "^^^count10Down"];
                                             }
                                             else
                                             {
                                                  row["類2加權總分高標"] += "/" + analytics[key + "^^^高標"];
                                                  row["類2加權總分均標"] += "/" + analytics[key + "^^^均標"];
                                                  row["類2加權總分低標"] += "/" + analytics[key + "^^^低標"];
                                                  row["類2加權總分標準差"] += "/" + analytics[key + "^^^標準差"];
                                                  row["類2加權總分組距count90"] += "/" + analytics[key + "^^^count90"];
                                                  row["類2加權總分組距count80"] += "/" + analytics[key + "^^^count80"];
                                                  row["類2加權總分組距count70"] += "/" + analytics[key + "^^^count70"];
                                                  row["類2加權總分組距count60"] += "/" + analytics[key + "^^^count60"];
                                                  row["類2加權總分組距count50"] += "/" + analytics[key + "^^^count50"];
                                                  row["類2加權總分組距count40"] += "/" + analytics[key + "^^^count40"];
                                                  row["類2加權總分組距count30"] += "/" + analytics[key + "^^^count30"];
                                                  row["類2加權總分組距count20"] += "/" + analytics[key + "^^^count20"];
                                                  row["類2加權總分組距count10"] += "/" + analytics[key + "^^^count10"];
                                                  row["類2加權總分組距count100Up"] += "/" + analytics[key + "^^^count100Up"];
                                                  row["類2加權總分組距count90Up"] += "/" + analytics[key + "^^^count90Up"];
                                                  row["類2加權總分組距count80Up"] += "/" + analytics[key + "^^^count80Up"];
                                                  row["類2加權總分組距count70Up"] += "/" + analytics[key + "^^^count70Up"];
                                                  row["類2加權總分組距count60Up"] += "/" + analytics[key + "^^^count60Up"];
                                                  row["類2加權總分組距count50Up"] += "/" + analytics[key + "^^^count50Up"];
                                                  row["類2加權總分組距count40Up"] += "/" + analytics[key + "^^^count40Up"];
                                                  row["類2加權總分組距count30Up"] += "/" + analytics[key + "^^^count30Up"];
                                                  row["類2加權總分組距count20Up"] += "/" + analytics[key + "^^^count20Up"];
                                                  row["類2加權總分組距count10Up"] += "/" + analytics[key + "^^^count10Up"];
                                                  row["類2加權總分組距count90Down"] += "/" + analytics[key + "^^^count90Down"];
                                                  row["類2加權總分組距count80Down"] += "/" + analytics[key + "^^^count80Down"];
                                                  row["類2加權總分組距count70Down"] += "/" + analytics[key + "^^^count70Down"];
                                                  row["類2加權總分組距count60Down"] += "/" + analytics[key + "^^^count60Down"];
                                                  row["類2加權總分組距count50Down"] += "/" + analytics[key + "^^^count50Down"];
                                                  row["類2加權總分組距count40Down"] += "/" + analytics[key + "^^^count40Down"];
                                                  row["類2加權總分組距count30Down"] += "/" + analytics[key + "^^^count30Down"];
                                                  row["類2加權總分組距count20Down"] += "/" + analytics[key + "^^^count20Down"];
                                                  row["類2加權總分組距count10Down"] += "/" + analytics[key + "^^^count10Down"];
                                             }
                                        }
                                        #endregion
                                        #region 類別2加權平均組距及高低標
                                        key = "類別2加權平均排名" + "^^^" + gradeYear + "^^^" + tag;
                                        if (rankStudents.ContainsKey(key))
                                        {
                                             if (i == 0)
                                             {
                                                  row["類2加權平均高標"] = analytics[key + "^^^高標"];
                                                  row["類2加權平均均標"] = analytics[key + "^^^均標"];
                                                  row["類2加權平均低標"] = analytics[key + "^^^低標"];
                                                  row["類2加權平均標準差"] = analytics[key + "^^^標準差"];
                                                  row["類2加權平均組距count90"] = analytics[key + "^^^count90"];
                                                  row["類2加權平均組距count80"] = analytics[key + "^^^count80"];
                                                  row["類2加權平均組距count70"] = analytics[key + "^^^count70"];
                                                  row["類2加權平均組距count60"] = analytics[key + "^^^count60"];
                                                  row["類2加權平均組距count50"] = analytics[key + "^^^count50"];
                                                  row["類2加權平均組距count40"] = analytics[key + "^^^count40"];
                                                  row["類2加權平均組距count30"] = analytics[key + "^^^count30"];
                                                  row["類2加權平均組距count20"] = analytics[key + "^^^count20"];
                                                  row["類2加權平均組距count10"] = analytics[key + "^^^count10"];
                                                  row["類2加權平均組距count100Up"] = analytics[key + "^^^count100Up"];
                                                  row["類2加權平均組距count90Up"] = analytics[key + "^^^count90Up"];
                                                  row["類2加權平均組距count80Up"] = analytics[key + "^^^count80Up"];
                                                  row["類2加權平均組距count70Up"] = analytics[key + "^^^count70Up"];
                                                  row["類2加權平均組距count60Up"] = analytics[key + "^^^count60Up"];
                                                  row["類2加權平均組距count50Up"] = analytics[key + "^^^count50Up"];
                                                  row["類2加權平均組距count40Up"] = analytics[key + "^^^count40Up"];
                                                  row["類2加權平均組距count30Up"] = analytics[key + "^^^count30Up"];
                                                  row["類2加權平均組距count20Up"] = analytics[key + "^^^count20Up"];
                                                  row["類2加權平均組距count10Up"] = analytics[key + "^^^count10Up"];
                                                  row["類2加權平均組距count90Down"] = analytics[key + "^^^count90Down"];
                                                  row["類2加權平均組距count80Down"] = analytics[key + "^^^count80Down"];
                                                  row["類2加權平均組距count70Down"] = analytics[key + "^^^count70Down"];
                                                  row["類2加權平均組距count60Down"] = analytics[key + "^^^count60Down"];
                                                  row["類2加權平均組距count50Down"] = analytics[key + "^^^count50Down"];
                                                  row["類2加權平均組距count40Down"] = analytics[key + "^^^count40Down"];
                                                  row["類2加權平均組距count30Down"] = analytics[key + "^^^count30Down"];
                                                  row["類2加權平均組距count20Down"] = analytics[key + "^^^count20Down"];
                                                  row["類2加權平均組距count10Down"] = analytics[key + "^^^count10Down"];
                                             }
                                             else
                                             {
                                                  row["類2加權平均高標"] += "/" + analytics[key + "^^^高標"];
                                                  row["類2加權平均均標"] += "/" + analytics[key + "^^^均標"];
                                                  row["類2加權平均低標"] += "/" + analytics[key + "^^^低標"];
                                                  row["類2加權平均標準差"] += "/" + analytics[key + "^^^標準差"];
                                                  row["類2加權平均組距count90"] += "/" + analytics[key + "^^^count90"];
                                                  row["類2加權平均組距count80"] += "/" + analytics[key + "^^^count80"];
                                                  row["類2加權平均組距count70"] += "/" + analytics[key + "^^^count70"];
                                                  row["類2加權平均組距count60"] += "/" + analytics[key + "^^^count60"];
                                                  row["類2加權平均組距count50"] += "/" + analytics[key + "^^^count50"];
                                                  row["類2加權平均組距count40"] += "/" + analytics[key + "^^^count40"];
                                                  row["類2加權平均組距count30"] += "/" + analytics[key + "^^^count30"];
                                                  row["類2加權平均組距count20"] += "/" + analytics[key + "^^^count20"];
                                                  row["類2加權平均組距count10"] += "/" + analytics[key + "^^^count10"];
                                                  row["類2加權平均組距count100Up"] += "/" + analytics[key + "^^^count100Up"];
                                                  row["類2加權平均組距count90Up"] += "/" + analytics[key + "^^^count90Up"];
                                                  row["類2加權平均組距count80Up"] += "/" + analytics[key + "^^^count80Up"];
                                                  row["類2加權平均組距count70Up"] += "/" + analytics[key + "^^^count70Up"];
                                                  row["類2加權平均組距count60Up"] += "/" + analytics[key + "^^^count60Up"];
                                                  row["類2加權平均組距count50Up"] += "/" + analytics[key + "^^^count50Up"];
                                                  row["類2加權平均組距count40Up"] += "/" + analytics[key + "^^^count40Up"];
                                                  row["類2加權平均組距count30Up"] += "/" + analytics[key + "^^^count30Up"];
                                                  row["類2加權平均組距count20Up"] += "/" + analytics[key + "^^^count20Up"];
                                                  row["類2加權平均組距count10Up"] += "/" + analytics[key + "^^^count10Up"];
                                                  row["類2加權平均組距count90Down"] += "/" + analytics[key + "^^^count90Down"];
                                                  row["類2加權平均組距count80Down"] += "/" + analytics[key + "^^^count80Down"];
                                                  row["類2加權平均組距count70Down"] += "/" + analytics[key + "^^^count70Down"];
                                                  row["類2加權平均組距count60Down"] += "/" + analytics[key + "^^^count60Down"];
                                                  row["類2加權平均組距count50Down"] += "/" + analytics[key + "^^^count50Down"];
                                                  row["類2加權平均組距count40Down"] += "/" + analytics[key + "^^^count40Down"];
                                                  row["類2加權平均組距count30Down"] += "/" + analytics[key + "^^^count30Down"];
                                                  row["類2加權平均組距count20Down"] += "/" + analytics[key + "^^^count20Down"];
                                                  row["類2加權平均組距count10Down"] += "/" + analytics[key + "^^^count10Down"];
                                             }
                                        }
                                        #endregion
                                   }
                                   #endregion
                                   table.Rows.Add(row);
                                   progressCount++;
                                   bkw.ReportProgress(70 + progressCount * 20 / selectedClasses.Count);

                              }
                              #endregion
                              bkw.ReportProgress(90);
                              document = conf.Template;
                              document.MailMerge.Execute(table);
                         }
                         catch (Exception exception)
                         {
                              exc = exception;
                         }
                    };
                    bkw.RunWorkerAsync();
               }
          }

          internal static void CreateFieldTemplate()
          {
               #region 產生欄位表

               Aspose.Words.Document doc = new Aspose.Words.Document(new System.IO.MemoryStream(Properties.Resources.Template));
               Aspose.Words.DocumentBuilder builder = new Aspose.Words.DocumentBuilder(doc);
               int maxSubjectNum = 15;
               int maxStuNum = 60;

               builder.Font.Size = 8;
               #region 基本欄位
               builder.Writeln("基本欄位");
               builder.StartTable();
               foreach (string field in new string[] { "學年度", "學期", "學校名稱", "學校地址", "學校電話", "科別名稱", "定期評量", "班級", "班導師", "類別排名1", "類別排名2" })
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
                         builder.InsertField("MERGEFIELD 科目成績" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«S" + i + "»");
                    }
                    builder.EndRow();
               }
               builder.EndTable();
               #endregion

               #region 科目成績排名
               foreach (string key in new string[] { "班", "科", "全校", "類別1", "類別2" })
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
               builder.Writeln("前次成績");
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
                         builder.InsertField("MERGEFIELD 前次成績" + stuIndex + "-" + i + " \\* MERGEFORMAT ", "«S" + i + "»");
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
               builder.Write("總分校排名");
               builder.InsertCell();
               builder.Write("平均");
               builder.InsertCell();
               builder.Write("平均班排名");
               builder.InsertCell();
               builder.Write("平均科排名");
               builder.InsertCell();
               builder.Write("平均校排名");
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
                    builder.InsertField("MERGEFIELD 總分全校排名" + stuIndex + " \\* MERGEFORMAT ", "«RS»");
                    builder.InsertField("MERGEFIELD 總分全校排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 平均" + stuIndex + " \\* MERGEFORMAT ", "«平均»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 平均班排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                    builder.InsertField("MERGEFIELD 平均班排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 平均科排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                    builder.InsertField("MERGEFIELD 平均科排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 平均全校排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                    builder.InsertField("MERGEFIELD 平均全校排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
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
               builder.Write("加權總分校排名");
               builder.InsertCell();
               builder.Write("加權平均");
               builder.InsertCell();
               builder.Write("加權平均班排名");
               builder.InsertCell();
               builder.Write("加權平均科排名");
               builder.InsertCell();
               builder.Write("加權平均校排名");
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
                    builder.InsertField("MERGEFIELD 加權總分全校排名" + stuIndex + " \\* MERGEFORMAT ", "«RS»");
                    builder.InsertField("MERGEFIELD 加權總分全校排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TS»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權平均" + stuIndex + " \\* MERGEFORMAT ", "«平均»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權平均班排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                    builder.InsertField("MERGEFIELD 加權平均班排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權平均科排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                    builder.InsertField("MERGEFIELD 加權平均科排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
                    builder.InsertCell();
                    builder.InsertField("MERGEFIELD 加權平均全校排名" + stuIndex + " \\* MERGEFORMAT ", "«RA»");
                    builder.InsertField("MERGEFIELD 加權平均全校排名母數" + stuIndex + " \\b /  \\* MERGEFORMAT ", "/«TA»");
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
               builder.Write("總分");
               builder.InsertCell();
               builder.Write("總分排名");
               builder.InsertCell();
               builder.Write("平均");
               builder.InsertCell();
               builder.Write("平均排名");
               builder.InsertCell();
               builder.Write("加權總分");
               builder.InsertCell();
               builder.Write("加權總分排名");
               builder.InsertCell();
               builder.Write("加權平均");
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
               builder.Write("總分");
               builder.InsertCell();
               builder.Write("總分排名");
               builder.InsertCell();
               builder.Write("平均");
               builder.InsertCell();
               builder.Write("平均排名");
               builder.InsertCell();
               builder.Write("加權總分");
               builder.InsertCell();
               builder.Write("加權總分排名");
               builder.InsertCell();
               builder.Write("加權平均");
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

               #region 各項科目成績分析
               foreach (string key in new string[] { "班", "科", "校", "類1", "類2" })
               {
                    builder.InsertBreak(Aspose.Words.BreakType.PageBreak);

                    builder.Writeln(key + "成績分析及組距");

                    builder.StartTable();
                    builder.InsertCell(); builder.Write("科目名稱");
                    for (int subjectIndex = 1; subjectIndex <= maxSubjectNum; subjectIndex++)
                    {
                         builder.InsertCell(); builder.InsertField("MERGEFIELD 科目名稱" + subjectIndex + " \\* MERGEFORMAT ", "«N" + subjectIndex + "»");
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
               builder.InsertBreak(Aspose.Words.BreakType.PageBreak);
               builder.Writeln("加總成績分析及組距");
               builder.StartTable();
               builder.InsertCell(); builder.Write("項目");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.Write(key);
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("高標");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "高標 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("均標");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "均標 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("低標");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "低標 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("標準差");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "標準差 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("100以上");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count100Up \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("90以上");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count90Up \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("90以上小於100");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count90 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("小於90");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count90Down \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("80以上");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count80Up \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("80以上小於90");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count80 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("小於80");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count80Down \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("70以上");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count70Up \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("70以上小於80");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count70 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("小於70");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count70Down \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("60以上");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count60Up \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("60以上小於70");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count60 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("小於60");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count60Down \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("50以上");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count50Up \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("50以上小於60");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count50 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("小於50");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count50Down \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("40以上");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count40Up \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("40以上小於50");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count40 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("小於40");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count40Down \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("30以上");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count30Up \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("30以上小於40");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count30 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("小於30");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count30Down \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("20以上");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count20Up \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("20以上小於30");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count20 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("小於20");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count20Down \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("10以上");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count10Up \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("10以上小於20");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count10 \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.InsertCell(); builder.Write("小於10");
               foreach (string key in new string[] { "總分班", "總分科", "總分校", "平均班", "平均科", "平均校", "加權總分班", "加權總分科", "加權總分校", "加權平均班", "加權平均科", "加權平均校", "類1總分", "類1平均", "類1加權總分", "類1加權平均", "類2總分", "類2平均", "類2加權總分", "類2加權平均" })
               {
                    builder.InsertCell(); builder.InsertField("MERGEFIELD " + key + "組距count10Down \\* MERGEFORMAT ", "«C»");
               }
               builder.EndRow();
               builder.EndTable();
               #endregion
               #endregion

               #region 儲存檔案
               string inputReportName = "班級評量成績單合併欄位總表";
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
                              FISCA.Presentation.Controls.MsgBox.Show("指定路徑無法存取。", "建立檔案失敗", System.Windows.Forms.MessageBoxButtons.OK, System.Windows.Forms.MessageBoxIcon.Error);
                              return;
                         }
                    }
               }
               #endregion
          }
     }
}


