﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Data;
using System.IO;
using SmartSchool.Customization.Data;
using System.Threading;
using SmartSchool.Customization.Data.StudentExtension;
using FISCA.Permission;
using SmartSchool;
using Campus.ePaperCloud;
using System.Xml.Linq;
using System.Xml.XPath;
using System.Xml;
using SH_SemesterScoreReportFixed.DAO;
using SHStudentExamExtension;

namespace SH_SemesterScoreReportFixed
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            var btn = K12.Presentation.NLDPanels.Student.RibbonBarItems["資料統計"]["報表"]["成績相關報表"]["期末成績通知單(固定排名)"];
            btn.Enable = false;
            K12.Presentation.NLDPanels.Student.SelectedSourceChanged += delegate
            {
                btn.Enable = Permissions.期末成績通知單_固定排名權限 && (K12.Presentation.NLDPanels.Student.SelectedSource.Count > 0);
            };
            btn.Click += new EventHandler(Program_Click);

            //權限設定
            Catalog permission = RoleAclSource.Instance["學生"]["功能按鈕"];
            permission.Add(new RibbonFeature(Permissions.期末成績通知單_固定排名, "期末成績通知單(固定排名)"));
        }

        // 四捨五入位數
        private static int r2ParseN = 2;
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

        static Dictionary<string, decimal> _studPassSumCreditDict1 = new Dictionary<string, decimal>();
        static Dictionary<string, decimal> _studPassSumCreditDictAll = new Dictionary<string, decimal>();

        // 累計取得必修學分
        static Dictionary<string, decimal> _studPassSumCreditDictC1 = new Dictionary<string, decimal>();
        // 累計取得選修學分
        static Dictionary<string, decimal> _studPassSumCreditDictC2 = new Dictionary<string, decimal>();

        // 本學期已修必修學分
        static Dictionary<string, decimal> StudentSemsReqSumCreditsDict = new Dictionary<string, decimal>();

        // 本學期已修選修學分
        static Dictionary<string, decimal> StudentSemsSelSumCreditsDict = new Dictionary<string, decimal>();

        // 本學期實得必修學分
        static Dictionary<string, decimal> StudentSemsPassReqSumCreditsDict = new Dictionary<string, decimal>();

        // 本學期實得選修學分
        static Dictionary<string, decimal> StudentSemsPassSelSumCreditsDict = new Dictionary<string, decimal>();

        // 在校期間實得必修學分
        static Dictionary<string, decimal> StudentAllPassReqSumCreditsDict = new Dictionary<string, decimal>();

        // 在校期間實得選修學分
        static Dictionary<string, decimal> StudentAllPassSelSumCreditsDict = new Dictionary<string, decimal>();




        static void Program_Click(object sender_, EventArgs e_)
        {
            AccessHelper helper = new AccessHelper();
            List<StudentRecord> lista = helper.StudentHelper.GetSelectedStudent();

            List<string> StudentIDList = new List<string>();
            //foreach (StudentRecord stud in lista)
            //    StudentIDList.Add(stud.StudentID);
            //// 取得學生修課及格標準
            //Dictionary<string, Dictionary<string, decimal>> StudentSCAttendApplyLimitDict = Utility.GetStudentSCAttendApplyLimitDict(StudentIDList);

            // 取得學生及格與補考標準
            Dictionary<string, Dictionary<string, decimal>> StudentApplyLimitDict = Utility.GetStudentApplyLimitDict(lista);


            ConfigForm form = new ConfigForm();
            if (form.ShowDialog() == DialogResult.OK)
            {
                AccessHelper accessHelper = new AccessHelper();


                //return;
                List<StudentRecord> overflowRecords = new List<StudentRecord>();
                //取得列印設定
                Configure conf = form.Configure;
                //建立測試的選取學生(先期不管怎麼選就是印這些人)
                List<string> selectedStudents = K12.Presentation.NLDPanels.Student.SelectedSource;

                // 取得評量成績排名、五標、分數區間
                Dictionary<string, Dictionary<string, DataRow>> RankMatrixDataDict = Utility.GetRankMatrixData(conf.SchoolYear, conf.Semester, conf.ExamRecord.ID, selectedStudents);

                // 取得學期成績排名、五標、分數區間
                Dictionary<string, Dictionary<string, DataRow>> SemsScoreRankMatrixDataDict = Utility.GetSemsScoreRankMatrixData(conf.SchoolYear, conf.Semester, selectedStudents);

                // 缺曠對照表key值(ex : 本學期一般_曠課、上學期集會_事假
                List<string> AttendanceMappingKeyList = new List<string>();

                //建立合併欄位總表
                DataTable table = new DataTable();
                #region 所有的合併欄位
                table.Columns.Add("學校名稱");
                table.Columns.Add("學校地址");
                table.Columns.Add("學校電話");
                table.Columns.Add("校長名稱");
                table.Columns.Add("教務主任");
                table.Columns.Add("學務主任");
                table.Columns.Add("收件人地址");
                //«通訊地址»«通訊地址郵遞區號»«通訊地址內容»
                //«戶籍地址»«戶籍地址郵遞區號»«戶籍地址內容»
                //«監護人»«父親»«母親»«科別名稱»
                table.Columns.Add("通訊地址");
                table.Columns.Add("通訊地址郵遞區號");
                table.Columns.Add("通訊地址內容");
                table.Columns.Add("戶籍地址");
                table.Columns.Add("戶籍地址郵遞區號");
                table.Columns.Add("戶籍地址內容");
                table.Columns.Add("監護人");
                table.Columns.Add("父親");
                table.Columns.Add("母親");
                table.Columns.Add("科別名稱");
                table.Columns.Add("試別");

                table.Columns.Add("系統編號");
                table.Columns.Add("收件人");
                table.Columns.Add("學年度");
                table.Columns.Add("學期");
                table.Columns.Add("班級科別名稱");
                table.Columns.Add("班級");
                table.Columns.Add("班導師");
                table.Columns.Add("座號");
                table.Columns.Add("學號");
                table.Columns.Add("姓名");
                table.Columns.Add("英文姓名");
                table.Columns.Add("國籍一");
                table.Columns.Add("國籍一護照名");
                table.Columns.Add("國籍二");
                table.Columns.Add("國籍二護照名");
                table.Columns.Add("國籍一英文");
                table.Columns.Add("國籍二英文");
                table.Columns.Add("定期評量");
                table.Columns.Add("本學期取得學分數");
                table.Columns.Add("累計取得學分數");
                table.Columns.Add("累計取得必修學分");
                table.Columns.Add("累計取得選修學分");
                table.Columns.Add("系統學年度");
                table.Columns.Add("系統學期");

                table.Columns.Add("本學期已修必修學分");
                table.Columns.Add("本學期已修選修學分");
                table.Columns.Add("本學期實得必修學分");
                table.Columns.Add("本學期實得選修學分");
                table.Columns.Add("在校期間實得必修學分");
                table.Columns.Add("在校期間實得選修學分");
                table.Columns.Add("學期成績排名採計成績欄位");
                table.Columns.Add("升級或應重讀");
                // 新增合併欄位
                List<string> r1List = new List<string>();
                List<string> r2List = new List<string>();
                List<string> r3List = new List<string>();

                List<string> r2ParseList = new List<string>();

                r1List.Add("總分班排名");
                r1List.Add("總分科排名");
                r1List.Add("總分全校排名");
                r1List.Add("平均班排名");
                r1List.Add("平均科排名");
                r1List.Add("平均全校排名");
                r1List.Add("加權總分班排名");
                r1List.Add("加權總分科排名");
                r1List.Add("加權總分全校排名");
                r1List.Add("加權平均班排名");
                r1List.Add("加權平均科排名");
                r1List.Add("加權平均全校排名");
                r1List.Add("類別1總分排名");
                r1List.Add("類別1平均排名");
                r1List.Add("類別1加權總分排名");
                r1List.Add("類別1加權平均排名");
                r1List.Add("類別2總分排名");
                r1List.Add("類別2平均排名");
                r1List.Add("類別2加權總分排名");
                r1List.Add("類別2加權平均排名");

                r2List.Add("pr");
                r2List.Add("percentile");
                r2List.Add("avg_top_25");
                r2List.Add("avg_top_50");
                r2List.Add("avg");
                r2List.Add("avg_bottom_50");
                r2List.Add("avg_bottom_25");
                r2List.Add("pr_88");
                r2List.Add("pr_75");
                r2List.Add("pr_50");
                r2List.Add("pr_25");
                r2List.Add("pr_12");
                r2List.Add("std_dev_pop");
                r2List.Add("level_gte100");
                r2List.Add("level_90");
                r2List.Add("level_80");
                r2List.Add("level_70");
                r2List.Add("level_60");
                r2List.Add("level_50");
                r2List.Add("level_40");
                r2List.Add("level_30");
                r2List.Add("level_20");
                r2List.Add("level_10");
                r2List.Add("level_lt10");

                //// 四捨五入位數
                //int r2ParseN = 2;

                r2ParseList.Add("avg_top_25");
                r2ParseList.Add("avg_top_50");
                r2ParseList.Add("avg");
                r2ParseList.Add("avg_bottom_50");
                r2ParseList.Add("avg_bottom_25");
                r2ParseList.Add("pr_88");
                r2ParseList.Add("pr_75");
                r2ParseList.Add("pr_50");
                r2ParseList.Add("pr_25");
                r2ParseList.Add("pr_12");
                r2ParseList.Add("std_dev_pop");


                r3List.Add("班排名");
                r3List.Add("科排名");
                r3List.Add("全校排名");
                r3List.Add("類別1排名");
                r3List.Add("類別2排名");

                foreach (string r1 in r1List)
                {
                    foreach (string r2 in r2List)
                    {
                        string it = r1 + "_" + r2;
                        table.Columns.Add(it);
                    }
                }

                for (int subjectIndex = 1; subjectIndex <= conf.SubjectLimit; subjectIndex++)
                {
                    table.Columns.Add("領域" + subjectIndex);
                    table.Columns.Add("分項類別" + subjectIndex);
                    table.Columns.Add("科目名稱" + subjectIndex);
                    table.Columns.Add("科目" + subjectIndex);
                    table.Columns.Add("科目級別" + subjectIndex);
                    table.Columns.Add("學分數" + subjectIndex);
                    table.Columns.Add("前次成績" + subjectIndex);
                    table.Columns.Add("科目成績" + subjectIndex);
                    // 新增學期科目相關成績--
                    table.Columns.Add("科目必選修" + subjectIndex);
                    table.Columns.Add("科目校部定/科目必選修" + subjectIndex);
                    table.Columns.Add("科目註記" + subjectIndex);
                    table.Columns.Add("科目取得學分" + subjectIndex);
                    table.Columns.Add("科目未取得學分註記" + subjectIndex);
                    table.Columns.Add("學期科目原始成績" + subjectIndex);
                    table.Columns.Add("學期科目補考成績" + subjectIndex);
                    table.Columns.Add("學期科目重修成績" + subjectIndex);
                    table.Columns.Add("學期科目手動調整成績" + subjectIndex);
                    table.Columns.Add("學期科目學年調整成績" + subjectIndex);
                    table.Columns.Add("學期科目成績" + subjectIndex);
                    table.Columns.Add("學期科目原始成績註記" + subjectIndex);
                    table.Columns.Add("學期科目補考成績註記" + subjectIndex);
                    table.Columns.Add("學期科目重修成績註記" + subjectIndex);
                    table.Columns.Add("學期科目手動成績註記" + subjectIndex);
                    table.Columns.Add("學期科目學年成績註記" + subjectIndex);
                    table.Columns.Add("學期科目需要補考註記" + subjectIndex);
                    table.Columns.Add("學期科目需要重修註記" + subjectIndex);
                    table.Columns.Add("學期科目需要補考標示" + subjectIndex);
                    table.Columns.Add("學期科目補考成績標示" + subjectIndex);
                    table.Columns.Add("學期科目未取得學分標示" + subjectIndex);
                    table.Columns.Add("學期科目需要重修標示" + subjectIndex);
                    table.Columns.Add("學期科目重修成績標示" + subjectIndex);
                    table.Columns.Add("學期科目補修成績標示" + subjectIndex);
                    table.Columns.Add("科目缺考原因" + subjectIndex);
                    table.Columns.Add("前次科目缺考原因" + subjectIndex);


                    // 新增學期科目排名
                    table.Columns.Add("學期科目排名成績" + subjectIndex);

                    table.Columns.Add("學期科目班排名" + subjectIndex);
                    table.Columns.Add("學期科目班排名母數" + subjectIndex);
                    table.Columns.Add("學期科目科排名" + subjectIndex);
                    table.Columns.Add("學期科目科排名母數" + subjectIndex);
                    table.Columns.Add("學期科目類別1排名" + subjectIndex);
                    table.Columns.Add("學期科目類別1排名母數" + subjectIndex);
                    table.Columns.Add("學期科目類別2排名" + subjectIndex);
                    table.Columns.Add("學期科目類別2排名母數" + subjectIndex);
                    table.Columns.Add("學期科目全校排名" + subjectIndex);
                    table.Columns.Add("學期科目全校排名母數" + subjectIndex);
                    table.Columns.Add("學期科目(原始)班排名" + subjectIndex);
                    table.Columns.Add("學期科目(原始)班排名母數" + subjectIndex);
                    table.Columns.Add("學期科目(原始)科排名" + subjectIndex);
                    table.Columns.Add("學期科目(原始)科排名母數" + subjectIndex);
                    table.Columns.Add("學期科目(原始)類別1排名" + subjectIndex);
                    table.Columns.Add("學期科目(原始)類別1排名母數" + subjectIndex);
                    table.Columns.Add("學期科目(原始)類別2排名" + subjectIndex);
                    table.Columns.Add("學期科目(原始)類別2排名母數" + subjectIndex);
                    table.Columns.Add("學期科目(原始)全校排名" + subjectIndex);
                    table.Columns.Add("學期科目(原始)全校排名母數" + subjectIndex);

                    // 學期科目五標與組距 
                    foreach (string item2 in r2List)
                    {
                        table.Columns.Add("學期科目班排名" + subjectIndex + "_" + item2);
                        table.Columns.Add("學期科目科排名" + subjectIndex + "_" + item2);
                        table.Columns.Add("學期科目類別1排名" + subjectIndex + "_" + item2);
                        table.Columns.Add("學期科目類別2排名" + subjectIndex + "_" + item2);
                        table.Columns.Add("學期科目全校排名" + subjectIndex + "_" + item2);
                        table.Columns.Add("學期科目(原始)班排名" + subjectIndex + "_" + item2);
                        table.Columns.Add("學期科目(原始)科排名" + subjectIndex + "_" + item2);
                        table.Columns.Add("學期科目(原始)類別1排名" + subjectIndex + "_" + item2);
                        table.Columns.Add("學期科目(原始)類別2排名" + subjectIndex + "_" + item2);
                        table.Columns.Add("學期科目(原始)全校排名" + subjectIndex + "_" + item2);

                    }


                    // 新增上學期科目相關成績--
                    table.Columns.Add("上學期科目原始成績" + subjectIndex);
                    table.Columns.Add("上學期科目補考成績" + subjectIndex);
                    table.Columns.Add("上學期科目重修成績" + subjectIndex);
                    table.Columns.Add("上學期科目手動調整成績" + subjectIndex);
                    table.Columns.Add("上學期科目學年調整成績" + subjectIndex);
                    table.Columns.Add("上學期科目成績" + subjectIndex);
                    table.Columns.Add("上學期科目原始成績註記" + subjectIndex);
                    table.Columns.Add("上學期科目補考成績註記" + subjectIndex);
                    table.Columns.Add("上學期科目重修成績註記" + subjectIndex);
                    table.Columns.Add("上學期科目手動成績註記" + subjectIndex);
                    table.Columns.Add("上學期科目學年成績註記" + subjectIndex);
                    table.Columns.Add("上學期科目取得學分" + subjectIndex);
                    table.Columns.Add("上學期科目未取得學分註記" + subjectIndex);
                    table.Columns.Add("上學期科目需要補考註記" + subjectIndex);
                    table.Columns.Add("上學期科目需要重修註記" + subjectIndex);
                    table.Columns.Add("上學期科目需要補考標示" + subjectIndex);
                    table.Columns.Add("上學期科目補考成績標示" + subjectIndex);
                    table.Columns.Add("上學期科目未取得學分標示" + subjectIndex);
                    table.Columns.Add("上學期科目需要重修標示" + subjectIndex);
                    table.Columns.Add("上學期科目重修成績標示" + subjectIndex);
                    table.Columns.Add("上學期科目補修成績標示" + subjectIndex);

                    // 新增學年科目成績--
                    table.Columns.Add("學年科目成績" + subjectIndex);

                    table.Columns.Add("班排名" + subjectIndex);
                    table.Columns.Add("班排名母數" + subjectIndex);
                    table.Columns.Add("科排名" + subjectIndex);
                    table.Columns.Add("科排名母數" + subjectIndex);
                    table.Columns.Add("類別1排名" + subjectIndex);
                    table.Columns.Add("類別1排名母數" + subjectIndex);
                    table.Columns.Add("類別2排名" + subjectIndex);
                    table.Columns.Add("類別2排名母數" + subjectIndex);
                    table.Columns.Add("全校排名" + subjectIndex);
                    table.Columns.Add("全校排名母數" + subjectIndex);

                    foreach (string r3 in r3List)
                    {
                        foreach (string r2 in r2List)
                        {
                            string it = r3 + subjectIndex + "_" + r2;
                            table.Columns.Add(it);
                        }
                    }
                }



                table.Columns.Add("總分");
                table.Columns.Add("總分班排名");
                table.Columns.Add("總分班排名母數");
                table.Columns.Add("總分科排名");
                table.Columns.Add("總分科排名母數");
                table.Columns.Add("總分全校排名");
                table.Columns.Add("總分全校排名母數");
                table.Columns.Add("平均");
                table.Columns.Add("平均班排名");
                table.Columns.Add("平均班排名母數");
                table.Columns.Add("平均科排名");
                table.Columns.Add("平均科排名母數");
                table.Columns.Add("平均全校排名");
                table.Columns.Add("平均全校排名母數");

                table.Columns.Add("加權總分");
                table.Columns.Add("加權總分班排名");
                table.Columns.Add("加權總分班排名母數");
                table.Columns.Add("加權總分科排名");
                table.Columns.Add("加權總分科排名母數");
                table.Columns.Add("加權總分全校排名");
                table.Columns.Add("加權總分全校排名母數");
                table.Columns.Add("加權平均");
                table.Columns.Add("加權平均班排名");
                table.Columns.Add("加權平均班排名母數");
                table.Columns.Add("加權平均科排名");
                table.Columns.Add("加權平均科排名母數");
                table.Columns.Add("加權平均全校排名");
                table.Columns.Add("加權平均全校排名母數");

                table.Columns.Add("類別排名1");
                table.Columns.Add("學期類別排名1");
                table.Columns.Add("類別1總分");
                table.Columns.Add("類別1總分排名");
                table.Columns.Add("類別1總分排名母數");
                table.Columns.Add("類別1平均");
                table.Columns.Add("類別1平均排名");
                table.Columns.Add("類別1平均排名母數");
                table.Columns.Add("類別1加權總分");
                table.Columns.Add("類別1加權總分排名");
                table.Columns.Add("類別1加權總分排名母數");
                table.Columns.Add("類別1加權平均");
                table.Columns.Add("類別1加權平均排名");
                table.Columns.Add("類別1加權平均排名母數");

                table.Columns.Add("類別排名2");
                table.Columns.Add("學期類別排名2");
                table.Columns.Add("類別2總分");
                table.Columns.Add("類別2總分排名");
                table.Columns.Add("類別2總分排名母數");
                table.Columns.Add("類別2平均");
                table.Columns.Add("類別2平均排名");
                table.Columns.Add("類別2平均排名母數");
                table.Columns.Add("類別2加權總分");
                table.Columns.Add("類別2加權總分排名");
                table.Columns.Add("類別2加權總分排名母數");
                table.Columns.Add("類別2加權平均");
                table.Columns.Add("類別2加權平均排名");
                table.Columns.Add("類別2加權平均排名母數");


                // 學期分項成績 --
                table.Columns.Add("學期學業成績");
                table.Columns.Add("學期體育成績");
                table.Columns.Add("學期國防通識成績");
                table.Columns.Add("學期健康與護理成績");
                table.Columns.Add("學期實習科目成績");
                table.Columns.Add("學期學業(原始)成績");
                table.Columns.Add("學期體育(原始)成績");
                table.Columns.Add("學期國防通識(原始)成績");
                table.Columns.Add("學期健康與護理(原始)成績");
                table.Columns.Add("學期實習科目(原始)成績");
                table.Columns.Add("學期專業科目成績");
                table.Columns.Add("學期專業科目(原始)成績");
                table.Columns.Add("學期德行成績");

                // 導師評語 --
                table.Columns.Add("導師評語");
                table.Columns.Add("上學期導師評語");
                // 獎懲統計 --
                table.Columns.Add("大功統計");
                table.Columns.Add("小功統計");
                table.Columns.Add("嘉獎統計");
                table.Columns.Add("大過統計");
                table.Columns.Add("小過統計");
                table.Columns.Add("警告統計");
                table.Columns.Add("留校察看");

                // 上學期獎懲統計 -- //2021-12-27 新增
                table.Columns.Add("上學期大功統計");
                table.Columns.Add("上學期小功統計");
                table.Columns.Add("上學期嘉獎統計");
                table.Columns.Add("上學期大過統計");
                table.Columns.Add("上學期小過統計");
                table.Columns.Add("上學期警告統計");
                table.Columns.Add("上學期留校察看");

                // 上學期分項成績 --
                table.Columns.Add("上學期學業成績");
                table.Columns.Add("上學期體育成績");
                table.Columns.Add("上學期國防通識成績");
                table.Columns.Add("上學期健康與護理成績");
                table.Columns.Add("上學期專業科目成績");
                table.Columns.Add("上學期實習科目成績");
                table.Columns.Add("上學期德行成績");

                // 上學期分項成績 -- //2017/8/28 穎驊補齊缺漏欄位
                table.Columns.Add("上學期學業(原始)成績");
                table.Columns.Add("上學期體育(原始)成績");
                table.Columns.Add("上學期國防通識(原始)成績");
                table.Columns.Add("上學期健康與護理(原始)成績");
                table.Columns.Add("上學期專業科目(原始)成績");
                table.Columns.Add("上學期實習科目(原始)成績");
                table.Columns.Add("上學期德行(原始)成績");


                // 學年分項成績 --
                table.Columns.Add("學年學業成績");
                table.Columns.Add("學年體育成績");
                table.Columns.Add("學年國防通識成績");
                table.Columns.Add("學年健康與護理成績");
                table.Columns.Add("學年實習科目成績");
                table.Columns.Add("學年專業科目成績");
                table.Columns.Add("學年德行成績");
                table.Columns.Add("學年學業成績班排名");

                // 服務學習時數
                table.Columns.Add("前學期服務學習時數");
                table.Columns.Add("本學期服務學習時數");
                table.Columns.Add("學年服務學習時數");

                // 缺曠統計
                // 動態新增缺曠統計，使用模式一般_曠課、一般_事假..
                foreach (string name in Utility.GetATMappingKey())
                {
                    table.Columns.Add("前學期" + name);
                    table.Columns.Add("本學期" + name);
                    table.Columns.Add("學年" + name);

                    AttendanceMappingKeyList.Add(name);
                }

                //  動態產生學期科目與學期分項合併欄位
                List<string> SemsItemNameList = new List<string>();

                SemsItemNameList.Add("學業");
                SemsItemNameList.Add("專業科目");
                SemsItemNameList.Add("實習科目");

                foreach (string name in SemsItemNameList)
                {
                    table.Columns.Add("學期" + name + "成績班排名");
                    table.Columns.Add("學期" + name + "成績科排名");
                    table.Columns.Add("學期" + name + "成績全校排名");
                    table.Columns.Add("學期" + name + "成績類別1排名");
                    table.Columns.Add("學期" + name + "成績類別2排名");
                    table.Columns.Add("學期" + name + "成績班排名母數");
                    table.Columns.Add("學期" + name + "成績科排名母數");
                    table.Columns.Add("學期" + name + "成績全校排名母數");
                    table.Columns.Add("學期" + name + "成績類別1排名母數");
                    table.Columns.Add("學期" + name + "成績類別2排名母數");
                    table.Columns.Add("學期" + name + "(原始)成績班排名");
                    table.Columns.Add("學期" + name + "(原始)成績科排名");
                    table.Columns.Add("學期" + name + "(原始)成績全校排名");
                    table.Columns.Add("學期" + name + "(原始)成績類別1排名");
                    table.Columns.Add("學期" + name + "(原始)成績類別2排名");
                    table.Columns.Add("學期" + name + "(原始)成績班排名母數");
                    table.Columns.Add("學期" + name + "(原始)成績科排名母數");
                    table.Columns.Add("學期" + name + "(原始)成績全校排名母數");
                    table.Columns.Add("學期" + name + "(原始)成績類別1排名母數");
                    table.Columns.Add("學期" + name + "(原始)成績類別2排名母數");

                    // 五標與組距
                    foreach (string item2 in r2List)
                    {
                        table.Columns.Add("學期" + name + "成績班排名" + "_" + item2);
                        table.Columns.Add("學期" + name + "成績科排名" + "_" + item2);
                        table.Columns.Add("學期" + name + "成績全校排名" + "_" + item2);
                        table.Columns.Add("學期" + name + "成績類別1排名" + "_" + item2);
                        table.Columns.Add("學期" + name + "成績類別2排名" + "_" + item2);
                        table.Columns.Add("學期" + name + "(原始)成績班排名" + "_" + item2);
                        table.Columns.Add("學期" + name + "(原始)成績科排名" + "_" + item2);
                        table.Columns.Add("學期" + name + "(原始)成績全校排名" + "_" + item2);
                        table.Columns.Add("學期" + name + "(原始)成績類別1排名" + "_" + item2);
                        table.Columns.Add("學期" + name + "(原始)成績類別2排名" + "_" + item2);
                    }
                }

                // 新增人數統計
                table.Columns.Add("班級人數");
                table.Columns.Add("科別人數");
                table.Columns.Add("年級人數");

                #endregion
                //宣告產生的報表
                Aspose.Words.Document document = new Aspose.Words.Document();
                //用一個BackgroundWorker包起來
                System.ComponentModel.BackgroundWorker bkw = new System.ComponentModel.BackgroundWorker();
                bkw.WorkerReportsProgress = true;
                System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 期末成績單產生 S");
                bkw.ProgressChanged += delegate (object sender, System.ComponentModel.ProgressChangedEventArgs e)
                {
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("期末成績單產生中", e.ProgressPercentage);
                    System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 期末成績單產生 " + e.ProgressPercentage);
                };
                Exception exc = null;
                bkw.RunWorkerCompleted += delegate
                {
                    //#region 將 DataTable 內合併欄位產生出來
                    //StreamWriter sw = new StreamWriter(Application.StartupPath + "\\期末學期成績單合併欄位.txt");
                    //foreach (DataColumn dc in table.Columns)
                    //    sw.WriteLine(dc.Caption);

                    //sw.Close();
                    //#endregion


                    System.Diagnostics.Trace.WriteLine(DateTime.Now.ToString("HH:mm:ss") + " 期末成績單產生 E");
                    string err = "下列學生因成績項目超過樣板支援上限，\n超出部分科目成績無法印出，建議調整樣板內容。";
                    if (overflowRecords.Count > 0)
                    {
                        foreach (var stuRec in overflowRecords)
                        {
                            err += "\n" + (stuRec.RefClass == null ? "" : (stuRec.RefClass.ClassName + "班" + stuRec.SeatNo + "號")) + "[" + stuRec.StudentNumber + "]" + stuRec.StudentName;
                        }
                    }
                    #region 儲存檔案
                    string inputReportName = conf.SchoolYear + "學年度第" + conf.Semester + "學期學期成績單";
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

                    MemoryStream memoryStream = new MemoryStream();
                    document.Save(memoryStream, Aspose.Words.SaveFormat.Doc);
                    ePaperCloud ePaperCloud = new ePaperCloud();
                    int schoolYear, semester;
                    schoolYear = Convert.ToInt32(conf.SchoolYear);
                    semester = Convert.ToInt32(conf.Semester);
                    ePaperCloud.upload_ePaper(schoolYear, semester, reportName, "", memoryStream, ePaperCloud.ViewerType.Student, ePaperCloud.FormatType.Docx);

                    #endregion
                    FISCA.Presentation.MotherForm.SetStatusBarMessage("期末成績單產生完成。", 100);
                    if (overflowRecords.Count > 0)
                        MessageBox.Show(err);
                    if (exc != null)
                    {
                        //throw new Exception("產生期末成績單發生錯誤", exc);
                    }
                };
                bkw.DoWork += delegate (object sender, System.ComponentModel.DoWorkEventArgs e)
                {
                    FISCA.Data.QueryHelper qh = new FISCA.Data.QueryHelper();
                    FISCA.Data.QueryHelper qh1 = new FISCA.Data.QueryHelper();

                    var studentRecords = accessHelper.StudentHelper.GetStudents(selectedStudents);

                    // 2019/01/30 穎驊新增學生英文姓名欄位
                    string strStudentEnglishNameSQL = "select id,english_name from student";
                    System.Data.DataTable StudentEnglishName_dt = qh.Select(strStudentEnglishNameSQL);
                    foreach (System.Data.DataRow dr in StudentEnglishName_dt.Rows)
                    {
                        var studentRecord = studentRecords.Find(s => s.StudentID == "" + dr["id"]);

                        if (studentRecord != null)
                        {
                            if (studentRecord.StudentID == "" + dr["id"])
                            {
                                studentRecord.Fields.Add("英文姓名", "" + dr["english_name"]);
                            }
                        }
                    }

                    // 學生評量缺考資料
                    Dictionary<string, SHStudentExamExtension.DAO.StudSceTakeInfo> StudSceTakeInfoDict = new Dictionary<string, SHStudentExamExtension.DAO.StudSceTakeInfo>();

                    Dictionary<string, Dictionary<string, Dictionary<string, ExamScoreInfo>>> studentExamSores = new Dictionary<string, Dictionary<string, Dictionary<string, ExamScoreInfo>>>();
                    Dictionary<string, Dictionary<string, ExamScoreInfo>> studentRefExamSores = new Dictionary<string, Dictionary<string, ExamScoreInfo>>();
                    //課程上指定學年科目名稱(KEY: courseID,)
                    Dictionary<string, Dictionary<int, string>> CourseSpecifySubjectNameDic = Utility.GetCourseSpecifySubjectNameDict(conf.SchoolYear, conf.Semester);

                    ManualResetEvent scoreReady = new ManualResetEvent(false);
                    ManualResetEvent elseReady = new ManualResetEvent(false);
                    #region 偷跑取得考試成績
                    // 有成績科目名稱對照
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
                                // 取得學生特定學期評量缺考資料
                                StudSceTakeInfoDict = StudentExam.GetStudSceTakeInfoDict(sSchoolYear, sSemester, selectedStudents);

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
                                            string newExamSubject = courseRecord.Subject;
                                            if (CourseSpecifySubjectNameDic.ContainsKey(courseRecord.Subject))
                                                if (CourseSpecifySubjectNameDic[courseRecord.Subject].ContainsKey(courseRecord.CourseID))
                                                    newExamSubject = CourseSpecifySubjectNameDic[courseRecord.Subject][courseRecord.CourseID];

                                            if (!studentExamSores.ContainsKey(attendStudent.StudentID))
                                                studentExamSores.Add(attendStudent.StudentID, new Dictionary<string, Dictionary<string, ExamScoreInfo>>());
                                            if (!studentExamSores[attendStudent.StudentID].ContainsKey(newExamSubject))
                                                studentExamSores[attendStudent.StudentID].Add(newExamSubject, new Dictionary<string, ExamScoreInfo>());
                                            studentExamSores[attendStudent.StudentID][newExamSubject].Add("" + attendStudent.CourseID, null);
                                        }
                                        foreach (var examScoreRec in courseRecord.ExamScoreList)
                                        {
                                            if (examScoreRec.ExamName == conf.ExamRecord.Name)
                                            {
                                                string newExamSubject = courseRecord.Subject;
                                                if (CourseSpecifySubjectNameDic.ContainsKey(courseRecord.Subject))
                                                    if (CourseSpecifySubjectNameDic[courseRecord.Subject].ContainsKey(courseRecord.CourseID))
                                                        newExamSubject = CourseSpecifySubjectNameDic[courseRecord.Subject][courseRecord.CourseID];

                                                studentExamSores[examScoreRec.StudentID][newExamSubject]["" + examScoreRec.CourseID] = examScoreRec;
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
                            // 拿掉 conf.WithSchoolYearScore 判斷，因為學校只有學期不需要學年成績
                            if (sSemester == 2)
                            {
                                accessHelper.StudentHelper.FillSchoolYearEntryScore(false, studentRecords);
                                accessHelper.StudentHelper.FillSchoolYearSubjectScore(false, studentRecords);
                            }
                            accessHelper.StudentHelper.FillSemesterEntryScore(false, studentRecords);
                            accessHelper.StudentHelper.FillSemesterSubjectScore(false, studentRecords);
                            accessHelper.StudentHelper.FillSemesterMoralScore(false, studentRecords);
                            //accessHelper.StudentHelper.FillField("SemesterEntryClassRating", studentRecords);
                            accessHelper.StudentHelper.FillField("SchoolYearEntryClassRating", studentRecords);

                            System.Xml.XmlDocument doc = new System.Xml.XmlDocument();
                            string sidList = "";
                            Dictionary<string, StudentRecord> stuDictionary = new Dictionary<string, StudentRecord>();
                            foreach (var stuRec in studentRecords)
                            {
                                sidList += (sidList == "" ? "" : ",") + stuRec.StudentID;
                                stuDictionary.Add(stuRec.StudentID, stuRec);
                            }

                            #region 學期學業成績排名




                            #endregion
                            #region 學期科目成績排名






                            #endregion
                            accessHelper.StudentHelper.FillAttendance(studentRecords);
                            accessHelper.StudentHelper.FillReward(studentRecords);
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
                        bkw.ReportProgress(0);
                        #region 日常行為表現資料表
                        SmartSchool.Customization.Data.SystemInformation.getField("文字評量對照表");
                        foreach (System.Xml.XmlElement ele in (SmartSchool.Customization.Data.SystemInformation.Fields["文字評量對照表"] as System.Xml.XmlElement).SelectNodes("Content/Morality"))
                        {
                            string face = ele.GetAttribute("Face");
                            if (!table.Columns.Contains("綜合表現：" + face))
                            {
                                table.Columns.Add("綜合表現：" + face);
                            }

                        }
                        // 2021-12-28 Cynthia 避免舊的變數不能使用，增加新版文字評量變數，舊版綜合表現OOO不移除
                        // 2023/8/15，CT,新增支援顯示上學期文字評量
                        for (int i = 1; i <= 10; i++)
                        {
                            table.Columns.Add("文字評量" + i);
                            table.Columns.Add("上學期文字評量" + i);
                        }
                        for (int i = 1; i <= 10; i++)
                        {
                            table.Columns.Add("文字評量名稱" + i);
                            table.Columns.Add("上學期文字評量名稱" + i);
                        }

                        #endregion
                        #region 缺曠對照表
                        List<K12.Data.PeriodMappingInfo> periodMappingInfos = K12.Data.PeriodMapping.SelectAll();
                        Dictionary<string, string> dicPeriodMappingType = new Dictionary<string, string>();
                        List<string> periodTypes = new List<string>();
                        foreach (K12.Data.PeriodMappingInfo periodMappingInfo in periodMappingInfos)
                        {
                            if (!dicPeriodMappingType.ContainsKey(periodMappingInfo.Name))
                                dicPeriodMappingType.Add(periodMappingInfo.Name, periodMappingInfo.Type);

                            if (!periodTypes.Contains(periodMappingInfo.Type))
                                periodTypes.Add(periodMappingInfo.Type);
                        }
                        foreach (var absence in K12.Data.AbsenceMapping.SelectAll())
                        {
                            foreach (var pt in periodTypes)
                            {
                                string attendanceKey = pt + "_" + absence.Name;
                                if (!table.Columns.Contains(attendanceKey))
                                {
                                    table.Columns.Add(attendanceKey);
                                }
                            }
                        }
                        #endregion
                        bkw.ReportProgress(3);
                        #region 整理學生住址
                        accessHelper.StudentHelper.FillContactInfo(studentRecords);
                        #endregion
                        #region 整理學生父母及監護人
                        accessHelper.StudentHelper.FillParentInfo(studentRecords);
                        #endregion
                        bkw.ReportProgress(10);
                        #region 整理同年級學生
                        //整理選取學生的年級
                        Dictionary<string, List<StudentRecord>> gradeyearStudents = new Dictionary<string, List<StudentRecord>>();
                        foreach (var studentRec in studentRecords)
                        {
                            string grade = "";
                            if (studentRec.RefClass != null)
                                grade = "" + studentRec.RefClass.GradeYear;
                            if (!gradeyearStudents.ContainsKey(grade))
                                gradeyearStudents.Add(grade, new List<StudentRecord>());
                            gradeyearStudents[grade].Add(studentRec);
                        }
                        foreach (var classRec in accessHelper.ClassHelper.GetAllClass())
                        {
                            if (gradeyearStudents.ContainsKey("" + classRec.GradeYear))
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
                        //List<string> gradeyearClasses = new List<string>();
                        //foreach (ClassRecord classRec in K12.Data.Class.SelectAll())
                        //{
                        //    if (gradeyearStudents.ContainsKey("" + classRec.GradeYear))
                        //    {
                        //        gradeyearClasses.Add(classRec.ID);
                        //    }
                        //}
                        //用班級去取出可能有相關的學生
                        //foreach (SHStudentRecord studentRec in SHStudent.SelectByClassIDs(gradeyearClasses))
                        //{
                        //    string grade = "";
                        //    if (studentRec.Class != null)
                        //        grade = "" + studentRec.Class.GradeYear;
                        //    if (!gradeyearStudents[grade].Contains(studentRec.ID))
                        //        gradeyearStudents[grade].Add(studentRec.ID);
                        //    if (!studentRecords.ContainsKey(studentRec.ID))
                        //        studentRecords.Add(studentRec.ID, studentRec);
                        //}
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
                        //Dictionary<string, string> studentTag1Group = new Dictionary<string, string>();
                        //Dictionary<string, string> studentTag2Group = new Dictionary<string, string>();
                        //Dictionary<string, List<decimal>> ranks = new Dictionary<string, List<decimal>>();
                        //Dictionary<string, List<string>> rankStudents = new Dictionary<string, List<string>>();
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
                        //Dictionary<string, decimal> analytics = new Dictionary<string, decimal>();
                        int total = 0;
                        foreach (var gss in gradeyearStudents.Values)
                        {
                            total += gss.Count;
                        }
                        bkw.ReportProgress(40);


                        // 取得學生ID
                        List<string> StudentID9DList = new List<string>();
                        foreach (string gradeyear in gradeyearStudents.Keys)
                        {
                            //找出全年級學生
                            foreach (var studentRec in gradeyearStudents[gradeyear])
                            {
                                StudentID9DList.Add(studentRec.StudentID);
                            }
                        }

                        // 取得人數統計
                        StudentCountInfo sci = new StudentCountInfo();
                        sci.SetStudentSourceIDs(StudentID9DList);
                        sci.CalcStudents();
                        foreach (string name in sci.GetPrefixList())
                        {
                            string colName = name + "人數";
                            if (!table.Columns.Contains(colName))
                                table.Columns.Add(colName);

                        }

                        foreach (string name in sci.GetFullNameTagList())
                        {
                            string colName = name + "人數";
                            if (!table.Columns.Contains(colName))
                                table.Columns.Add(colName);

                        }

                        // 取得學生課程規劃表9D科目名稱
                        Dictionary<string, List<string>> Student9DSubjectNameDict = Utility.GetStudent9DSubjectNameByID(StudentID9DList);

                        foreach (string gradeyear in gradeyearStudents.Keys)
                        {
                            //找出全年級學生
                            foreach (var studentRec in gradeyearStudents[gradeyear])
                            {
                                string studentID = studentRec.StudentID;
                                bool rank = true;

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
                                        // 需要過濾9D課程不計算總分與平均成績
                                        bool isClac = true;
                                        if (Student9DSubjectNameDict.ContainsKey(studentID))
                                        {
                                            if (Student9DSubjectNameDict[studentID].Contains(subjectName))
                                            {
                                                isClac = false;
                                            }
                                        }

                                        //if (conf.PrintSubjectList.Contains(subjectName))
                                        //{
                                        #region 是列印科目
                                        foreach (var sceTakeRecord in studentExamSores[studentID][subjectName].Values)
                                        {
                                            if (sceTakeRecord != null && sceTakeRecord.ExamScore != -2)
                                            {

                                                if (isClac)
                                                {
                                                    decimal examScore = sceTakeRecord.ExamScore;
                                                    if (sceTakeRecord.SpecialCase == "缺")
                                                        examScore = 0;
                                                    
                                                    if (examScore == -1)
                                                        examScore = 0;

                                                    printSubjectSum += examScore;//計算總分
                                                    printSubjectCount++;
                                                    //計算加權總分
                                                    printSubjectSumW += examScore * sceTakeRecord.CreditDec();
                                                    printSubjectCreditSum += sceTakeRecord.CreditDec();
                                                }


                                                //    Console.WriteLine(studentID + "," + sceTakeRecord.ExamName + "," + sceTakeRecord.ExamScore + "," + printSubjectSumW + "," + printSubjectCreditSum + "," + sceTakeRecord.Subject);
                                            }
                                            else
                                            {
                                                summaryRank = false;
                                            }
                                        }
                                        #endregion
                                        // }

                                    }
                                    if (printSubjectCount > 0)
                                    {
                                        #region 有列印科目處理加總成績
                                        //總分
                                        studentPrintSubjectSum.Add(studentID, printSubjectSum);
                                        //平均四捨五入至小數點第二位
                                        studentPrintSubjectAvg.Add(studentID, Math.Round(printSubjectSum / printSubjectCount, 2, MidpointRounding.AwayFromZero));

                                        #endregion
                                        if (printSubjectCreditSum > 0)
                                        {
                                            #region 有總學分數處理加總
                                            //加權總分
                                            studentPrintSubjectSumW.Add(studentID, printSubjectSumW);
                                            //加權平均四捨五入至小數點第二位
                                            studentPrintSubjectAvgW.Add(studentID, Math.Round(printSubjectSumW / printSubjectCreditSum, 2, MidpointRounding.AwayFromZero));

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


                                    }
                                    //類別2總分平均排名
                                    if (tag2SubjectCount > 0)
                                    {
                                        //總分
                                        studentTag2SubjectSum.Add(studentID, tag2SubjectSum);
                                        //平均四捨五入至小數點第二位
                                        studentTag2SubjectAvg.Add(studentID, Math.Round(tag2SubjectSum / tag2SubjectCount, 2, MidpointRounding.AwayFromZero));

                                        //類別2加權總分平均排名
                                        if (tag2SubjectCreditSum > 0)
                                        {
                                            studentTag2SubjectSumW.Add(studentID, tag2SubjectSumW);
                                            studentTag2SubjectAvgW.Add(studentID, Math.Round(tag2SubjectSumW / tag2SubjectCreditSum, 2, MidpointRounding.AwayFromZero));

                                        }
                                    }
                                }
                                progressCount++;
                                bkw.ReportProgress(40 + progressCount * 30 / total);
                            }
                        }

                        #endregion

                        // 先取得 K12 StudentRec,因為後面透過 k12.data 取資料有的傳入ID,有的傳入 Record 有點亂
                        List<K12.Data.StudentRecord> StudRecList = new List<K12.Data.StudentRecord>();
                        List<string> StudIDList = (from data in studentRecords select data.StudentID).ToList();
                        StudRecList = K12.Data.Student.SelectByIDs(StudIDList);

                        int SchoolYear, Semester;
                        int.TryParse(conf.SchoolYear, out SchoolYear);
                        int.TryParse(conf.Semester, out Semester);

                        Dictionary<string, decimal> ServiceLearningByDateDict2 = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> ServiceLearningByDateDict1 = new Dictionary<string, decimal>();
                        Dictionary<string, decimal> ServiceLearningByDateDict = new Dictionary<string, decimal>();
                        Dictionary<string, Dictionary<string, int>> AttendanceCountDict = new Dictionary<string, Dictionary<string, int>>();
                        Dictionary<string, Dictionary<string, int>> AttendanceCountDict1 = new Dictionary<string, Dictionary<string, int>>();
                        Dictionary<string, Dictionary<string, int>> AttendanceCountDict2 = new Dictionary<string, Dictionary<string, int>>();

                        // 取得暫存資料 學習服務區間時數                       
                        // 本學期
                        ServiceLearningByDateDict2 = Utility.GetServiceLearningBySchoolYearSemester(StudIDList, SchoolYear, Semester);

                        // 取得缺曠本學期
                        AttendanceCountDict2 = Utility.GetAttendanceCountBySchoolYearSemester(StudRecList, SchoolYear, Semester);

                        if (Semester == 2)
                        {
                            // 前學期
                            ServiceLearningByDateDict1 = Utility.GetServiceLearningBySchoolYearSemester(StudIDList, SchoolYear, 1);
                            // 學年
                            ServiceLearningByDateDict = Utility.GetServiceLearningBySchoolYear(StudIDList, SchoolYear);

                            // 取得學年缺曠
                            AttendanceCountDict = Utility.GetAttendanceCountBySchoolYear(StudRecList, SchoolYear);

                            // 取得缺曠前學期
                            AttendanceCountDict1 = Utility.GetAttendanceCountBySchoolYearSemester(StudRecList, SchoolYear, 1);
                        }

                        List<K12.Data.PeriodMappingInfo> PeriodMappingList = K12.Data.PeriodMapping.SelectAll();
                        // 節次>類別
                        Dictionary<string, string> PeriodMappingDict = new Dictionary<string, string>();
                        foreach (K12.Data.PeriodMappingInfo rec in PeriodMappingList)
                        {
                            if (!PeriodMappingDict.ContainsKey(rec.Name))
                                PeriodMappingDict.Add(rec.Name, rec.Type);
                        }

                        bkw.ReportProgress(70);
                        elseReady.WaitOne();

                        _studPassSumCreditDict1.Clear();
                        _studPassSumCreditDictAll.Clear();
                        _studPassSumCreditDictC1.Clear();
                        _studPassSumCreditDictC2.Clear();

                        StudentSemsReqSumCreditsDict.Clear();
                        StudentSemsSelSumCreditsDict.Clear();
                        StudentSemsPassReqSumCreditsDict.Clear();
                        StudentSemsPassSelSumCreditsDict.Clear();
                        StudentAllPassReqSumCreditsDict.Clear();
                        StudentAllPassSelSumCreditsDict.Clear();


                        progressCount = 0;
                        #region 填入資料表
                        foreach (var stuRec in studentRecords)
                        {
                            // 本學期取得學分數
                            if (!_studPassSumCreditDict1.ContainsKey(stuRec.StudentID))
                                _studPassSumCreditDict1.Add(stuRec.StudentID, 0);

                            // 累計取得學分數
                            if (!_studPassSumCreditDictAll.ContainsKey(stuRec.StudentID))
                                _studPassSumCreditDictAll.Add(stuRec.StudentID, 0);

                            if (!_studPassSumCreditDictC1.ContainsKey(stuRec.StudentID))
                                _studPassSumCreditDictC1.Add(stuRec.StudentID, 0);

                            if (!_studPassSumCreditDictC2.ContainsKey(stuRec.StudentID))
                                _studPassSumCreditDictC2.Add(stuRec.StudentID, 0);


                            if (!StudentSemsReqSumCreditsDict.ContainsKey(stuRec.StudentID))
                                StudentSemsReqSumCreditsDict.Add(stuRec.StudentID, 0);
                            if (!StudentSemsSelSumCreditsDict.ContainsKey(stuRec.StudentID))
                                StudentSemsSelSumCreditsDict.Add(stuRec.StudentID, 0);
                            if (!StudentSemsPassReqSumCreditsDict.ContainsKey(stuRec.StudentID))
                                StudentSemsPassReqSumCreditsDict.Add(stuRec.StudentID, 0);
                            if (!StudentSemsPassSelSumCreditsDict.ContainsKey(stuRec.StudentID))
                                StudentSemsPassSelSumCreditsDict.Add(stuRec.StudentID, 0);
                            if (!StudentAllPassReqSumCreditsDict.ContainsKey(stuRec.StudentID))
                                StudentAllPassReqSumCreditsDict.Add(stuRec.StudentID, 0);
                            if (!StudentAllPassSelSumCreditsDict.ContainsKey(stuRec.StudentID))
                                StudentAllPassSelSumCreditsDict.Add(stuRec.StudentID, 0);



                            string studentID = stuRec.StudentID;
                            string gradeYear = (stuRec.RefClass == null ? "" : "" + stuRec.RefClass.GradeYear);
                            DataRow row = table.NewRow();
                            #region 基本資料

                            // 服務學習時數
                            if (ServiceLearningByDateDict1.ContainsKey(studentID))
                            {
                                // 處理學生上學習服務時數	
                                row["前學期服務學習時數"] = ServiceLearningByDateDict1[studentID];
                            }
                            else
                            {
                                row["前學期服務學習時數"] = 0;
                            }

                            if (ServiceLearningByDateDict2.ContainsKey(studentID))
                            {
                                // 處理學生下學習服務時數	
                                row["本學期服務學習時數"] = ServiceLearningByDateDict2[studentID];
                            }
                            else
                            {
                                row["本學期服務學習時數"] = 0;
                            }

                            if (ServiceLearningByDateDict.ContainsKey(studentID))
                            {
                                // 處理學生學年學習服務時數	
                                row["學年服務學習時數"] = ServiceLearningByDateDict[studentID];
                            }
                            else
                            {
                                row["學年服務學習時數"] = 0;
                            }

                            // 處理缺曠
                            if (AttendanceCountDict.ContainsKey(studentID))
                            {
                                Dictionary<string, int> attenceCount = AttendanceCountDict[studentID];

                                foreach (string attendMappingKey in AttendanceMappingKeyList)
                                {
                                    string keyS = "學年" + attendMappingKey;
                                    if (attenceCount.Keys.Contains(attendMappingKey))
                                    {
                                        row[keyS] = attenceCount[attendMappingKey] == 0 ? 0 : attenceCount[attendMappingKey];
                                    }
                                    else
                                    {
                                        row[keyS] = "0";
                                    }
                                }
                            }

                            if (AttendanceCountDict1.ContainsKey(studentID))
                            {
                                Dictionary<string, int> attenceCount = AttendanceCountDict1[studentID];

                                foreach (string attendMappingKey in AttendanceMappingKeyList)
                                {
                                    string keyS = "前學期" + attendMappingKey;
                                    if (attenceCount.Keys.Contains(attendMappingKey))
                                    {
                                        row[keyS] = attenceCount[attendMappingKey] == 0 ? 0 : attenceCount[attendMappingKey];
                                    }
                                    else
                                    {
                                        row[keyS] = "0";
                                    }
                                }
                            }

                            if (AttendanceCountDict2.ContainsKey(studentID))
                            {
                                Dictionary<string, int> attenceCount = AttendanceCountDict2[studentID];

                                foreach (string attendMappingKey in AttendanceMappingKeyList)
                                {
                                    string keyS = "本學期" + attendMappingKey;
                                    if (attenceCount.Keys.Contains(attendMappingKey))
                                    {
                                        row[keyS] = attenceCount[attendMappingKey] == 0 ? 0 : attenceCount[attendMappingKey];
                                    }
                                    else
                                    {
                                        row[keyS] = "0";
                                    }
                                }
                            }

                            row["學校名稱"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChineseName").InnerText;
                            row["學校地址"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("Address").InnerText;
                            row["學校電話"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("Telephone").InnerText;
                            row["校長名稱"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("ChancellorChineseName").InnerText;
                            row["學務主任"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("StuDirectorName").InnerText;
                            row["教務主任"] = K12.Data.School.Configuration["學校資訊"].PreviousData.SelectSingleNode("EduDirectorName").InnerText;
                            row["收件人地址"] = stuRec.ContactInfo.MailingAddress.FullAddress != "" ?
                                                stuRec.ContactInfo.MailingAddress.FullAddress : stuRec.ContactInfo.PermanentAddress.FullAddress;
                            row["收件人"] = stuRec.ParentInfo.CustodianName != "" ? stuRec.ParentInfo.CustodianName :
                                                (stuRec.ParentInfo.FatherName != "" ? stuRec.ParentInfo.FatherName :
                                                    (stuRec.ParentInfo.FatherName != "" ? stuRec.ParentInfo.MotherName : stuRec.StudentName));
                            //«通訊地址»«通訊地址郵遞區號»«通訊地址內容»
                            //«戶籍地址»«戶籍地址郵遞區號»«戶籍地址內容»
                            //«監護人»«父親»«母親»«科別名稱»
                            row["通訊地址"] = stuRec.ContactInfo.MailingAddress.FullAddress;
                            row["通訊地址郵遞區號"] = stuRec.ContactInfo.MailingAddress.ZipCode;
                            row["通訊地址內容"] = stuRec.ContactInfo.MailingAddress.County + stuRec.ContactInfo.MailingAddress.Town + stuRec.ContactInfo.MailingAddress.DetailAddress;
                            row["戶籍地址"] = stuRec.ContactInfo.PermanentAddress.FullAddress;
                            row["戶籍地址郵遞區號"] = stuRec.ContactInfo.PermanentAddress.ZipCode;
                            row["戶籍地址內容"] = stuRec.ContactInfo.PermanentAddress.County + stuRec.ContactInfo.PermanentAddress.Town + stuRec.ContactInfo.PermanentAddress.DetailAddress;
                            row["監護人"] = stuRec.ParentInfo.CustodianName;
                            row["父親"] = stuRec.ParentInfo.FatherName;
                            row["母親"] = stuRec.ParentInfo.MotherName;
                            row["科別名稱"] = stuRec.Department;
                            row["試別"] = conf.ExamRecord.Name;

                            row["系統編號"] = "系統編號{" + stuRec.StudentID + "}";
                            row["學年度"] = conf.SchoolYear;
                            row["學期"] = conf.Semester;
                            row["系統學年度"] = K12.Data.School.DefaultSchoolYear;
                            row["系統學期"] = K12.Data.School.DefaultSemester;
                            row["班級科別名稱"] = stuRec.RefClass == null ? "" : stuRec.RefClass.Department;
                            row["班級"] = stuRec.RefClass == null ? "" : stuRec.RefClass.ClassName;
                            row["班導師"] = (stuRec.RefClass == null || stuRec.RefClass.RefTeacher == null) ? "" : stuRec.RefClass.RefTeacher.TeacherName;
                            row["座號"] = stuRec.SeatNo;
                            row["學號"] = stuRec.StudentNumber;
                            row["姓名"] = stuRec.StudentName;
                            row["英文姓名"] = stuRec.Fields.ContainsKey("英文姓名") ? stuRec.Fields["英文姓名"] : "";

                            //護照資料

                            string strSQL1 = "select nationality1, passport_name1, nat1.eng_name as nat_eng1, nationality2, passport_name2, nat2.eng_name as nat_eng2, nationality3, passport_name3, nat3.eng_name as nat_eng3 from student_info_ext  as stud_info left outer join $ischool.mapping.nationality as nat1 on nat1.name = stud_info.nationality1 left outer join $ischool.mapping.nationality as nat2 on nat2.name = stud_info.nationality2 left outer join $ischool.mapping.nationality as nat3 on nat3.name = stud_info.nationality3 WHERE ref_student_id=" + stuRec.StudentID;
                            DataTable student_info_ext = qh1.Select(strSQL1);
                            if (student_info_ext.Rows.Count > 0)
                            {
                                row["國籍一"] = student_info_ext.Rows[0]["nationality1"];
                                row["國籍一護照名"] = student_info_ext.Rows[0]["passport_name1"];
                                row["國籍二"] = student_info_ext.Rows[0]["nationality2"];
                                row["國籍二護照名"] = student_info_ext.Rows[0]["passport_name2"];
                                row["國籍一英文"] = student_info_ext.Rows[0]["nat_eng1"];
                                row["國籍二英文"] = student_info_ext.Rows[0]["nat_eng2"];
                            }
                            else
                            {
                                row["國籍一"] = "";
                                row["國籍一護照名"] = "";
                                row["國籍二"] = "";
                                row["國籍二護照名"] = "";
                                row["國籍一英文"] = "";
                                row["國籍二英文"] = "";
                            }

                            row["定期評量"] = conf.ExamRecord.Name;
                            #endregion
                            #region 成績資料
                            #region 各科成績資料
                            #region 分項成績
                            int currentGradeYear = -1;
                            foreach (var semesterEntryScore in stuRec.SemesterEntryScoreList)
                            {
                                if (("" + semesterEntryScore.SchoolYear) == conf.SchoolYear && ("" + semesterEntryScore.Semester) == conf.Semester)
                                {
                                    row["學期" + semesterEntryScore.Entry + "成績"] = semesterEntryScore.Score;
                                    currentGradeYear = semesterEntryScore.GradeYear;
                                }
                            }

                            #region 學期學業成績排名


                            #endregion

                            if (conf.Semester == "2")
                            {
                                #region 學年學業成績及排名

                                // 先不判斷 WithSchoolYearScore，因為有報表只有學期
                                //if (conf.WithSchoolYearScore)
                                //{
                                foreach (var schoolYearEntryScore in stuRec.SchoolYearEntryScoreList)
                                {
                                    if (("" + schoolYearEntryScore.SchoolYear) == conf.SchoolYear)
                                    {
                                        string keyY = "學年" + schoolYearEntryScore.Entry + "成績";
                                        if (table.Columns.Contains(keyY))
                                        {
                                            row[keyY] = schoolYearEntryScore.Score;
                                        }


                                    }
                                }
                                if (stuRec.Fields.ContainsKey("SchoolYearEntryClassRating"))
                                {
                                    System.Xml.XmlElement _sems_ratings = stuRec.Fields["SchoolYearEntryClassRating"] as System.Xml.XmlElement;
                                    string path = string.Format("SchoolYearEntryScore[SchoolYear='{0}']/ClassRating/Rating/Item[@分項='學業']/@排名", conf.SchoolYear);
                                    System.Xml.XmlNode result = _sems_ratings.SelectSingleNode(path);
                                    if (result != null)
                                    {
                                        if (table.Columns.Contains("學年學業成績班排名"))
                                            row["學年學業成績班排名"] = result.InnerText;
                                    }
                                }
                                //}


                                #endregion
                                if (conf.WithPrevSemesterScore)
                                {
                                    foreach (var semesterEntryScore in stuRec.SemesterEntryScoreList)
                                    {
                                        if (semesterEntryScore.Semester == 1 && semesterEntryScore.GradeYear == currentGradeYear)
                                        {
                                            row["上學期" + semesterEntryScore.Entry + "成績"] = semesterEntryScore.Score;
                                        }
                                    }
                                }
                            }






                            #endregion
                            #region 整理科目順序
                            List<string> subjects1 = new List<string>();//本學期
                            List<string> subjects2 = new List<string>();//上學期
                            List<string> subjects3 = new List<string>();//學年
                            // 進校判斷升級重讀: key= 科目名稱, Value=上下學期實得學分加總
                            Dictionary<string, decimal> subjectCreditDic = new Dictionary<string, decimal>();

                            foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                            {
                                if (semesterSubjectScore.SchoolYear.ToString() == conf.SchoolYear && semesterSubjectScore.Detail.GetAttribute("不計學分") != "是")
                                {
                                    string skey = semesterSubjectScore.Subject;
                                    if (semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱") != "")
                                        skey = semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱");

                                    if (!subjectCreditDic.ContainsKey(skey))
                                        subjectCreditDic.Add(skey, 0);
                                    if (semesterSubjectScore.Pass)
                                        subjectCreditDic[skey] += semesterSubjectScore.CreditDec();

                                }

                                if (("" + semesterSubjectScore.SchoolYear) == conf.SchoolYear && ("" + semesterSubjectScore.Semester) == conf.Semester)
                                {
                                    if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是")
                                    {
                                        //科目_科目代碼
                                        //subjects1.Add(semesterSubjectScore.Subject);
                                        //指定學年科目名稱
                                        if (semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱") != "")
                                            subjects1.Add(semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱"));
                                        else
                                            subjects1.Add(semesterSubjectScore.Subject);
                                        currentGradeYear = semesterSubjectScore.GradeYear;
                                    }
                                }
                            }
                            if (studentExamSores.ContainsKey(stuRec.StudentID))
                            {
                                foreach (var subjectName in studentExamSores[studentID].Keys)
                                {
                                    foreach (var courseID in studentExamSores[studentID][subjectName].Keys)
                                    {
                                        if (conf.PrintSubjectList.Contains(subjectName))
                                        {
                                            #region 跟學期成績做差異新增
                                            string newExamSubjectName = subjectName;
                                            if (CourseSpecifySubjectNameDic.ContainsKey(subjectName))
                                                if (CourseSpecifySubjectNameDic[subjectName].ContainsKey(accessHelper.CourseHelper.GetCourse(courseID)[0].CourseID))
                                                    newExamSubjectName = CourseSpecifySubjectNameDic[subjectName][accessHelper.CourseHelper.GetCourse(courseID)[0].CourseID];

                                            bool match = false;
                                            foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                                            {
                                                string newSemesterSubjectName = semesterSubjectScore.Subject;
                                                if (semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱") != "")
                                                    newSemesterSubjectName = semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱");

                                                if (("" + semesterSubjectScore.SchoolYear) == conf.SchoolYear
                                                    && ("" + semesterSubjectScore.Semester) == conf.Semester
                                                    && newSemesterSubjectName == newExamSubjectName
                                                    && semesterSubjectScore.Level == accessHelper.CourseHelper.GetCourse(courseID)[0].SubjectLevel)
                                                {
                                                    match = true;
                                                    break;
                                                }
                                            }
                                            if (!match)
                                            {
                                                //指定學年科目名稱
                                                subjects1.Add(newExamSubjectName);
                                            }
                                            #endregion
                                        }
                                    }
                                }
                            }
                            if (conf.Semester == "2")
                            {
                                if (conf.WithPrevSemesterScore)
                                {
                                    foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                                    {
                                        if (semesterSubjectScore.Semester == 1 && semesterSubjectScore.GradeYear == currentGradeYear)
                                        {
                                            if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是")
                                                if (semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱") != "")
                                                    subjects2.Add(semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱"));
                                                else
                                                    subjects2.Add(semesterSubjectScore.Subject);
                                        }
                                    }
                                }
                                if (conf.WithSchoolYearScore)
                                {
                                    decimal entryScore = 0;
                                    foreach (var schoolYearEntryScore in stuRec.SchoolYearEntryScoreList)
                                    {
                                        if (schoolYearEntryScore.SchoolYear.ToString() == conf.SchoolYear)
                                        {
                                            if (schoolYearEntryScore.Entry == "學業")
                                                entryScore = schoolYearEntryScore.Score;
                                        }
                                    }
                                    //RetainInTheSameGrade(string pStudentID, int schoolYear, decimal entryScore, List<SchoolYearSubjectScoreInfo> schoolYearSubjectScoreInfos, Dictionary<string, decimal> subjectCreditDic)
                                    row["升級或應重讀"] = RetainInTheSameGrade(stuRec.StudentID, conf.SchoolYear, currentGradeYear, entryScore, stuRec.SchoolYearSubjectScoreList, subjectCreditDic) ? "重讀" : "升級";
                                    foreach (var schoolYearSubjectScore in stuRec.SchoolYearSubjectScoreList)
                                    {
                                        if (("" + schoolYearSubjectScore.SchoolYear) == conf.SchoolYear)
                                        {
                                            subjects3.Add(schoolYearSubjectScore.Subject);
                                        }
                                    }
                                }
                            }
                            var subjectNameList = new List<string>();
                            subjectNameList.AddRange(subjects1);
                            foreach (var subject in subjects1)
                            {
                                if (subjects2.Contains(subject)) subjects2.Remove(subject);
                                if (subjects3.Contains(subject)) subjects3.Remove(subject);
                            }
                            subjectNameList.AddRange(subjects2);
                            foreach (var subject in subjects2)
                            {
                                if (subjects3.Contains(subject)) subjects3.Remove(subject);
                            }
                            subjectNameList.AddRange(subjects3);

                            //subjectNameList.Sort(new StringComparer("國文"
                            //                , "英文"
                            //                , "數學"
                            //                , "理化"
                            //                , "生物"
                            //                , "社會"
                            //                , "物理"
                            //                , "化學"
                            //                , "歷史"
                            //                , "地理"
                            //                , "公民"));

                            // 取得中英文對照表中文名稱
                            List<string> ChineseSubjectNameMapList = Utility.GetChineseSubjectNameList();

                            // 科目排序
                            subjectNameList.Sort(new StringComparer(ChineseSubjectNameMapList.ToArray()));

                            #endregion


                            // 處理本學期取得學分與累計取得學分
                            foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                            {
                                if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是")
                                {
                                    // 本學期已修
                                    if (semesterSubjectScore.SchoolYear.ToString() == conf.SchoolYear && semesterSubjectScore.Semester.ToString() == conf.Semester)
                                    {
                                        if (semesterSubjectScore.Require)
                                            StudentSemsReqSumCreditsDict[stuRec.StudentID] += semesterSubjectScore.CreditDec();
                                        else
                                            StudentSemsSelSumCreditsDict[stuRec.StudentID] += semesterSubjectScore.CreditDec();
                                    }


                                    // 本學期取得
                                    if (semesterSubjectScore.SchoolYear.ToString() == conf.SchoolYear && semesterSubjectScore.Semester.ToString() == conf.Semester && semesterSubjectScore.Pass)
                                    {
                                        _studPassSumCreditDict1[stuRec.StudentID] += semesterSubjectScore.CreditDec();

                                        if (semesterSubjectScore.Require)
                                            StudentSemsPassReqSumCreditsDict[stuRec.StudentID] += semesterSubjectScore.CreditDec();
                                        else
                                            StudentSemsPassSelSumCreditsDict[stuRec.StudentID] += semesterSubjectScore.CreditDec();
                                    }


                                    // 累計取得
                                    if (semesterSubjectScore.Pass)
                                    {
                                        _studPassSumCreditDictAll[stuRec.StudentID] += semesterSubjectScore.CreditDec();

                                        if (semesterSubjectScore.Require)
                                        {
                                            _studPassSumCreditDictC1[stuRec.StudentID] += semesterSubjectScore.CreditDec();
                                            StudentAllPassReqSumCreditsDict[stuRec.StudentID] += semesterSubjectScore.CreditDec();
                                        }
                                        else
                                        {
                                            _studPassSumCreditDictC2[stuRec.StudentID] += semesterSubjectScore.CreditDec();
                                            StudentAllPassSelSumCreditsDict[stuRec.StudentID] += semesterSubjectScore.CreditDec();
                                        }
                                    }
                                }
                            }

                            row["本學期取得學分數"] = _studPassSumCreditDict1[stuRec.StudentID];
                            row["累計取得學分數"] = _studPassSumCreditDictAll[stuRec.StudentID];
                            row["累計取得必修學分"] = _studPassSumCreditDictC1[stuRec.StudentID];
                            row["累計取得選修學分"] = _studPassSumCreditDictC2[stuRec.StudentID];

                            row["本學期已修必修學分"] = StudentSemsReqSumCreditsDict[stuRec.StudentID];
                            row["本學期已修選修學分"] = StudentSemsSelSumCreditsDict[stuRec.StudentID];
                            row["本學期實得必修學分"] = StudentSemsPassReqSumCreditsDict[stuRec.StudentID];
                            row["本學期實得選修學分"] = StudentSemsPassSelSumCreditsDict[stuRec.StudentID];
                            row["在校期間實得必修學分"] = StudentAllPassReqSumCreditsDict[stuRec.StudentID];
                            row["在校期間實得選修學分"] = StudentAllPassSelSumCreditsDict[stuRec.StudentID];

                            // 取得學生及格與補考標準
                            // 及格
                            decimal o_scA = 0;
                            // 補考
                            decimal o_scB = 0;
                            if (StudentApplyLimitDict.ContainsKey(stuRec.StudentID))
                            {
                                string sA = stuRec.RefClass.GradeYear + "_及";
                                string sB = stuRec.RefClass.GradeYear + "_補";

                                if (StudentApplyLimitDict[stuRec.StudentID].ContainsKey(sA))
                                    o_scA = StudentApplyLimitDict[stuRec.StudentID][sA];

                                if (StudentApplyLimitDict[stuRec.StudentID].ContainsKey(sB))
                                    o_scB = StudentApplyLimitDict[stuRec.StudentID][sB];
                            }

                            int subjectIndex = 1;
                            // 學期科目與定期評量
                            foreach (string subjectName in subjectNameList)
                            {
                                if (subjectIndex <= conf.SubjectLimit)
                                {
                                    decimal? subjectNumber = null;
                                    bool findInSemesterSubjectScore = false;
                                    bool findInSemester1SubjectScore = false;
                                    bool findInExamScores = false;
                                    #region 本學期學期成績
                                    foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                                    {
                                        string newSemesterSubjectName = semesterSubjectScore.Subject;
                                        if (semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱") != "")
                                            newSemesterSubjectName = semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱");

                                        if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是"
                                            && newSemesterSubjectName == subjectName
                                            && ("" + semesterSubjectScore.SchoolYear) == conf.SchoolYear
                                            && ("" + semesterSubjectScore.Semester) == conf.Semester)
                                        {
                                            findInSemesterSubjectScore = true;


                                            decimal level;
                                            subjectNumber = decimal.TryParse(semesterSubjectScore.Level, out level) ? (decimal?)level : null;

                                            row["領域" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("領域");
                                            row["分項類別" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("開課分項類別");
                                            row["科目名稱" + subjectIndex] = newSemesterSubjectName + GetNumber(subjectNumber);
                                            if (!conf.IsShowLevel)  //2021-12-27 Cynthia 不顯示級別
                                                row["科目名稱" + subjectIndex] = newSemesterSubjectName;
                                            row["科目" + subjectIndex] = newSemesterSubjectName;
                                            row["科目級別" + subjectIndex] = GetNumber(subjectNumber);
                                            row["學分數" + subjectIndex] = semesterSubjectScore.CreditDec();
                                            string required = semesterSubjectScore.Require ? "必修" : "選修";
                                            string requiredBy = semesterSubjectScore.Detail.GetAttribute("修課校部訂");
                                            row["科目必選修" + subjectIndex] = required;
                                            row["科目校部定/科目必選修" + subjectIndex] = "" + requiredBy[0] + required[0];
                                            row["科目註記" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("註記");
                                            row["科目取得學分" + subjectIndex] = semesterSubjectScore.Pass ? "是" : "否";
                                            row["科目未取得學分註記" + subjectIndex] = semesterSubjectScore.Pass ? "" : "\f";

                                            //"原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績"
                                            //if (semesterSubjectScore.Detail.GetAttribute("不需評分") != "是")
                                            {
                                                if (semesterSubjectScore.Pass == false)
                                                {
                                                    row["學期科目未取得學分標示" + subjectIndex] = conf.FailScoreMark;
                                                }

                                                row["學期科目原始成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("原始成績");
                                                row["學期科目補考成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("補考成績");
                                                row["學期科目重修成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("重修成績");
                                                row["學期科目手動調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("擇優採計成績");
                                                row["學期科目學年調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("學年調整成績");
                                                row["學期科目成績" + subjectIndex] = semesterSubjectScore.Score;

                                                if ("" + semesterSubjectScore.Detail.GetAttribute("是否補修成績") == "是")
                                                {
                                                    row["學期科目補修成績標示" + subjectIndex] = conf.RepairScoreMark;
                                                }

                                                if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("原始成績"))
                                                    row["學期科目原始成績註記" + subjectIndex] = "\f";
                                                if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("補考成績"))
                                                {
                                                    row["學期科目補考成績註記" + subjectIndex] = "\f";
                                                    row["學期科目補考成績標示" + subjectIndex] = conf.ReScoreMark;
                                                }
                                                if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("重修成績"))
                                                {
                                                    row["學期科目重修成績註記" + subjectIndex] = "\f";
                                                    row["學期科目重修成績標示" + subjectIndex] = conf.RereadScoreMark;
                                                }

                                                if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("擇優採計成績"))
                                                    row["學期科目手動成績註記" + subjectIndex] = "\f";
                                                if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("學年調整成績"))
                                                    row["學期科目學年成績註記" + subjectIndex] = "\f";

                                                decimal scA = o_scA;
                                                decimal scB = o_scB;
                                                //string keySa = semesterSubjectScore.SchoolYear + "_" + semesterSubjectScore.Semester + "_" + semesterSubjectScore.Subject + "_" + semesterSubjectScore.Level + "_及";
                                                //string keySb = semesterSubjectScore.SchoolYear + "_" + semesterSubjectScore.Semester + "_" + semesterSubjectScore.Subject + "_" + semesterSubjectScore.Level + "_補";

                                                //if (StudentSCAttendApplyLimitDict.ContainsKey(stuRec.StudentID))
                                                //{
                                                //    if (StudentSCAttendApplyLimitDict[stuRec.StudentID].ContainsKey(keySa))
                                                //        scA = StudentSCAttendApplyLimitDict[stuRec.StudentID][keySa];

                                                //    if (StudentSCAttendApplyLimitDict[stuRec.StudentID].ContainsKey(keySb))
                                                //        scB = StudentSCAttendApplyLimitDict[stuRec.StudentID][keySb];
                                                //}

                                                decimal.TryParse(semesterSubjectScore.Detail.GetAttribute("修課及格標準"), out scA);
                                                decimal.TryParse(semesterSubjectScore.Detail.GetAttribute("修課補考標準"), out scB);

                                                // 不及格
                                                if (semesterSubjectScore.Score < scA)
                                                {

                                                    // 可補考
                                                    if (semesterSubjectScore.Score >= scB)
                                                    {
                                                        // 未取得學分、沒有補考成績、沒有重修成績才需要標示
                                                        if (semesterSubjectScore.Detail.GetAttribute("補考成績") == "" && semesterSubjectScore.Detail.GetAttribute("重修成績") == "" && semesterSubjectScore.Pass == false)
                                                        {
                                                            row["學期科目需要補考註記" + subjectIndex] = "\f";
                                                            row["學期科目需要補考標示" + subjectIndex] = conf.NeedReScoreMark;
                                                        }

                                                    }

                                                    decimal ss = 0;
                                                    decimal.TryParse(semesterSubjectScore.Detail.GetAttribute("原始成績"), out ss);


                                                    if ((ss < scB || semesterSubjectScore.Detail.GetAttribute("補考成績") != "" || semesterSubjectScore.Detail.GetAttribute("重修成績") != "") && semesterSubjectScore.Pass == false)
                                                    {
                                                        // 不可補考，須重修，加入沒有取得學分，需要重修

                                                        row["學期科目需要補考標示" + subjectIndex] = "";
                                                        row["學期科目需要重修註記" + subjectIndex] = "\f";
                                                        row["學期科目需要重修標示" + subjectIndex] = conf.NeedRereadScoreMark;

                                                    }


                                                }
                                            }

                                            #region 學期科目班、科、校、類別1、類別2排名、五標、組距

                                            key = "學期科目排名成績" + semesterSubjectScore.Subject + "^^^" + semesterSubjectScore.Level;
                                            if (stuRec.Fields.ContainsKey(key))
                                                row["學期科目排名成績" + subjectIndex] = "" + stuRec.Fields[key];

                                            string ssKey = "";

                                            if (SemsScoreRankMatrixDataDict.ContainsKey(studentID))
                                            {

                                                #region 班排
                                                ssKey = "學期/科目成績_" + semesterSubjectScore.Subject + "_班排名";
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(ssKey))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"] != null)
                                                        row["學期科目班排名" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"] != null)
                                                        row["學期科目班排名母數" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"].ToString();

                                                    // 五標、組距
                                                    foreach (string item2 in r2List)
                                                    {
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey][item2] != null)
                                                        {
                                                            if (r2ParseList.Contains(item2))
                                                            {
                                                                decimal dd;

                                                                if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString(), out dd))
                                                                {
                                                                    row["學期科目班排名" + subjectIndex + "_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                }
                                                            }
                                                            else
                                                            {

                                                                row["學期科目班排名" + subjectIndex + "_" + item2] = SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString();
                                                            }
                                                        }
                                                    }
                                                }

                                                ssKey = "學期/科目成績(原始)_" + semesterSubjectScore.Subject + "_班排名";
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(ssKey))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"] != null)
                                                        row["學期科目(原始)班排名" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"] != null)
                                                        row["學期科目(原始)班排名母數" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"].ToString();

                                                    // 五標、組距
                                                    foreach (string item2 in r2List)
                                                    {
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey][item2] != null)
                                                        {
                                                            if (r2ParseList.Contains(item2))
                                                            {
                                                                decimal dd;

                                                                if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString(), out dd))
                                                                {
                                                                    row["學期科目(原始)班排名" + subjectIndex + "_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                row["學期科目(原始)班排名" + subjectIndex + "_" + item2] = SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString();
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region 科排
                                                ssKey = "學期/科目成績_" + semesterSubjectScore.Subject + "_科排名";
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(ssKey))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"] != null)
                                                        row["學期科目科排名" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"] != null)
                                                        row["學期科目科排名母數" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"].ToString();

                                                    // 五標、組距
                                                    foreach (string item2 in r2List)
                                                    {
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey][item2] != null)
                                                        {
                                                            if (r2ParseList.Contains(item2))
                                                            {
                                                                decimal dd;

                                                                if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString(), out dd))
                                                                {
                                                                    row["學期科目科排名" + subjectIndex + "_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                row["學期科目科排名" + subjectIndex + "_" + item2] = SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString();
                                                            }
                                                        }
                                                    }
                                                }

                                                ssKey = "學期/科目成績(原始)_" + semesterSubjectScore.Subject + "_科排名";
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(ssKey))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"] != null)
                                                        row["學期科目(原始)科排名" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"] != null)
                                                        row["學期科目(原始)科排名母數" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"].ToString();

                                                    // 五標、組距
                                                    foreach (string item2 in r2List)
                                                    {
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey][item2] != null)
                                                        {
                                                            if (r2ParseList.Contains(item2))
                                                            {
                                                                decimal dd;

                                                                if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString(), out dd))
                                                                {
                                                                    row["學期科目(原始)科排名" + subjectIndex + "_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                row["學期科目(原始)科排名" + subjectIndex + "_" + item2] = SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString();
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region 校排
                                                ssKey = "學期/科目成績_" + semesterSubjectScore.Subject + "_年排名";
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(ssKey))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"] != null)
                                                        row["學期科目全校排名" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"] != null)
                                                        row["學期科目全校排名母數" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["picked_grade"] != null)
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey]["picked_grade"].ToString() != "")
                                                            row["學期成績排名採計成績欄位"] = SemsScoreRankMatrixDataDict[studentID][ssKey]["picked_grade"].ToString();

                                                    // 五標、組距
                                                    foreach (string item2 in r2List)
                                                    {
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey][item2] != null)
                                                        {
                                                            if (r2ParseList.Contains(item2))
                                                            {
                                                                decimal dd;

                                                                if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString(), out dd))
                                                                {
                                                                    row["學期科目全校排名" + subjectIndex + "_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                row["學期科目全校排名" + subjectIndex + "_" + item2] = SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString();
                                                            }
                                                        }
                                                    }
                                                }

                                                ssKey = "學期/科目成績(原始)_" + semesterSubjectScore.Subject + "_年排名";
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(ssKey))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"] != null)
                                                        row["學期科目(原始)全校排名" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"] != null)
                                                        row["學期科目(原始)全校排名母數" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"].ToString();

                                                    // 五標、組距
                                                    foreach (string item2 in r2List)
                                                    {
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey][item2] != null)
                                                        {
                                                            if (r2ParseList.Contains(item2))
                                                            {
                                                                decimal dd;

                                                                if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString(), out dd))
                                                                {
                                                                    row["學期科目(原始)全校排名" + subjectIndex + "_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                row["學期科目(原始)全校排名" + subjectIndex + "_" + item2] = SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString();
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region 類1
                                                ssKey = "學期/科目成績_" + semesterSubjectScore.Subject + "_類別1排名";
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(ssKey))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"] != null)
                                                        row["學期科目類別1排名" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"] != null)
                                                        row["學期科目類別1排名母數" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"].ToString();

                                                    // 五標、組距
                                                    foreach (string item2 in r2List)
                                                    {
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey][item2] != null)
                                                        {
                                                            if (r2ParseList.Contains(item2))
                                                            {
                                                                decimal dd;

                                                                if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString(), out dd))
                                                                {
                                                                    row["學期科目類別1排名" + subjectIndex + "_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                row["學期科目類別1排名" + subjectIndex + "_" + item2] = SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString();
                                                            }
                                                        }

                                                    }
                                                }

                                                ssKey = "學期/科目成績(原始)_" + semesterSubjectScore.Subject + "_類別1排名";
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(ssKey))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"] != null)
                                                        row["學期科目(原始)類別1排名" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"] != null)
                                                        row["學期科目(原始)類別1排名母數" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"].ToString();

                                                    // 五標、組距
                                                    foreach (string item2 in r2List)
                                                    {
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey][item2] != null)
                                                        {
                                                            if (r2ParseList.Contains(item2))
                                                            {
                                                                decimal dd;

                                                                if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString(), out dd))
                                                                {
                                                                    row["學期科目(原始)類別1排名" + subjectIndex + "_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                row["學期科目(原始)類別1排名" + subjectIndex + "_" + item2] = SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString();
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion

                                                #region 類2
                                                ssKey = "學期/科目成績_" + semesterSubjectScore.Subject + "_類別2排名";
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(ssKey))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"] != null)
                                                        row["學期科目類別2排名" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"] != null)
                                                        row["學期科目類別2排名母數" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"].ToString();

                                                    // 五標、組距
                                                    foreach (string item2 in r2List)
                                                    {
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey][item2] != null)
                                                        {
                                                            if (r2ParseList.Contains(item2))
                                                            {
                                                                decimal dd;

                                                                if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString(), out dd))
                                                                {
                                                                    row["學期科目類別2排名" + subjectIndex + "_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                row["學期科目類別2排名" + subjectIndex + "_" + item2] = SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString();
                                                            }
                                                        }
                                                    }
                                                }

                                                ssKey = "學期/科目成績(原始)_" + semesterSubjectScore.Subject + "_類別2排名";
                                                if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(ssKey))
                                                {
                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"] != null)
                                                        row["學期科目(原始)類別2排名" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["rank"].ToString();

                                                    if (SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"] != null)
                                                        row["學期科目(原始)類別2排名母數" + subjectIndex] = SemsScoreRankMatrixDataDict[studentID][ssKey]["matrix_count"].ToString();

                                                    // 五標、組距
                                                    foreach (string item2 in r2List)
                                                    {
                                                        if (SemsScoreRankMatrixDataDict[studentID][ssKey][item2] != null)
                                                        {
                                                            if (r2ParseList.Contains(item2))
                                                            {
                                                                decimal dd;

                                                                if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString(), out dd))
                                                                {
                                                                    row["學期科目(原始)類別2排名" + subjectIndex + "_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                }
                                                            }
                                                            else
                                                            {
                                                                row["學期科目(原始)類別2排名" + subjectIndex + "_" + item2] = SemsScoreRankMatrixDataDict[studentID][ssKey][item2].ToString();
                                                            }
                                                        }
                                                    }
                                                }
                                                #endregion
                                            }


                                            #endregion
                                            stuRec.SemesterSubjectScoreList.Remove(semesterSubjectScore);
                                            break;
                                        }
                                    }
                                    #endregion
                                    #region 定期評量成績
                                    // 檢查畫面上定期評量列印科目

                                    //if (conf.PrintSubjectList.Contains(subjectInfoArray[0]))//subjectName要取原本的名稱
                                    {

                                        if (studentExamSores.ContainsKey(studentID))
                                        {
                                            if (studentExamSores[studentID].ContainsKey(subjectName))
                                            {
                                                foreach (var courseID in studentExamSores[studentID][subjectName].Keys)
                                                {
                                                    var sceTakeRecord = studentExamSores[studentID][subjectName][courseID];
                                                    if (sceTakeRecord != null)
                                                    {//有輸入
                                                        if (conf.PrintSubjectList.Contains(sceTakeRecord.Subject))
                                                        {

                                                            if (findInSemesterSubjectScore)
                                                            {
                                                                if (sceTakeRecord.SubjectLevel != "" + subjectNumber)
                                                                {
                                                                    continue;
                                                                }
                                                            }
                                                            findInExamScores = true;
                                                            if (!findInSemesterSubjectScore)
                                                            {
                                                                decimal level;
                                                                subjectNumber = decimal.TryParse(sceTakeRecord.SubjectLevel, out level) ? (decimal?)level : null;
                                                                row["領域" + subjectIndex] = sceTakeRecord.Domain;
                                                                row["分項類別" + subjectIndex] = sceTakeRecord.Entry;
                                                                row["科目名稱" + subjectIndex] = sceTakeRecord.Subject + GetNumber(subjectNumber);
                                                                if (!conf.IsShowLevel)  //2021-12-27 Cynthia 不顯示級別
                                                                    row["科目名稱" + subjectIndex] = sceTakeRecord.Subject;
                                                                row["科目" + subjectIndex] = sceTakeRecord.Subject;
                                                                row["科目級別" + subjectIndex] = GetNumber(subjectNumber);
                                                                string required = sceTakeRecord.Required ? "必修" : "選修";
                                                                string requiredBy = sceTakeRecord.RequiredBy == "部訂" ? "部定" : sceTakeRecord.RequiredBy;
                                                                row["科目必選修" + subjectIndex] = required;
                                                                row["科目校部定/科目必選修" + subjectIndex] = "" + requiredBy[0] + required[0];
                                                                row["學分數" + subjectIndex] = sceTakeRecord.CreditDec();
                                                            }

                                                            // 因為缺免0分或免試調整，SpecialCase="缺",Score-1,0分，Score=-2免試，空白。
                                                            if (sceTakeRecord.SpecialCase == "缺")
                                                            {
                                                                row["科目成績" + subjectIndex] = 0;
                                                            }
                                                            else
                                                            {
                                                                if (sceTakeRecord.ExamScore == -1)
                                                                {
                                                                    row["科目成績" + subjectIndex] = 0;
                                                                }
                                                                else if (sceTakeRecord.ExamScore == -2)
                                                                {
                                                                    row["科目成績" + subjectIndex] = "免";
                                                                }
                                                                else
                                                                {
                                                                    row["科目成績" + subjectIndex] = sceTakeRecord.ExamScore;
                                                                }
                                                            }

                                                            //row["科目成績" + subjectIndex] = sceTakeRecord.SpecialCase == "" ? ("" + sceTakeRecord.ExamScore) : sceTakeRecord.SpecialCase;

                                                            // 判斷如果有缺考原因，使用缺考輸入文字與原因
                                                            string examUseTextKey = sceTakeRecord.StudentID + "_" + sceTakeRecord.CourseID + "_" + sceTakeRecord.ExamName;
                                                            if (StudSceTakeInfoDict.ContainsKey(examUseTextKey))
                                                            {
                                                                row["科目成績" + subjectIndex] = StudSceTakeInfoDict[examUseTextKey].UseText;
                                                                row["科目缺考原因" + subjectIndex] = StudSceTakeInfoDict[examUseTextKey].ReportValue;
                                                            }

                                                            #region 班排名及落點分析
                                                            string k1 = "";
                                                            if (RankMatrixDataDict.ContainsKey(studentID))
                                                            {
                                                                k1 = "定期評量/科目成績_" + sceTakeRecord.Subject + "_班排名";
                                                                if (RankMatrixDataDict[studentID].ContainsKey(k1))
                                                                {
                                                                    if (RankMatrixDataDict[studentID][k1]["rank"] != null)
                                                                        row["班排名" + subjectIndex] = RankMatrixDataDict[studentID][k1]["rank"].ToString();

                                                                    if (RankMatrixDataDict[studentID][k1]["matrix_count"] != null)
                                                                        row["班排名母數" + subjectIndex] = RankMatrixDataDict[studentID][k1]["matrix_count"].ToString();

                                                                    // 五標PR填值
                                                                    foreach (string rItem in r2List)
                                                                    {
                                                                        if (RankMatrixDataDict[studentID][k1][rItem] != null)
                                                                        {

                                                                            // RankMatrixDataDict[studentID][k1][rItem].ToString();

                                                                            if (r2ParseList.Contains(rItem))
                                                                            {
                                                                                decimal dd;
                                                                                if (decimal.TryParse(RankMatrixDataDict[studentID][k1][rItem].ToString(), out dd))
                                                                                {
                                                                                    row["班排名" + subjectIndex + "_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                row["班排名" + subjectIndex + "_" + rItem] = RankMatrixDataDict[studentID][k1][rItem].ToString();
                                                                            }
                                                                        }
                                                                    }
                                                                }

                                                            }
                                                            #endregion
                                                            #region 科排名及落點分析
                                                            if (RankMatrixDataDict.ContainsKey(studentID))
                                                            {
                                                                k1 = "定期評量/科目成績_" + sceTakeRecord.Subject + "_科排名";
                                                                if (RankMatrixDataDict[studentID].ContainsKey(k1))
                                                                {
                                                                    if (RankMatrixDataDict[studentID][k1]["rank"] != null)
                                                                        row["科排名" + subjectIndex] = RankMatrixDataDict[studentID][k1]["rank"].ToString();

                                                                    if (RankMatrixDataDict[studentID][k1]["matrix_count"] != null)
                                                                        row["科排名母數" + subjectIndex] = RankMatrixDataDict[studentID][k1]["matrix_count"].ToString();

                                                                    // 五標PR填值
                                                                    foreach (string rItem in r2List)
                                                                    {
                                                                        if (RankMatrixDataDict[studentID][k1][rItem] != null)
                                                                        {
                                                                            if (r2ParseList.Contains(rItem))
                                                                            {
                                                                                decimal dd;
                                                                                if (decimal.TryParse(RankMatrixDataDict[studentID][k1][rItem].ToString(), out dd))
                                                                                {
                                                                                    row["科排名" + subjectIndex + "_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                row["科排名" + subjectIndex + "_" + rItem] = RankMatrixDataDict[studentID][k1][rItem].ToString();
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                            #region 全校排名及落點分析
                                                            if (RankMatrixDataDict.ContainsKey(studentID))
                                                            {
                                                                k1 = "定期評量/科目成績_" + sceTakeRecord.Subject + "_年排名";
                                                                if (RankMatrixDataDict[studentID].ContainsKey(k1))
                                                                {
                                                                    if (RankMatrixDataDict[studentID][k1]["rank"] != null)
                                                                        row["全校排名" + subjectIndex] = RankMatrixDataDict[studentID][k1]["rank"].ToString();

                                                                    if (RankMatrixDataDict[studentID][k1]["matrix_count"] != null)
                                                                        row["全校排名母數" + subjectIndex] = RankMatrixDataDict[studentID][k1]["matrix_count"].ToString();

                                                                    // 五標PR填值
                                                                    foreach (string rItem in r2List)
                                                                    {
                                                                        if (RankMatrixDataDict[studentID][k1][rItem] != null)
                                                                        {
                                                                            if (r2ParseList.Contains(rItem))
                                                                            {
                                                                                decimal dd;
                                                                                if (decimal.TryParse(RankMatrixDataDict[studentID][k1][rItem].ToString(), out dd))
                                                                                {
                                                                                    row["全校排名" + subjectIndex + "_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                row["全校排名" + subjectIndex + "_" + rItem] = RankMatrixDataDict[studentID][k1][rItem].ToString();
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                            #region 類別1排名及落點分析
                                                            if (RankMatrixDataDict.ContainsKey(studentID))
                                                            {
                                                                k1 = "定期評量/科目成績_" + sceTakeRecord.Subject + "_類別1排名";
                                                                if (RankMatrixDataDict[studentID].ContainsKey(k1))
                                                                {
                                                                    if (RankMatrixDataDict[studentID][k1]["rank"] != null)
                                                                        row["類別1排名" + subjectIndex] = RankMatrixDataDict[studentID][k1]["rank"].ToString();

                                                                    if (RankMatrixDataDict[studentID][k1]["matrix_count"] != null)
                                                                        row["類別1排名母數" + subjectIndex] = RankMatrixDataDict[studentID][k1]["matrix_count"].ToString();

                                                                    // 五標PR填值
                                                                    foreach (string rItem in r2List)
                                                                    {
                                                                        if (RankMatrixDataDict[studentID][k1][rItem] != null)
                                                                        {
                                                                            if (r2ParseList.Contains(rItem))
                                                                            {
                                                                                decimal dd;
                                                                                if (decimal.TryParse(RankMatrixDataDict[studentID][k1][rItem].ToString(), out dd))
                                                                                {
                                                                                    row["類別1排名" + subjectIndex + "_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                row["類別1排名" + subjectIndex + "_" + rItem] = RankMatrixDataDict[studentID][k1][rItem].ToString();
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion
                                                            #region 類別2排名及落點分析
                                                            if (RankMatrixDataDict.ContainsKey(studentID))
                                                            {
                                                                k1 = "定期評量/科目成績_" + sceTakeRecord.Subject + "_類別2排名";
                                                                if (RankMatrixDataDict[studentID].ContainsKey(k1))
                                                                {
                                                                    if (RankMatrixDataDict[studentID][k1]["rank"] != null)
                                                                        row["類別2排名" + subjectIndex] = RankMatrixDataDict[studentID][k1]["rank"].ToString();

                                                                    if (RankMatrixDataDict[studentID][k1]["matrix_count"] != null)
                                                                        row["類別2排名母數" + subjectIndex] = RankMatrixDataDict[studentID][k1]["matrix_count"].ToString();

                                                                    // 五標PR填值
                                                                    foreach (string rItem in r2List)
                                                                    {
                                                                        if (RankMatrixDataDict[studentID][k1][rItem] != null)
                                                                        {
                                                                            if (r2ParseList.Contains(rItem))
                                                                            {
                                                                                decimal dd;
                                                                                if (decimal.TryParse(RankMatrixDataDict[studentID][k1][rItem].ToString(), out dd))
                                                                                {
                                                                                    row["類別2排名" + subjectIndex + "_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                                                }
                                                                            }
                                                                            else
                                                                            {
                                                                                row["類別2排名" + subjectIndex + "_" + rItem] = RankMatrixDataDict[studentID][k1][rItem].ToString();
                                                                            }
                                                                        }
                                                                    }
                                                                }
                                                            }
                                                            #endregion

                                                        }
                                                    }
                                                    else
                                                    {//修課有該考試但沒有成績資料
                                                        var courseRecs = accessHelper.CourseHelper.GetCourse(courseID);
                                                        if (courseRecs.Count > 0)
                                                        {
                                                            var courseRec = courseRecs[0];
                                                            if (conf.PrintSubjectList.Contains(courseRec.Subject))
                                                            {
                                                                if (findInSemesterSubjectScore)
                                                                {
                                                                    if (courseRec.SubjectLevel != "" + subjectNumber)
                                                                    {
                                                                        continue;
                                                                    }
                                                                }
                                                                findInExamScores = true;
                                                                if (!findInSemesterSubjectScore)
                                                                {
                                                                    decimal level;
                                                                    subjectNumber = decimal.TryParse(courseRec.SubjectLevel, out level) ? (decimal?)level : null;
                                                                    row["領域" + subjectIndex] = courseRec.Domain;
                                                                    row["分項類別" + subjectIndex] = courseRec.Entry;
                                                                    row["科目名稱" + subjectIndex] = courseRec.Subject + GetNumber(subjectNumber);
                                                                    if (!conf.IsShowLevel)  //2021-12-27 Cynthia 不顯示級別
                                                                        row["科目名稱" + subjectIndex] = courseRec.Subject;
                                                                    row["科目" + subjectIndex] = courseRec.Subject;
                                                                    row["科目級別" + subjectIndex] = GetNumber(subjectNumber);
                                                                    string required = courseRec.Required ? "必修" : "選修";
                                                                    string requiredBy = courseRec.RequiredBy == "部訂" ? "部定" : courseRec.RequiredBy;
                                                                    row["科目必選修" + subjectIndex] = required;
                                                                    row["科目校部定/科目必選修" + subjectIndex] = "" + requiredBy[0] + required[0];
                                                                    row["學分數" + subjectIndex] = courseRec.CreditDec();
                                                                }
                                                                row["科目成績" + subjectIndex] = "未輸入";
                                                            }
                                                        }
                                                    }
                                                    if (studentRefExamSores.ContainsKey(studentID) && studentRefExamSores[studentID].ContainsKey(courseID))
                                                    {

                                                        if (studentRefExamSores[studentID][courseID].SpecialCase == "缺")
                                                        {
                                                            row["前次成績" + subjectIndex] = 0;
                                                        }
                                                        else
                                                        {
                                                            if (studentRefExamSores[studentID][courseID].ExamScore == -1)
                                                            {
                                                                row["前次成績" + subjectIndex] = 0;
                                                            }
                                                            else if (studentRefExamSores[studentID][courseID].ExamScore == -2)
                                                            {
                                                                row["前次成績" + subjectIndex] = "免";
                                                            }
                                                            else
                                                            {
                                                                row["前次成績" + subjectIndex] = studentRefExamSores[studentID][courseID].ExamScore;
                                                            }
                                                        }
                                                        //row["前次成績" + subjectIndex] =
                                                        //    studentRefExamSores[studentID][courseID].SpecialCase == ""
                                                        //    ? ("" + studentRefExamSores[studentID][courseID].ExamScore)
                                                        //    : studentRefExamSores[studentID][courseID].SpecialCase;

                                                        // 判斷如果有缺考原因，使用缺考輸入文字與原因
                                                        string examUseTextKey = studentID + "_" + studentID + "_" + studentRefExamSores[studentID][courseID].ExamName;
                                                        if (StudSceTakeInfoDict.ContainsKey(examUseTextKey))
                                                        {
                                                            row["前次成績" + subjectIndex] = StudSceTakeInfoDict[examUseTextKey].UseText;
                                                            row["前次科目缺考原因" + subjectIndex] = StudSceTakeInfoDict[examUseTextKey].ReportValue;
                                                        }

                                                    }
                                                    studentExamSores[studentID][subjectName].Remove(courseID);
                                                    break;
                                                }
                                            }
                                        }
                                    }
                                    #endregion
                                    #region 上學期學期成績
                                    if (conf.Semester == "2" && conf.WithPrevSemesterScore)
                                    {
                                        foreach (var semesterSubjectScore in stuRec.SemesterSubjectScoreList)
                                        {
                                            string newSemesterSubjectName = semesterSubjectScore.Subject;
                                            if (semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱") != "")
                                                newSemesterSubjectName = semesterSubjectScore.Detail.GetAttribute("指定學年科目名稱");

                                            if (semesterSubjectScore.Detail.GetAttribute("不計學分") != "是"
                                                && newSemesterSubjectName == subjectName
                                                && semesterSubjectScore.Semester == 1
                                                && semesterSubjectScore.GradeYear == currentGradeYear)
                                            {
                                                findInSemester1SubjectScore = true;
                                                if (!findInSemesterSubjectScore
                                                    && !findInExamScores)
                                                {
                                                    decimal level;
                                                    subjectNumber = decimal.TryParse(semesterSubjectScore.Level, out level) ? (decimal?)level : null;
                                                    row["領域" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("領域");
                                                    row["分項類別" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("開課分項類別");
                                                    row["科目名稱" + subjectIndex] = newSemesterSubjectName + GetNumber(subjectNumber);
                                                    if (!conf.IsShowLevel)  //2021-12-27 Cynthia 不顯示級別
                                                        row["科目名稱" + subjectIndex] = newSemesterSubjectName;
                                                    row["科目" + subjectIndex] = newSemesterSubjectName;
                                                    row["科目級別" + subjectIndex] = GetNumber(subjectNumber);
                                                    row["學分數" + subjectIndex] = semesterSubjectScore.CreditDec();
                                                    string required = semesterSubjectScore.Require ? "必修" : "選修";
                                                    string requiredBy = semesterSubjectScore.Detail.GetAttribute("修課校部訂");
                                                    row["科目必選修" + subjectIndex] = required;
                                                    row["科目校部定/科目必選修" + subjectIndex] = "" + requiredBy[0] + required[0];
                                                    row["科目註記" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("註記");
                                                }
                                                row["上學期科目取得學分" + subjectIndex] = semesterSubjectScore.Pass ? "是" : "否";
                                                row["上學期科目未取得學分註記" + subjectIndex] = semesterSubjectScore.Pass ? "" : "\f";

                                                //"原始成績", "學年調整成績", "擇優採計成績", "補考成績", "重修成績"
                                                //if (semesterSubjectScore.Detail.GetAttribute("不需評分") != "是")
                                                {

                                                    if (semesterSubjectScore.Pass == false)
                                                    {
                                                        row["上學期科目未取得學分標示" + subjectIndex] = conf.FailScoreMark;
                                                    }

                                                    row["上學期科目原始成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("原始成績");
                                                    row["上學期科目補考成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("補考成績");
                                                    row["上學期科目重修成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("重修成績");
                                                    row["上學期科目手動調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("擇優採計成績");
                                                    row["上學期科目學年調整成績" + subjectIndex] = semesterSubjectScore.Detail.GetAttribute("學年調整成績");
                                                    row["上學期科目成績" + subjectIndex] = semesterSubjectScore.Score;

                                                    if ("" + semesterSubjectScore.Detail.GetAttribute("是否補修成績") == "是")
                                                    {
                                                        row["上學期科目補修成績標示" + subjectIndex] = conf.RepairScoreMark;
                                                    }

                                                    if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("原始成績"))
                                                        row["上學期科目原始成績註記" + subjectIndex] = "\f";
                                                    if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("補考成績"))
                                                    {
                                                        row["上學期科目補考成績註記" + subjectIndex] = "\f";
                                                        row["上學期科目補考成績標示" + subjectIndex] = conf.ReScoreMark;
                                                    }

                                                    if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("重修成績"))
                                                    {
                                                        row["上學期科目重修成績註記" + subjectIndex] = "\f";
                                                        row["上學期科目重修成績標示" + subjectIndex] = conf.RereadScoreMark;
                                                    }

                                                    if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("擇優採計成績"))
                                                        row["上學期科目手動成績註記" + subjectIndex] = "\f";
                                                    if ("" + semesterSubjectScore.Score == semesterSubjectScore.Detail.GetAttribute("學年調整成績"))
                                                        row["上學期科目學年成績註記" + subjectIndex] = "\f";

                                                    decimal scA = o_scA;
                                                    decimal scB = o_scB;
                                                    //string keySa = semesterSubjectScore.SchoolYear + "_" + semesterSubjectScore.Semester + "_" + semesterSubjectScore.Subject + "_" + semesterSubjectScore.Level + "_及";
                                                    //string keySb = semesterSubjectScore.SchoolYear + "_" + semesterSubjectScore.Semester + "_" + semesterSubjectScore.Subject + "_" + semesterSubjectScore.Level + "_補";

                                                    //if (StudentSCAttendApplyLimitDict.ContainsKey(stuRec.StudentID))
                                                    //{
                                                    //    if (StudentSCAttendApplyLimitDict[stuRec.StudentID].ContainsKey(keySa))
                                                    //        scA = StudentSCAttendApplyLimitDict[stuRec.StudentID][keySa];

                                                    //    if (StudentSCAttendApplyLimitDict[stuRec.StudentID].ContainsKey(keySb))
                                                    //        scB = StudentSCAttendApplyLimitDict[stuRec.StudentID][keySb];
                                                    //}

                                                    decimal.TryParse(semesterSubjectScore.Detail.GetAttribute("修課及格標準"), out scA);
                                                    decimal.TryParse(semesterSubjectScore.Detail.GetAttribute("修課補考標準"), out scB);

                                                    // 不及格
                                                    if (semesterSubjectScore.Score < scA)
                                                    {

                                                        // 可補考
                                                        if (semesterSubjectScore.Score >= scB)
                                                        {
                                                            if (semesterSubjectScore.Detail.GetAttribute("補考成績") == "" && semesterSubjectScore.Detail.GetAttribute("重修成績") == "" && semesterSubjectScore.Pass == false)
                                                            {
                                                                row["上學期科目需要補考註記" + subjectIndex] = "\f";
                                                                row["上學期科目需要補考標示" + subjectIndex] = conf.NeedReScoreMark;
                                                            }

                                                        }

                                                        decimal ss = 0;
                                                        decimal.TryParse(semesterSubjectScore.Detail.GetAttribute("原始成績"), out ss);


                                                        // 不可補考需要重修，沒有取得學分才需要重修
                                                        if ((ss < scB || semesterSubjectScore.Detail.GetAttribute("補考成績") != "" || semesterSubjectScore.Detail.GetAttribute("重修成績") != "") && semesterSubjectScore.Pass == false)
                                                        {
                                                            row["上學期科目需要補考標示" + subjectIndex] = "";
                                                            row["上學期科目需要重修註記" + subjectIndex] = "\f";
                                                            row["上學期科目需要重修標示" + subjectIndex] = conf.NeedRereadScoreMark;

                                                        }

                                                    }
                                                }
                                                stuRec.SemesterSubjectScoreList.Remove(semesterSubjectScore);
                                                break;
                                            }
                                        }
                                    }
                                    #endregion
                                    #region 學年成績
                                    if (conf.Semester == "2" && conf.WithSchoolYearScore)
                                    {
                                        foreach (var schoolYearSubjectScore in stuRec.SchoolYearSubjectScoreList)
                                        {
                                            if (("" + schoolYearSubjectScore.SchoolYear) == conf.SchoolYear
                                            && schoolYearSubjectScore.Subject == subjectName)
                                            {
                                                if (!findInSemesterSubjectScore
                                                    && !findInSemester1SubjectScore
                                                    && !findInExamScores)
                                                {
                                                    row["科目名稱" + subjectIndex] = schoolYearSubjectScore.Subject;
                                                }
                                                row["學年科目成績" + subjectIndex] = schoolYearSubjectScore.Score;
                                                stuRec.SchoolYearSubjectScoreList.Remove(schoolYearSubjectScore);
                                                break;
                                            }
                                        }
                                    }
                                    #endregion
                                    subjectIndex++;
                                }
                                else
                                {
                                    //重要!!發現資料在樣板中印不下時一定要記錄起來，否則使用者自己不會去發現的
                                    if (!overflowRecords.Contains(stuRec))
                                        overflowRecords.Add(stuRec);
                                }
                            }
                            #endregion

                            #region 總分
                            if (studentPrintSubjectSum.ContainsKey(studentID))
                            {
                                row["總分"] = studentPrintSubjectSum[studentID];

                                if (RankMatrixDataDict.ContainsKey(studentID))
                                {
                                    string skey = "定期評量/總計成績_總分_班排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["總分班排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["總分班排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["總分班排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["總分班排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }
                                    }

                                    skey = "定期評量/總計成績_總分_科排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["總分科排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["總分科排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["總分科排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["總分科排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }
                                    }

                                    skey = "定期評量/總計成績_總分_年排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["總分全校排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["總分全校排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["總分全校排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["總分全校排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                            #endregion
                            #region 平均
                            if (studentPrintSubjectAvg.ContainsKey(studentID))
                            {
                                row["平均"] = studentPrintSubjectAvg[studentID];
                                if (RankMatrixDataDict.ContainsKey(studentID))
                                {
                                    string skey = "定期評量/總計成績_平均_班排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["平均班排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["平均班排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["平均班排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["平均班排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }

                                    }

                                    skey = "定期評量/總計成績_平均_科排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["平均科排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["平均科排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["平均科排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["平均科排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }
                                    }

                                    skey = "定期評量/總計成績_平均_年排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["平均全校排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["平均全校排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["平均全校排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["平均全校排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }

                                    }
                                }

                            }
                            #endregion
                            #region 加權總分
                            if (studentPrintSubjectSumW.ContainsKey(studentID))
                            {
                                row["加權總分"] = studentPrintSubjectSumW[studentID];
                                if (RankMatrixDataDict.ContainsKey(studentID))
                                {
                                    string skey = "定期評量/總計成績_加權總分_班排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權總分班排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權總分班排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["加權總分班排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["加權總分班排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }
                                    }

                                    skey = "定期評量/總計成績_加權總分_科排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權總分科排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權總分科排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["加權總分科排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["加權總分科排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }
                                    }

                                    skey = "定期評量/總計成績_加權總分_年排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權總分全校排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權總分全校排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["加權總分全校排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["加權總分全校排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }

                                    }
                                }
                            }
                            #endregion
                            #region 加權平均
                            if (studentPrintSubjectAvgW.ContainsKey(studentID))
                            {
                                row["加權平均"] = studentPrintSubjectAvgW[studentID];
                                if (RankMatrixDataDict.ContainsKey(studentID))
                                {

                                    string skey = "定期評量/總計成績_加權平均_班排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權平均班排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權平均班排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["加權平均班排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["加權平均班排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }
                                    }

                                    skey = "定期評量/總計成績_加權平均_科排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權平均科排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權平均科排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["加權平均科排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["加權平均科排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }

                                    }

                                    skey = "定期評量/總計成績_加權平均_年排名";
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["加權平均全校排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["加權平均全校排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();
                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["加權平均全校排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["加權平均全校排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }
                                    }

                                }
                            }
                            #endregion
                            #region 類別1綜合成績

                            // 評量成績排名、五標、組距
                            if (RankMatrixDataDict.ContainsKey(studentID))
                            {
                                string skey = "定期評量/總計成績_總分_類別1排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別1總分排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別1總分排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                        {
                                            if (r2ParseList.Contains(rItem))
                                            {
                                                decimal dd;
                                                if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                {
                                                    row["類別1總分排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                }
                                            }
                                            else
                                            {
                                                row["類別1總分排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                            }
                                        }
                                    }

                                }

                                skey = "定期評量/總計成績_平均_類別1排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別1平均排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別1平均排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                        {
                                            if (r2ParseList.Contains(rItem))
                                            {
                                                decimal dd;
                                                if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                {
                                                    row["類別1平均排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                }
                                            }
                                            else
                                            {
                                                row["類別1平均排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                            }
                                        }
                                    }
                                }

                                skey = "定期評量/總計成績_加權總分_類別1排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別1加權總分排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別1加權總分排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                        {
                                            if (r2ParseList.Contains(rItem))
                                            {
                                                decimal dd;
                                                if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                {
                                                    row["類別1加權總分排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                }
                                            }
                                            else
                                            {
                                                row["類別1加權總分排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                            }
                                        }
                                    }
                                }

                                skey = "定期評量/總計成績_加權平均_類別1排名";
                                if (RankMatrixDataDict.ContainsKey(studentID))
                                {
                                    if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (RankMatrixDataDict[studentID][skey]["rank_name"] != null)
                                        {
                                            row["類別排名1"] = RankMatrixDataDict[studentID][skey]["rank_name"].ToString();
                                        }

                                        if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["類別1加權平均排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["類別1加權平均排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標PR填值
                                        foreach (string rItem in r2List)
                                        {
                                            if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                            {
                                                if (r2ParseList.Contains(rItem))
                                                {
                                                    decimal dd;
                                                    if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                    {
                                                        row["類別1加權平均排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    }
                                                }
                                                else
                                                {
                                                    row["類別1加權平均排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                                }
                                            }
                                        }

                                    }
                                }
                            }


                            #endregion
                            #region 類別2綜合成績

                            if (RankMatrixDataDict.ContainsKey(studentID))
                            {
                                string skey = "定期評量/總計成績_總分_類別2排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別2總分排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別2總分排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                        {
                                            if (r2ParseList.Contains(rItem))
                                            {
                                                decimal dd;
                                                if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                {
                                                    row["類別2總分排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                }
                                            }
                                            else
                                            {
                                                row["類別2總分排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                            }
                                        }
                                    }
                                }

                                skey = "定期評量/總計成績_平均_類別2排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別2平均排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別2平均排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                        {
                                            if (r2ParseList.Contains(rItem))
                                            {
                                                decimal dd;
                                                if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                {
                                                    row["類別2平均排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                }
                                            }
                                            else
                                            {
                                                row["類別2平均排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                            }
                                        }
                                    }
                                }

                                skey = "定期評量/總計成績_加權總分_類別2排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {
                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別2加權總分排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別2加權總分排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                        {
                                            if (r2ParseList.Contains(rItem))
                                            {
                                                decimal dd;
                                                if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                {
                                                    row["類別2加權總分排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                }
                                            }
                                            else
                                            {
                                                row["類別2加權總分排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                            }
                                        }
                                    }
                                }

                                skey = "定期評量/總計成績_加權平均_類別2排名";
                                if (RankMatrixDataDict[studentID].ContainsKey(skey))
                                {

                                    if (RankMatrixDataDict[studentID][skey]["rank_name"] != null)
                                    {
                                        row["類別排名2"] = RankMatrixDataDict[studentID][skey]["rank_name"].ToString();
                                    }

                                    if (RankMatrixDataDict[studentID][skey]["rank"] != null)
                                        row["類別2加權平均排名"] = RankMatrixDataDict[studentID][skey]["rank"].ToString();

                                    if (RankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                        row["類別2加權平均排名母數"] = RankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                    // 五標PR填值
                                    foreach (string rItem in r2List)
                                    {
                                        if (RankMatrixDataDict[studentID][skey][rItem] != null)
                                        {
                                            if (r2ParseList.Contains(rItem))
                                            {
                                                decimal dd;
                                                if (decimal.TryParse(RankMatrixDataDict[studentID][skey][rItem].ToString(), out dd))
                                                {
                                                    row["類別2加權平均排名_" + rItem] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                }
                                            }
                                            else
                                            {
                                                row["類別2加權平均排名_" + rItem] = RankMatrixDataDict[studentID][skey][rItem].ToString();
                                            }
                                        }
                                    }
                                }
                            }

                            #endregion



                            #region 學期分項成績


                            // 學期分項成績排名、五標、組距
                            if (SemsScoreRankMatrixDataDict.ContainsKey(studentID))
                            {
                                // 分項
                                foreach (string sname in SemsItemNameList)
                                {
                                    string skey = "";

                                    #region 班排名
                                    skey = "學期/分項成績_" + sname + "_班排名";
                                    if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(skey))
                                    {
                                        if (SemsScoreRankMatrixDataDict[studentID][skey]["rank"] != null)
                                            row["學期" + sname + "成績班排名"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank"].ToString();

                                        if (SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                            row["學期" + sname + "成績班排名母數"] = SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                        // 五標、組距
                                        foreach (string item2 in r2List)
                                        {
                                            if (SemsScoreRankMatrixDataDict[studentID][skey][item2] != null)
                                            {
                                                //if (r2ParseList.Contains(item2))
                                                //{
                                                //    decimal dd;
                                                //    if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString(), out dd))
                                                //    {

                                                //        row["學期" + sname + "成績班排名_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                //    }
                                                //    else
                                                //    {
                                                //        row["學期" + sname + "成績班排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                //    }
                                                //}
                                                if (r2ParseList.Contains(item2))
                                                    row["學期" + sname + "成績班排名_" + item2] = ParseScore(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString());
                                                else
                                                    row["學期" + sname + "成績班排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                            }
                                        }

                                        skey = "學期/分項成績(原始)_" + sname + "_班排名";
                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(skey))
                                        {
                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["rank"] != null)
                                                row["學期" + sname + "(原始)成績班排名"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank"].ToString();

                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                                row["學期" + sname + "(原始)成績班排名母數"] = SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                            // 五標、組距
                                            foreach (string item2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey][item2] != null)
                                                {
                                                    //if (r2ParseList.Contains(item2))
                                                    //{
                                                    //    decimal dd;
                                                    //    if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString(), out dd))
                                                    //    {

                                                    //        row["學期" + sname + "(原始)成績班排名_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        row["學期" + sname + "(原始)成績班排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                    //    }
                                                    //}
                                                    if (r2ParseList.Contains(item2))
                                                        row["學期" + sname + "(原始)成績班排名_" + item2] = ParseScore(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString());
                                                    else
                                                        row["學期" + sname + "(原始)成績班排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                }
                                            }
                                        }
                                        #endregion


                                        #region 科排名
                                        skey = "學期/分項成績_" + sname + "_科排名";
                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(skey))
                                        {
                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["rank"] != null)
                                                row["學期" + sname + "成績科排名"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank"].ToString();

                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                                row["學期" + sname + "成績科排名母數"] = SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                            // 五標、組距
                                            foreach (string item2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey][item2] != null)
                                                {
                                                    //row["學期" + sname + "成績科排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                    if (r2ParseList.Contains(item2))
                                                        row["學期" + sname + "成績科排名_" + item2] = ParseScore(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString());
                                                    else
                                                        row["學期" + sname + "成績科排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                }
                                            }

                                        }

                                        skey = "學期/分項成績(原始)_" + sname + "_科排名";
                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(skey))
                                        {
                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["rank"] != null)
                                                row["學期" + sname + "(原始)成績科排名"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank"].ToString();

                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                                row["學期" + sname + "(原始)成績科排名母數"] = SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"].ToString();

                                            // 五標、組距
                                            foreach (string item2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey][item2] != null)
                                                {
                                                    //if (r2ParseList.Contains(item2))
                                                    //{
                                                    //    decimal dd;
                                                    //    if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString(), out dd))
                                                    //    {

                                                    //        row["學期" + sname + "(原始)成績科排名_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        row["學期" + sname + "(原始)成績科排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                    //    }
                                                    //}
                                                    if (r2ParseList.Contains(item2))
                                                        row["學期" + sname + "(原始)成績科排名_" + item2] = ParseScore(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString());
                                                    else
                                                        row["學期" + sname + "(原始)成績科排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                }
                                            }
                                        }
                                        #endregion

                                        #region 年排名
                                        skey = "學期/分項成績_" + sname + "_年排名";
                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(skey))
                                        {
                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["rank"] != null)
                                                row["學期" + sname + "成績全校排名"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank"].ToString();

                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                                row["學期" + sname + "成績全校排名母數"] = SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"].ToString();
                                            // 五標、組距
                                            foreach (string item2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey][item2] != null)
                                                {
                                                    //if (r2ParseList.Contains(item2))
                                                    //{
                                                    //    decimal dd;
                                                    //    if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString(), out dd))
                                                    //    {

                                                    //        row["學期" + sname + "成績全校排名_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        row["學期" + sname + "成績全校排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                    //    }
                                                    //}
                                                    if (r2ParseList.Contains(item2))
                                                        row["學期" + sname + "成績全校排名_" + item2] = ParseScore(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString());
                                                    else
                                                        row["學期" + sname + "成績全校排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                }
                                            }
                                        }

                                        skey = "學期/分項成績(原始)_" + sname + "_年排名";
                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(skey))
                                        {
                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["rank"] != null)
                                                row["學期" + sname + "(原始)成績全校排名"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank"].ToString();

                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                                row["學期" + sname + "(原始)成績全校排名母數"] = SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"].ToString();
                                            // 五標、組距
                                            foreach (string item2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey][item2] != null)
                                                {
                                                    //if (r2ParseList.Contains(item2))
                                                    //{
                                                    //    decimal dd;
                                                    //    if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString(), out dd))
                                                    //    {

                                                    //        row["學期" + sname + "(原始)成績全校排名_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        row["學期" + sname + "(原始)成績全校排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                    //    }
                                                    //}
                                                    if (r2ParseList.Contains(item2))
                                                        row["學期" + sname + "(原始)成績全校排名_" + item2] = ParseScore(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString());
                                                    else
                                                        row["學期" + sname + "(原始)成績全校排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                }
                                            }
                                        }
                                        #endregion

                                        #region 類1排名
                                        skey = "學期/分項成績_" + sname + "_類別1排名";
                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(skey))
                                        {
                                            if (sname == "學業")
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey]["rank_name"] != null)
                                                {
                                                    row["學期類別排名1"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank_name"].ToString();
                                                }
                                            }


                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["rank"] != null)
                                                row["學期" + sname + "成績類別1排名"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank"].ToString();

                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                                row["學期" + sname + "成績類別1排名母數"] = SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"].ToString();
                                            // 五標、組距
                                            foreach (string item2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey][item2] != null)
                                                {
                                                    //if (r2ParseList.Contains(item2))
                                                    //{
                                                    //    decimal dd;
                                                    //    if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString(), out dd))
                                                    //    {

                                                    //        row["學期" + sname + "成績類別1排名_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        row["學期" + sname + "成績類別1排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                    //    }
                                                    //}

                                                    if (r2ParseList.Contains(item2))
                                                        row["學期" + sname + "成績類別1排名_" + item2] = ParseScore(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString());
                                                    else
                                                        row["學期" + sname + "成績類別1排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                }
                                            }

                                        }

                                        skey = "學期/分項成績(原始)_" + sname + "_類別1排名";
                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(skey))
                                        {
                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["rank"] != null)
                                                row["學期" + sname + "(原始)成績類別1排名"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank"].ToString();

                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                                row["學期" + sname + "(原始)成績類別1排名母數"] = SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"].ToString();
                                            // 五標、組距
                                            foreach (string item2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey][item2] != null)
                                                {
                                                    //if (r2ParseList.Contains(item2))
                                                    //{
                                                    //    decimal dd;
                                                    //    if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString(), out dd))
                                                    //    {

                                                    //        row["學期" + sname + "(原始)成績類別1排名_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        row["學期" + sname + "(原始)成績類別1排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                    //    }
                                                    //}

                                                    if (r2ParseList.Contains(item2))
                                                        row["學期" + sname + "(原始)成績類別1排名_" + item2] = ParseScore(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString());
                                                    else
                                                        row["學期" + sname + "(原始)成績類別1排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();

                                                }
                                            }
                                        }
                                        #endregion

                                        #region 類2排名
                                        skey = "學期/分項成績_" + sname + "_類別2排名";
                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(skey))
                                        {
                                            if (sname == "學業")
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey]["rank_name"] != null)
                                                {
                                                    row["學期類別排名2"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank_name"].ToString();
                                                }
                                            }

                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["rank"] != null)
                                                row["學期" + sname + "成績類別2排名"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank"].ToString();

                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                                row["學期" + sname + "成績類別2排名母數"] = SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"].ToString();
                                            // 五標、組距
                                            foreach (string item2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey][item2] != null)
                                                {
                                                    //if (r2ParseList.Contains(item2))
                                                    //{
                                                    //    decimal dd;
                                                    //    if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString(), out dd))
                                                    //    {

                                                    //        row["學期" + sname + "成績類別2排名_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        row["學期" + sname + "成績類別2排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                    //    }
                                                    //}
                                                    if (r2ParseList.Contains(item2))
                                                        row["學期" + sname + "成績類別2排名_" + item2] = ParseScore(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString());
                                                    else
                                                        row["學期" + sname + "成績類別2排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                }
                                            }

                                        }

                                        skey = "學期/分項成績(原始)_" + sname + "_類別2排名";
                                        if (SemsScoreRankMatrixDataDict[studentID].ContainsKey(skey))
                                        {
                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["rank"] != null)
                                                row["學期" + sname + "(原始)成績類別2排名"] = SemsScoreRankMatrixDataDict[studentID][skey]["rank"].ToString();

                                            if (SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"] != null)
                                                row["學期" + sname + "(原始)成績類別2排名母數"] = SemsScoreRankMatrixDataDict[studentID][skey]["matrix_count"].ToString();
                                            // 五標、組距
                                            foreach (string item2 in r2List)
                                            {
                                                if (SemsScoreRankMatrixDataDict[studentID][skey][item2] != null)
                                                {
                                                    //if (r2ParseList.Contains(item2))
                                                    //{
                                                    //    decimal dd;
                                                    //    if (decimal.TryParse(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString(), out dd))
                                                    //    {

                                                    //        row["學期" + sname + "(原始)成績類別2排名_" + item2] = Math.Round(dd, r2ParseN, MidpointRounding.AwayFromZero);
                                                    //    }
                                                    //    else
                                                    //    {
                                                    //        row["學期" + sname + "(原始)成績類別2排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();
                                                    //    }
                                                    //}
                                                    if (r2ParseList.Contains(item2))
                                                        row["學期" + sname + "(原始)成績類別2排名_" + item2] = ParseScore(SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString());
                                                    else
                                                        row["學期" + sname + "(原始)成績類別2排名_" + item2] = SemsScoreRankMatrixDataDict[studentID][skey][item2].ToString();

                                                }
                                            }
                                        }
                                        #endregion

                                    }

                                }
                            }
                            #endregion

                            #endregion

                            #region 學務資料
                            #region 綜合表現
                            List<string> commentList = new List<string>();
                            List<string> faceList = new List<string>();
                            foreach (SemesterMoralScoreInfo info in stuRec.SemesterMoralScoreList)
                            {
                                // 處理上學期文字評量
                                if (conf.Semester == "2" && ("" + info.SchoolYear) == conf.SchoolYear && info.Semester == 1)
                                {
                                    faceList.Clear();
                                    commentList.Clear();

                                    row["上學期導師評語"] = info.SupervisedByComment;
                                    System.Xml.XmlElement xml = info.Detail;
                                    foreach (System.Xml.XmlElement each in xml.SelectNodes("TextScore/Morality"))
                                    {
                                        string face = each.GetAttribute("Face");
                                        if ((SmartSchool.Customization.Data.SystemInformation.Fields["文字評量對照表"] as System.Xml.XmlElement).SelectSingleNode("Content/Morality[@Face='" + face + "']") != null)
                                        {
                                            string comment = each.InnerText;
                                            row["綜合表現：" + face] = each.InnerText;

                                            // 2021-12-28 Cynthia 避免舊的變數不能使用，增加新版文字評量變數，舊版綜合表現OOO不移除
                                            commentList.Add(comment);
                                            faceList.Add(face);
                                            for (int i = 1; i <= faceList.Count; i++)
                                            {
                                                if (faceList.Count <= 10)
                                                    row["上學期文字評量名稱" + i] = faceList[i - 1];
                                            }
                                            for (int i = 1; i <= commentList.Count; i++)
                                            {
                                                if (commentList.Count <= 10)
                                                    row["上學期文字評量" + i] = commentList[i - 1];
                                            }
                                        }
                                    }

                                }

                                if (("" + info.Semester) == conf.Semester && ("" + info.SchoolYear) == conf.SchoolYear)
                                {
                                    faceList.Clear();
                                    commentList.Clear();

                                    row["導師評語"] = info.SupervisedByComment;
                                    System.Xml.XmlElement xml = info.Detail;
                                    foreach (System.Xml.XmlElement each in xml.SelectNodes("TextScore/Morality"))
                                    {
                                        string face = each.GetAttribute("Face");
                                        if ((SmartSchool.Customization.Data.SystemInformation.Fields["文字評量對照表"] as System.Xml.XmlElement).SelectSingleNode("Content/Morality[@Face='" + face + "']") != null)
                                        {
                                            string comment = each.InnerText;
                                            row["綜合表現：" + face] = each.InnerText;

                                            // 2021-12-28 Cynthia 避免舊的變數不能使用，增加新版文字評量變數，舊版綜合表現OOO不移除
                                            commentList.Add(comment);
                                            faceList.Add(face);
                                            for (int i = 1; i <= faceList.Count; i++)
                                            {
                                                if (faceList.Count <= 10)
                                                    row["文字評量名稱" + i] = faceList[i - 1];
                                            }
                                            for (int i = 1; i <= commentList.Count; i++)
                                            {
                                                if (commentList.Count <= 10)
                                                    row["文字評量" + i] = commentList[i - 1];
                                            }
                                        }
                                    }
                                    // break;
                                }
                            }
                            #endregion
                            #region 獎懲統計
                            int 大功 = 0;
                            int 小功 = 0;
                            int 嘉獎 = 0;
                            int 大過 = 0;
                            int 小過 = 0;
                            int 警告 = 0;
                            bool 留校察看 = false;

                            int previous大功 = 0;
                            int previous小功 = 0;
                            int previous嘉獎 = 0;
                            int previous大過 = 0;
                            int previous小過 = 0;
                            int previous警告 = 0;
                            bool previous留校察看 = false;
                            foreach (RewardInfo info in stuRec.RewardList)
                            {
                                if (("" + info.Semester) == conf.Semester && ("" + info.SchoolYear) == conf.SchoolYear)
                                {
                                    大功 += info.AwardA;
                                    小功 += info.AwardB;
                                    嘉獎 += info.AwardC;
                                    if (!info.Cleared)
                                    {
                                        大過 += info.FaultA;
                                        小過 += info.FaultB;
                                        警告 += info.FaultC;
                                    }
                                    if (info.UltimateAdmonition)
                                        留校察看 = true;
                                }

                                ////2021-12-27 Cynthia  新增上學期獎懲統計
                                if (conf.Semester == "2")
                                {
                                    if (("" + info.Semester) == "1" && ("" + info.SchoolYear) == conf.SchoolYear)
                                    {
                                        previous大功 += info.AwardA;
                                        previous小功 += info.AwardB;
                                        previous嘉獎 += info.AwardC;
                                        if (!info.Cleared)
                                        {
                                            previous大過 += info.FaultA;
                                            previous小過 += info.FaultB;
                                            previous警告 += info.FaultC;
                                        }
                                        if (info.UltimateAdmonition)
                                            previous留校察看 = true;
                                    }
                                }

                            }
                            // 本學期
                            row["大功統計"] = 大功 == 0 ? "0" : ("" + 大功);
                            row["小功統計"] = 小功 == 0 ? "0" : ("" + 小功);
                            row["嘉獎統計"] = 嘉獎 == 0 ? "0" : ("" + 嘉獎);
                            row["大過統計"] = 大過 == 0 ? "0" : ("" + 大過);
                            row["小過統計"] = 小過 == 0 ? "0" : ("" + 小過);
                            row["警告統計"] = 警告 == 0 ? "0" : ("" + 警告);
                            row["留校察看"] = 留校察看 ? "是" : "否";

                            //2021-12-27 Cynthia 上學期
                            row["上學期大功統計"] = previous大功 == 0 ? "0" : ("" + previous大功);
                            row["上學期小功統計"] = previous小功 == 0 ? "0" : ("" + previous小功);
                            row["上學期嘉獎統計"] = previous嘉獎 == 0 ? "0" : ("" + previous嘉獎);
                            row["上學期大過統計"] = previous大過 == 0 ? "0" : ("" + previous大過);
                            row["上學期小過統計"] = previous小過 == 0 ? "0" : ("" + previous小過);
                            row["上學期警告統計"] = previous警告 == 0 ? "0" : ("" + previous警告);
                            row["上學期留校察看"] = previous留校察看 ? "是" : "否";
                            #endregion

                            #region 缺曠統計
                            Dictionary<string, int> 缺曠項目統計 = new Dictionary<string, int>();
                            foreach (AttendanceInfo info in stuRec.AttendanceList)
                            {
                                if (("" + info.Semester) == conf.Semester && ("" + info.SchoolYear) == conf.SchoolYear)
                                {
                                    string infoType = "";
                                    if (dicPeriodMappingType.ContainsKey(info.Period))
                                        infoType = dicPeriodMappingType[info.Period];
                                    else
                                        infoType = "";
                                    string attendanceKey = "" + infoType + "_" + info.Absence;
                                    if (!缺曠項目統計.ContainsKey(attendanceKey))
                                        缺曠項目統計.Add(attendanceKey, 0);
                                    缺曠項目統計[attendanceKey]++;
                                }
                            }
                            foreach (string attendanceKey in 缺曠項目統計.Keys)
                            {
                                row[attendanceKey] = 缺曠項目統計[attendanceKey] == 0 ? "0" : ("" + 缺曠項目統計[attendanceKey]);
                            }
                            #endregion


                            #endregion

                            // 處理人數統計
                            if (stuRec.RefClass != null)
                            {
                                row["班級人數"] = sci.GetClassStudentCount(stuRec.RefClass.ClassName);
                            }
                            if (stuRec.RefClass != null)
                            {
                                string deptName = stuRec.RefClass.Department;
                                if (stuRec.Department != "")
                                    deptName = stuRec.Department;

                                row["科別人數"] = sci.GetDeptStudentCount(gradeYear, deptName);
                            }
                            row["年級人數"] = sci.GetGradeStudentCount(gradeYear);

                            // 處理類別人數
                            List<string> prefixList = sci.GetStudentPrefixList(stuRec.StudentID);
                            if (prefixList.Count > 0)
                            {
                                foreach (string name in prefixList)
                                {
                                    string keyName = name + "人數";

                                    row[keyName] = sci.GetTagStudentCount(gradeYear, name);
                                }
                            }

                            // 處理完整類別人數
                            List<string> FullNameList = sci.GetStudentFullNameTagList(stuRec.StudentID);
                            if (FullNameList.Count > 0)
                            {
                                foreach (string name in FullNameList)
                                {
                                    string keyName = name + "人數";

                                    row[keyName] = sci.GetFullNameTagStudentCount(gradeYear, name);
                                }
                            }

                            #endregion

                            table.Rows.Add(row);
                            progressCount++;
                            bkw.ReportProgress(70 + progressCount * 20 / selectedStudents.Count);

                            table.TableName = "test";
                            table.WriteXml(Application.StartupPath + "\\debug.xml");
                        }
                        bkw.ReportProgress(90);
                        document = conf.Template;
                        document.MailMerge.Execute(table);
                        document.MailMerge.DeleteFields();

                    }
                    catch (Exception exception)
                    {
                        //  exc = exception
                        throw exception;
                    }
                };
                bkw.RunWorkerAsync();
            }
        }

        /// <summary>
        /// 四捨五入至使用者指定位數
        /// </summary>
        /// <param name="str"></param>
        /// <returns></returns>
        public static string ParseScore(string str)
        {
            string value = "";
            decimal dc;
            if (decimal.TryParse(str, out dc))
            {
                value = Math.Round(dc, r2ParseN, MidpointRounding.AwayFromZero).ToString();
            }

            return value;
        }

        /***********************************
 * 學生之學年學業成績符合下列各款規定之一者，准予升級：
 *   一、各科目學年成績均及格。
 *   二、學年成績不及格之各科目每週教學總節數，未超過修習各科目每週教學總節數二分之一，且不及格之科目成績無零分，而學年總平均成績及格。
 *   三、學年成績不及格之科目未超過修習全部科目二分之一，且其不及格之科目成績無零分，而學年總平均成績及格。
 ***********************************/
        public static bool RetainInTheSameGrade(string pStudentID, string schoolYear, int gradeYear, decimal entryScore, List<SchoolYearSubjectScoreInfo> schoolYearSubjectScoreInfos, Dictionary<string, decimal> subjectCreditDic)
        {
            decimal? totalCredit = 0;
            decimal? notPassTotalCredit = 0;
            int totalSubjectAmount = 0;
            int notPassTotalSubjectAmount = 0;
            bool hasZeroScoreSubject = false;

            foreach (SchoolYearSubjectScoreInfo ss in schoolYearSubjectScoreInfos)
            {
                if (ss.SchoolYear.ToString() == schoolYear && ss.GradeYear == gradeYear)
                {
                    string sKey = ss.Subject;

                    //  所有科目每週教學總節數(ischool 無課時，故取學分)
                    if (subjectCreditDic.ContainsKey(sKey))
                        totalCredit += subjectCreditDic[sKey];

                    //  所有科目數
                    totalSubjectAmount++;

                    if (ss.Score < GetPassingStandard(pStudentID, ss.GradeYear))
                    {
                        if (subjectCreditDic.ContainsKey(sKey))
                            notPassTotalCredit += subjectCreditDic[sKey];

                        notPassTotalSubjectAmount += 1;
                    }

                    if (ss.Score == 0)
                        hasZeroScoreSubject = true;
                }
            }

            //  各科目學年成績均及格：升級
            if (notPassTotalSubjectAmount == 0)
                return false;

            //  剛好 1/2：升級
            //  學年成績不及格之各科目每週教學總節數，未超過修習各科目每週教學總節數二分之一，且不及格之科目成績無零分，而學年總平均成績(系統學年學業成績)及格
            if (((notPassTotalCredit * 2) <= totalCredit) && (!hasZeroScoreSubject) && (entryScore >= (GetPassingStandard(pStudentID, gradeYear))))
                return false;

            //  剛好 1/2：升級
            //  學年成績不及格之科目未超過修習全部科目二分之一，且其不及格之科目成績無零分，而學年總平均成績(系統學年學業成績)及格
            if (((notPassTotalSubjectAmount * 2) <= totalSubjectAmount) && (!hasZeroScoreSubject) && (entryScore >= (GetPassingStandard(pStudentID, gradeYear))))
                return false;

            return true;
        }

        /****<學生類別 一年級及格標準=\"60\" 一年級補考標準=\"0\" 三年級及格標準=\"60\" 三年級補考標準=\"0\" 二年級及格標準=\"60\" 二年級補考標準=\"0\" 四年級及格標準=\"60\" 四年級補考標準=\"0\" 類別=\"預設\" />
 * 比對學生成績身份，若比對不到則及格標準使用「預設」
 ****/
        public static decimal GetPassingStandard(string pStudentID, int? GradeYear)
        {
            GradeYear = GradeYear.HasValue ? GradeYear.Value : 0;
            decimal passingStandard = 60;
            decimal tmpPassingStandard = 60;

            List<SHSchool.Data.SHStudentTagRecord> pStudentTag = SHSchool.Data.SHStudentTag.SelectByStudentID(pStudentID);
            if (pStudentTag == null)
                return passingStandard;

            Dictionary<string, SHSchool.Data.SHScoreCalcRuleRecord> _SHScoreCalcRule = SHSchool.Data.SHScoreCalcRule.SelectAll().ToDictionary(x => x.ID);

            XmlElement scoreCalcRule = SmartSchool.Evaluation.ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(pStudentID) == null ? null : SmartSchool.Evaluation.ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(pStudentID).ScoreCalcRuleElement;
            if (scoreCalcRule == null)
                return passingStandard;

            List<string> pStudentTagFullName = new List<string>();
            List<string> pStudentTagPrefix = new List<string>();
            foreach (SHSchool.Data.SHStudentTagRecord sr in pStudentTag)
            {
                pStudentTagFullName.Add(sr.FullName);
                pStudentTagPrefix.Add(sr.Prefix);
            }

            XDocument doc = XDocument.Parse("<root>" + scoreCalcRule.InnerXml + "</root>");

            if (doc.Document == null)
                return passingStandard;

            if (doc.Document.Element("root").Element("及格標準") == null)
                return passingStandard;

            if (doc.Document.Element("root").Element("及格標準").Elements("學生類別") == null)
                return passingStandard;

            foreach (XElement x in doc.Document.Element("root").Element("及格標準").Elements("學生類別"))
            {
                if (x.Attribute("類別") != null)
                {
                    if (pStudentTagFullName.Contains(x.Attribute("類別").Value) || pStudentTagPrefix.Contains(x.Attribute("類別").Value))
                    {
                        try
                        {
                            if (x.Attribute(NumericToChinese(GradeYear) + "年級及格標準") != null)
                            {
                                bool result = decimal.TryParse(x.Attribute(NumericToChinese(GradeYear) + "年級及格標準").Value, out tmpPassingStandard);
                                if (result)
                                {
                                    if (tmpPassingStandard < passingStandard)
                                    {
                                        passingStandard = tmpPassingStandard;
                                    }
                                }
                            }
                        }
                        catch
                        {
                            return passingStandard;
                        }
                    }
                }
            }

            try
            {
                var elem = doc.XPathSelectElement("root/及格標準/學生類別[@類別='預設']");
                bool result = decimal.TryParse(elem.Attribute(NumericToChinese(GradeYear) + "年級及格標準").Value, out tmpPassingStandard);
                if (result)
                {
                    if (passingStandard > tmpPassingStandard)
                        passingStandard = tmpPassingStandard;
                }
            }
            catch
            {
                return passingStandard;
            }

            return passingStandard;
        }

        public static char NumericToChinese(int? pNumber)
        {
            int qNumber = pNumber.HasValue ? pNumber.Value : 0;
            if (pNumber > 10 || pNumber < 0)
                throw new Exception("轉換的數字必須介於 0~10");

            char[] number = new char[] { '○', '一', '二', '三', '四', '五', '六', '七', '八', '九', '十' };

            return number[qNumber];
        }
    }
}
