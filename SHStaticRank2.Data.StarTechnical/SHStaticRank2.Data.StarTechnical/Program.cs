using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using SmartSchool.Customization.Data;
using System.IO;
using FISCA.DSAUtil;
using SHStaticRank2.Data;
using System.Data;

namespace SHStaticRank2.Data.StarTechnical
{
    public class Program
    {
        [FISCA.MainMethod]
        public static void Main()
        {
            //FISCA.Permission.Catalog cat = FISCA.Permission.RoleAclSource.Instance["教務作業"]["功能按鈕"];
            //cat.Add(new FISCA.Permission.RibbonFeature("SHSchool.SHStaticRank2.Data", "計算固定排名(測試版)"));

            var button = FISCA.Presentation.MotherForm.RibbonBarItems["教務作業", "批次作業/檢視"]["成績作業"]["計算固定排名"]["計算多學期成績固定排名(104學年度技職繁星)"];
            button.Enable = FISCA.Permission.UserAcl.Current["SHSchool.SHStaticRank2.Data"].Executable;
            button.Click += delegate
            {
                var conf = new StarTechnical();
                conf.ShowDialog();
                if (conf.DialogResult == System.Windows.Forms.DialogResult.OK)
                {
                    CalcMutilSemeSubjectRank.OneClassCompleted += delegate
                    {
                        #region delegate
                        if (string.IsNullOrWhiteSpace(conf.Configure.Rank1Tag))
                            return;
                        #region 取得設定值 , MapRecord
                        string rank1tag = conf.Configure.Rank1Tag.Replace("[", "").Replace("]", "");
                        int subjectLimit = Math.Max(conf.Configure.useSubjectPrintList.Count, conf.Configure.SubjectLimit);
                        string path = Path.Combine(System.Windows.Forms.Application.StartupPath + "\\Reports",
                                                   CalcMutilSemeSubjectRank.FolderName);
                        string path_prefix = "E_群別";
                        string path_suffix = "報名表.xls";

                        FISCA.UDT.AccessHelper _AccessHelper = new FISCA.UDT.AccessHelper();
                        Dictionary<string, MapRecord> dmr;
                        try
                        {
                            dmr = _AccessHelper.Select<MapRecord>()
                                .FindAll(delegate(MapRecord mr) { return !string.IsNullOrWhiteSpace(mr.student_tag); })
                                .ToDictionary(x => x.student_tag, x => x);
                        }
                        catch (Exception)
                        {
                            return;
                        }

                        int seq;
                        Dictionary<string, Aspose.Cells.Workbook> wbs = new Dictionary<string, Aspose.Cells.Workbook>();
                        Dictionary<string, Aspose.Cells.Workbook> wbCheck = new Dictionary<string, Aspose.Cells.Workbook>();
                        Aspose.Cells.Workbook wb;
                        Aspose.Cells.Worksheet ws;
                        string fulltag;
                        string group_code;
                        string full_doc_path;
                        #endregion
                        foreach (DataRow student in CalcMutilSemeSubjectRank._table.Rows)
                        {
                            fulltag = student.Table.Columns.Contains("類別一分類") ? rank1tag + ":" + student["類別一分類"] : "";
                            //只處理有設定群別的類別
                            if (dmr.ContainsKey(fulltag))
                            {
                                #region 處理上傳表
                                group_code = dmr[fulltag].code1;
                                full_doc_path = Path.Combine(path, path_prefix + group_code + path_suffix);
                                try
                                {
                                    #region 開啟xls檔案加入Dictionary,或新增
                                    if (wbs.ContainsKey(group_code))
                                        wb = wbs[group_code];
                                    else if (File.Exists(full_doc_path))
                                    {
                                        wb = new Aspose.Cells.Workbook();
                                        wb.Open(full_doc_path);
                                        wbs.Add(group_code, wb);
                                    }
                                    else
                                    {
                                        wb = new Aspose.Cells.Workbook();
                                        ws = wb.Worksheets[0];
                                        #region 報表header初始化
                                        ws.Cells[0, 0].PutValue("序號");
                                        ws.Cells[0, 1].PutValue("學號");
                                        ws.Cells[0, 2].PutValue("學生姓名");
                                        ws.Cells[0, 3].PutValue("群別代碼");
                                        ws.Cells[0, 4].PutValue("學制代碼");
                                        ws.Cells[0, 5].PutValue("科(組)、學程名稱");
                                        ws.Cells[0, 6].PutValue("班級名稱");
                                        ws.Cells[0, 7].PutValue("學業平均成績科(組)、學程名次");
                                        ws.Cells[0, 8].PutValue("學業平均成績群名次");
                                        ws.Cells[0, 9].PutValue("專業及實習科目平均成績群名次");
                                        ws.Cells[0, 10].PutValue("英文平均成績群名次");
                                        ws.Cells[0, 11].PutValue("國文平均成績群名次");
                                        ws.Cells[0, 12].PutValue("數學平均成績群名次");
                                        #endregion
                                        wbs.Add(group_code, wb);
                                    }
                                    #endregion
                                    ws = wb.Worksheets[0];
                                    seq = 1;//序號
                                    while (!string.IsNullOrWhiteSpace("" + ws.Cells[seq, 0].Value)) { seq++; }
                                    #region 填入資料
                                    ws.Cells[seq, 0].PutValue("" + seq);//1.序號
                                    ws.Cells[seq, 1].PutValue("" + student["學號"]);//2.學號
                                    ws.Cells[seq, 2].PutValue("" + student["姓名"]);//3.學生姓名
                                    ws.Cells[seq, 3].PutValue(group_code);//4.群別代碼
                                    ws.Cells[seq, 4].PutValue(dmr[fulltag].code2);//5.學制代碼
                                    ws.Cells[seq, 5].PutValue("" + student["科別"]);//6.科(組),學程名稱
                                    ws.Cells[seq, 6].PutValue("" + student["班級"]);//7.班級名稱
                                    ws.Cells[seq, 7].PutValue("" + student["學業原始平均科排名"]);//8.平均科排名
                                    ws.Cells[seq, 8].PutValue("" + student["學業原始平均類別一排名"]);//9.學業平均成績群名次
                                    ws.Cells[seq, 9].PutValue("" + student["篩選科目原始成績加權平均平均類別二排名"]);//10.專業及實習平均成績群名次
                                    for (int i = 1; i <= subjectLimit; i++)
                                    {
                                        switch ("" + student["科目名稱" + i])
                                        {
                                            case "英文":
                                                ws.Cells[seq, 10].PutValue("" + student["科目平均類別一排名" + i]);//11.英文平均成績群名次
                                                break;
                                            case "國文":
                                                ws.Cells[seq, 11].PutValue("" + student["科目平均類別一排名" + i]);//12.國文平均成績群名次
                                                break;
                                            case "數學":
                                                ws.Cells[seq, 12].PutValue("" + student["科目平均類別一排名" + i]);//13.數學平均群名次
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                    #endregion
                                    wb = null;
                                }
                                catch (Exception)
                                {
                                    #region save all opened excel
                                    foreach (KeyValuePair<string, Aspose.Cells.Workbook> item in wbs)
                                    {
                                        item.Value.Save(Path.Combine(path, path_prefix + item.Key + path_suffix), Aspose.Cells.FileFormatType.Excel97To2003);
                                    }
                                    #endregion
                                    throw;
                                }
                                #endregion
                                #region 處理驗算表
                                group_code = dmr[fulltag].code1;
                                full_doc_path = Path.Combine(Path.Combine(path, path_prefix + "成績明細" + group_code + ".xls"));
                                try
                                {
                                    #region 開啟xls檔案加入Dictionary,或新增
                                    if (wbCheck.ContainsKey(group_code))
                                        wb = wbCheck[group_code];
                                    else if (File.Exists(full_doc_path))
                                    {
                                        wb = new Aspose.Cells.Workbook();
                                        wb.Open(full_doc_path);
                                        wbCheck.Add(group_code, wb);
                                    }
                                    else
                                    {
                                        wb = new Aspose.Cells.Workbook();
                                        ws = wb.Worksheets[0];
                                        #region 報表header初始化
                                        ws.Cells[0, 0].PutValue("序號");
                                        ws.Cells[0, 1].PutValue("學號");
                                        ws.Cells[0, 2].PutValue("學生姓名");
                                        ws.Cells[0, 3].PutValue("群別代碼");
                                        ws.Cells[0, 4].PutValue("學制代碼");
                                        ws.Cells[0, 5].PutValue("科(組)、學程名稱");
                                        ws.Cells[0, 6].PutValue("班級名稱");

                                        ws.Cells[0, 7].PutValue("類別一分類");
                                        ws.Cells[0, 8].PutValue("類別二分類");

                                        ws.Cells[0, 9].PutValue("學業原始平均成績");
                                        ws.Cells[0, 10].PutValue("學業平均成績科(組)、學程名次");
                                        ws.Cells[0, 11].PutValue("學業平均成績群名次");

                                        ws.Cells[0, 12].PutValue("專業及實習科目平均成績");
                                        ws.Cells[0, 13].PutValue("專業及實習科目平均成績群名次");

                                        ws.Cells[0, 14].PutValue("英文科平均成績");
                                        ws.Cells[0, 15].PutValue("英文平均成績群名次");
                                        ws.Cells[0, 16].PutValue("國文科平均成績");
                                        ws.Cells[0, 17].PutValue("國文平均成績群名次");
                                        ws.Cells[0, 18].PutValue("數學科平均成績");
                                        ws.Cells[0, 19].PutValue("數學平均成績群名次");
                                        #endregion
                                        wbCheck.Add(group_code, wb);
                                    }
                                    #endregion
                                    ws = wb.Worksheets[0];
                                    seq = 1;//序號
                                    while (!string.IsNullOrWhiteSpace("" + ws.Cells[seq, 0].Value)) { seq++; }
                                    #region 填入資料
                                    ws.Cells[seq, 0].PutValue("" + seq);//1.序號
                                    ws.Cells[seq, 1].PutValue("" + student["學號"]);//2.學號
                                    ws.Cells[seq, 2].PutValue("" + student["姓名"]);//3.學生姓名
                                    ws.Cells[seq, 3].PutValue(group_code);//4.群別代碼
                                    ws.Cells[seq, 4].PutValue(dmr[fulltag].code2);//5.學制代碼
                                    ws.Cells[seq, 5].PutValue("" + student["科別"]);//6.科(組),學程名稱
                                    ws.Cells[seq, 6].PutValue("" + student["班級"]);//7.班級名稱

                                    ws.Cells[seq, 7].PutValue("" + student["類別一分類"]);
                                    ws.Cells[seq, 8].PutValue("" + student["類別二分類"]);

                                    ws.Cells[seq, 9].PutValue("" + student["學業原始平均"]);//學業原始平均成績
                                    ws.Cells[seq, 10].PutValue("" + student["學業原始平均科排名"]);//8.平均科排名
                                    ws.Cells[seq, 11].PutValue("" + student["學業原始平均類別一排名"]);//9.學業平均成績群名次

                                    ws.Cells[seq, 12].PutValue("" + student["篩選科目原始成績加權平均平均類別二"]);//專業及實習平均成績
                                    ws.Cells[seq, 13].PutValue("" + student["篩選科目原始成績加權平均平均類別二排名"]);//10.專業及實習平均成績群名次
                                    for (int i = 1; i <= subjectLimit; i++)
                                    {
                                        switch ("" + student["科目名稱" + i])
                                        {
                                            case "英文":
                                                ws.Cells[seq, 14].PutValue("" + student["科目平均" + i]);//英文平均成績
                                                ws.Cells[seq, 15].PutValue("" + student["科目平均類別一排名" + i]);//11.英文平均成績群名次
                                                break;
                                            case "國文":
                                                ws.Cells[seq, 16].PutValue("" + student["科目平均" + i]);//國文平均成績
                                                ws.Cells[seq, 17].PutValue("" + student["科目平均類別一排名" + i]);//12.國文平均成績群名次
                                                break;
                                            case "數學":
                                                ws.Cells[seq, 18].PutValue("" + student["科目平均" + i]);//數學平均成績
                                                ws.Cells[seq, 19].PutValue("" + student["科目平均類別一排名" + i]);//13.數學平均群名次
                                                break;
                                            default:
                                                break;
                                        }
                                    }
                                    #endregion
                                    wb = null;
                                }
                                catch (Exception)
                                {
                                    #region save all opened excel
                                    foreach (KeyValuePair<string, Aspose.Cells.Workbook> item in wbCheck)
                                    {
                                        item.Value.Save(Path.Combine(path, path_prefix + "成績明細" + item.Key + ".xls"), Aspose.Cells.FileFormatType.Excel97To2003);
                                    }
                                    #endregion
                                    throw;
                                }
                                #endregion
                            }
                        }
                        #region save all opened excel
                        foreach (KeyValuePair<string, Aspose.Cells.Workbook> item in wbs)
                        {
                            item.Value.Save(Path.Combine(path, path_prefix + item.Key + path_suffix), Aspose.Cells.FileFormatType.Excel97To2003);
                        }
                        #endregion
                        #region save all opened excel wbCheck
                        foreach (KeyValuePair<string, Aspose.Cells.Workbook> item in wbCheck)
                        {
                            item.Value.Save(Path.Combine(path, path_prefix + "成績明細" + item.Key + ".xls"), Aspose.Cells.FileFormatType.Excel97To2003);
                        }
                        #endregion
                        wbs = null;
                        #endregion
                    };
                    CalcMutilSemeSubjectRank.Setup(conf.Configure);
                }
            };
        }

    }
}
