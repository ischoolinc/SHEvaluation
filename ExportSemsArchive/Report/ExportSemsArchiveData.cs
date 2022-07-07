using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Aspose.Cells;
using FISCA.Data;
using FISCA.Permission;
using FISCA.Presentation;

namespace ExportSemsArchive.Report
{
    //[FeatureCode("ExportSemsArchiveData", "學期成績(封存)報表")]
    class ExportSemsArchiveData
    {
        List<string> studentIDList = new List<string>();
        BackgroundWorker _bgWorker;
        Workbook _wb;


        QueryHelper queryHelper = new QueryHelper();
        Aspose.Cells.Workbook wb = new Aspose.Cells.Workbook();
        int rowIndex = 1;

        public ExportSemsArchiveData(List<string> stuIDs)
        {
            studentIDList = stuIDs;
            _bgWorker = new BackgroundWorker();
            _bgWorker.DoWork += _bgWorker_DoWork;
            _bgWorker.ProgressChanged += _bgWorker_ProgressChanged;
            _bgWorker.RunWorkerCompleted += _bgWorker_RunWorkerCompleted;
            _bgWorker.WorkerReportsProgress = true;
        }

        private void _bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("學期成績(封存)報表 產生完成。");

            if (_wb != null)
            {
                Utility.ExprotXls("學期成績(封存)報表", _wb);
            }
        }

        private void _bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            MotherForm.SetStatusBarMessage("學期成績(封存)報表 產生中...", e.ProgressPercentage);
        }


        private void _bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            _bgWorker.ReportProgress(10);
            string sql = @"SELECT
											sems_subj_score_ext.ref_student_id AS 學生系統編號
											, student.student_number AS 學號
											, class.class_name AS 班級
											, student.seat_no AS 座號
											, CASE WHEN	student.ref_dept_id IS NULL THEN  g.name
														 WHEN student.ref_dept_id IS NOT NULL THEN h.name 	END AS  科別
											, student.name AS 姓名
										   , sems_subj_score_ext.last_update AS 封存日期
											, sems_subj_score_ext.school_year AS 學年度
											, sems_subj_score_ext.semester AS 學期
											, sems_subj_score_ext.grade_year AS 成績年級
											, array_to_string(xpath('//Subject/@開課分項類別', subj_score_ele), '')::text AS 分項類別
											, array_to_string(xpath('//Subject/@科目', subj_score_ele), '')::text AS 科目
											, array_to_string(xpath('//Subject/@科目級別', subj_score_ele), '')::text AS 科目級別
											, array_to_string(xpath('//Subject/@開課學分數', subj_score_ele), '')::text AS 學分數
											, array_to_string(xpath('//Subject/@修課校部訂', subj_score_ele), '')::text AS 校部訂
											, array_to_string(xpath('//Subject/@修課必選修', subj_score_ele), '')::text AS 必選修
											, array_to_string(xpath('//Subject/@是否取得學分', subj_score_ele), '')::text AS 取得學分
											, array_to_string(xpath('//Subject/@原始成績', subj_score_ele), '')::text AS 原始成績
											, array_to_string(xpath('//Subject/@補考成績', subj_score_ele), '')::text AS 補考成績
											, array_to_string(xpath('//Subject/@重修成績', subj_score_ele), '')::text AS 重修成績
											, array_to_string(xpath('//Subject/@擇優採計成績', subj_score_ele), '')::text AS 手動調整成績
											, array_to_string(xpath('//Subject/@學年調整成績', subj_score_ele), '')::text AS 學年調整成績
											, array_to_string(xpath('//Subject/@不計學分', subj_score_ele), '')::text AS 不計學分
											, array_to_string(xpath('//Subject/@不需評分', subj_score_ele), '')::text AS 不需評分
											, array_to_string(xpath('//Subject/@註記', subj_score_ele), '')::text AS 註記
											, array_to_string(xpath('//Subject/@修課及格標準', subj_score_ele), '')::text AS 修課及格標準
											, array_to_string(xpath('//Subject/@修課補考標準', subj_score_ele), '')::text AS 修課補考標準
											, array_to_string(xpath('//Subject/@修課直接指定總成績', subj_score_ele), '')::text AS 修課直接指定總成績
											, array_to_string(xpath('//Subject/@修課備註', subj_score_ele), '')::text AS 修課備註
											, array_to_string(xpath('//Subject/@修課科目代碼', subj_score_ele), '')::text AS 修課科目代碼	
											, array_to_string(xpath('//Subject/@是否補修成績', subj_score_ele), '')::text AS 是否補修成績
											, array_to_string(xpath('//Subject/@重修學年度', subj_score_ele), '')::text AS 重修學年度
											, array_to_string(xpath('//Subject/@重修學期', subj_score_ele), '')::text AS 重修學期 
											, array_to_string(xpath('//Subject/@補修學年度', subj_score_ele), '')::text AS 補修學年度
											, array_to_string(xpath('//Subject/@補修學期', subj_score_ele), '')::text AS 補修學期 
											, array_to_string(xpath('//Subject/@免修', subj_score_ele), '')::text AS 免修
											, array_to_string(xpath('//Subject/@抵免', subj_score_ele), '')::text AS 抵免
											, array_to_string(xpath('//Subject/@指定學年科目名稱', subj_score_ele), '')::text AS 指定學年科目名稱
										FROM (
												SELECT 
													$semester_subject_score_archive.*
													, 	unnest(xpath('//SemesterSubjectScoreInfo/Subject', xmlparse(content score_info))) as subj_score_ele
												FROM 
													$semester_subject_score_archive
												WHERE ref_student_id IN( {0}) 
											) as sems_subj_score_ext
										LEFT JOIN student ON sems_subj_score_ext.ref_student_id=student.id
										LEFT JOIN class ON student.ref_class_id=class.id
										LEFT JOIN dept g ON g.id= class.ref_dept_id  
										LEFT JOIN dept h ON h.id=student.ref_dept_id 
										ORDER BY class.grade_year,  class.display_order, student.seat_no, sems_subj_score_ext.last_update";
            sql = string.Format(sql, string.Join(",", studentIDList));

            try
            {
                var dt = queryHelper.Select(sql);
                _bgWorker.ReportProgress(50);

                _wb = new Workbook(new MemoryStream(Properties.Resources.學期成績封存樣板));

                //填樣板
                foreach (DataRow dr in dt.Rows)
                {
                    _wb.Worksheets[0].Cells[rowIndex, 0].PutValue(dr["學生系統編號"]);
                    _wb.Worksheets[0].Cells[rowIndex, 1].PutValue(dr["學號"]);
                    _wb.Worksheets[0].Cells[rowIndex, 2].PutValue(dr["班級"]);
                    _wb.Worksheets[0].Cells[rowIndex, 3].PutValue(dr["座號"]);
                    _wb.Worksheets[0].Cells[rowIndex, 4].PutValue(dr["科別"]);
                    _wb.Worksheets[0].Cells[rowIndex, 5].PutValue(dr["姓名"]);
                    _wb.Worksheets[0].Cells[rowIndex, 6].PutValue(dr["封存日期"]);
                    _wb.Worksheets[0].Cells[rowIndex, 7].PutValue(dr["學年度"]);
                    _wb.Worksheets[0].Cells[rowIndex, 8].PutValue(dr["學期"]);
                    _wb.Worksheets[0].Cells[rowIndex, 9].PutValue(dr["成績年級"]);
                    _wb.Worksheets[0].Cells[rowIndex, 10].PutValue(dr["分項類別"]);
                    _wb.Worksheets[0].Cells[rowIndex, 11].PutValue(dr["科目"]);
                    _wb.Worksheets[0].Cells[rowIndex, 12].PutValue(dr["科目級別"]);
                    _wb.Worksheets[0].Cells[rowIndex, 13].PutValue(dr["學分數"]);
                    _wb.Worksheets[0].Cells[rowIndex, 14].PutValue(dr["必選修"]);
                    _wb.Worksheets[0].Cells[rowIndex, 15].PutValue(dr["校部訂"].ToString() == "部訂" ? "部定" : dr["校部訂"].ToString()) ;
                    _wb.Worksheets[0].Cells[rowIndex, 16].PutValue(dr["取得學分"]);
                    _wb.Worksheets[0].Cells[rowIndex, 17].PutValue(dr["原始成績"]);
                    _wb.Worksheets[0].Cells[rowIndex, 18].PutValue(dr["補考成績"]);
                    _wb.Worksheets[0].Cells[rowIndex, 19].PutValue(dr["重修成績"]);
                    _wb.Worksheets[0].Cells[rowIndex, 20].PutValue(dr["手動調整成績"]);
                    _wb.Worksheets[0].Cells[rowIndex, 21].PutValue(dr["學年調整成績"]);
                    _wb.Worksheets[0].Cells[rowIndex, 22].PutValue(dr["不計學分"]);
                    _wb.Worksheets[0].Cells[rowIndex, 23].PutValue(dr["不需評分"]);
                    _wb.Worksheets[0].Cells[rowIndex, 24].PutValue(dr["註記"]);
                    _wb.Worksheets[0].Cells[rowIndex, 25].PutValue(dr["修課及格標準"]);
                    _wb.Worksheets[0].Cells[rowIndex, 26].PutValue(dr["修課補考標準"]);
                    _wb.Worksheets[0].Cells[rowIndex, 27].PutValue(dr["修課直接指定總成績"]);
                    _wb.Worksheets[0].Cells[rowIndex, 28].PutValue(dr["修課備註"]);
                    //_wb.Worksheets[0].Cells[rowIndex, 29].PutValue(dr["修課科目代碼"]);
                    _wb.Worksheets[0].Cells[rowIndex, 29].PutValue(dr["是否補修成績"]);
                    _wb.Worksheets[0].Cells[rowIndex, 30].PutValue(dr["重修學年度"]);
                    _wb.Worksheets[0].Cells[rowIndex, 31].PutValue(dr["重修學期"]);
                    _wb.Worksheets[0].Cells[rowIndex, 32].PutValue(dr["補修學年度"]);
                    _wb.Worksheets[0].Cells[rowIndex, 33].PutValue(dr["補修學期"]);
                    _wb.Worksheets[0].Cells[rowIndex, 34].PutValue(dr["免修"]);
                    _wb.Worksheets[0].Cells[rowIndex, 35].PutValue(dr["抵免"]);
                    _wb.Worksheets[0].Cells[rowIndex, 36].PutValue(dr["指定學年科目名稱"]);
                    rowIndex++;
                }

                _bgWorker.ReportProgress(100);

            }
            catch (Exception ex)
            {
                FISCA.Presentation.Controls.MsgBox.Show("取得封存資料發生錯誤");
            }
        }

        public void Run()
        {
            _bgWorker.RunWorkerAsync();
        }

    }
}
