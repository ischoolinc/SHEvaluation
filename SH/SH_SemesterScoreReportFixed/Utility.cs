﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using FISCA.Data;
using System.Data;
using K12.Data;
using SmartSchool.Evaluation.WearyDogComputerHelper;
using FISCA.DSAUtil;
using System.Xml;
using SmartSchool.Customization.Data;


namespace SH_SemesterScoreReportFixed
{
    class Utility
    {
        /// <summary>
        /// 取得學生分項成績(學業、體育、國防通識、健康與護理、實習科目、德行)
        /// </summary>
        /// <param name="studentIDLis"></param>
        /// <param name="SchoolYear"></param>
        /// <param name="Semester"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, decimal?>> GetStudentSemsEntryScore(List<string> studentIDLis, int SchoolYear, int Semester)
        {
            // DB XML
            //<Entry 分項="學業" 成績="81.5"/><Entry 分項="體育" 成績="90"/><Entry 分項="健康與護理" 成績="95"/><Entry 分項="國防通識" 成績="89"/><Entry 分項="實習科目" 成績="96"/></SemesterEntryScore>"
            //<SemesterEntryScore>
            //<Entry 分項="德行" 成績="83.4" 鎖定="False"/>
            //</SemesterEntryScore>
            Dictionary<string, Dictionary<string, decimal?>> retVal = new Dictionary<string, Dictionary<string, decimal?>>();

            if (studentIDLis.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string strSQL = "select ref_student_id,score_info from sems_entry_score where ref_student_id in(" + string.Join(",", studentIDLis.ToArray()) + ") and school_year=" + SchoolYear + " and semester=" + Semester + " order by ref_student_id";
                DataTable dt = qh.Select(strSQL);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr[0].ToString();

                    if (!retVal.ContainsKey(sid))
                    {
                        Dictionary<string, decimal?> val = new Dictionary<string, decimal?>();
                        retVal.Add(sid, val);
                    }

                    string strScore = dr[1].ToString();
                    if (!string.IsNullOrWhiteSpace(strScore))
                    {
                        XElement elmScore = XElement.Parse(strScore);
                        foreach (XElement elm in elmScore.Elements("Entry"))
                        {
                            if (elm.Attribute("分項") != null)
                            {
                                string name = elm.Attribute("分項").Value;
                                decimal dd;
                                retVal[sid].Add(name, null);
                                if (elm.Attribute("成績") != null)
                                    if (decimal.TryParse(elm.Attribute("成績").Value, out dd))
                                        retVal[sid][name] = dd;
                            }
                        }
                    }
                }
            }

            return retVal;
        }

        /// <summary>
        /// 透過學年度學期取得學生缺曠統計(傳入學生系統編號、開始日期、結束日期；回傳：學生系統編號、獎懲名稱,統計值
        /// </summary>
        /// <param name="StudIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, int>> GetAttendanceCountBySchoolYearSemester(List<K12.Data.StudentRecord> StudRecordList, int SchoolYear, int Semester)
        {
            Dictionary<string, Dictionary<string, int>> retVal = new Dictionary<string, Dictionary<string, int>>();

            List<PeriodMappingInfo> PeriodMappingList = PeriodMapping.SelectAll();
            // 節次>類別
            Dictionary<string, string> PeriodMappingDict = new Dictionary<string, string>();
            foreach (PeriodMappingInfo rec in PeriodMappingList)
            {
                if (!PeriodMappingDict.ContainsKey(rec.Name))
                    PeriodMappingDict.Add(rec.Name, rec.Type);
            }


            List<AttendanceRecord> attendList = K12.Data.Attendance.SelectBySchoolYearAndSemester(StudRecordList, SchoolYear, Semester);

            // 計算統計資料
            foreach (AttendanceRecord rec in attendList)
            {
                if (!retVal.ContainsKey(rec.RefStudentID))
                    retVal.Add(rec.RefStudentID, new Dictionary<string, int>());

                foreach (AttendancePeriod per in rec.PeriodDetail)
                {
                    if (!PeriodMappingDict.ContainsKey(per.Period))
                        continue;

                    // ex.一般:曠課
                    string key = PeriodMappingDict[per.Period] + "_" + per.AbsenceType;

                    if (!retVal[rec.RefStudentID].ContainsKey(key))
                        retVal[rec.RefStudentID].Add(key, 0);

                    retVal[rec.RefStudentID][key]++;
                }
            }

            return retVal;
        }

        /// <summary>
        /// 透過學年度取得學生缺曠統計(傳入學生系統編號、開始日期、結束日期；回傳：學生系統編號、獎懲名稱,統計值
        /// </summary>
        /// <param name="StudIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, int>> GetAttendanceCountBySchoolYear(List<K12.Data.StudentRecord> StudRecordList, int SchoolYear)
        {
            Dictionary<string, Dictionary<string, int>> retVal = new Dictionary<string, Dictionary<string, int>>();

            List<PeriodMappingInfo> PeriodMappingList = PeriodMapping.SelectAll();
            // 節次>類別
            Dictionary<string, string> PeriodMappingDict = new Dictionary<string, string>();
            foreach (PeriodMappingInfo rec in PeriodMappingList)
            {
                if (!PeriodMappingDict.ContainsKey(rec.Name))
                    PeriodMappingDict.Add(rec.Name, rec.Type);
            }


            List<AttendanceRecord> attendList = K12.Data.Attendance.SelectBySchoolYearAndSemester(StudRecordList, SchoolYear, null);

            // 計算統計資料
            foreach (AttendanceRecord rec in attendList)
            {
                if (!retVal.ContainsKey(rec.RefStudentID))
                    retVal.Add(rec.RefStudentID, new Dictionary<string, int>());

                foreach (AttendancePeriod per in rec.PeriodDetail)
                {
                    if (!PeriodMappingDict.ContainsKey(per.Period))
                        continue;

                    // ex.一般:曠課
                    string key = PeriodMappingDict[per.Period] + "_" + per.AbsenceType;

                    if (!retVal[rec.RefStudentID].ContainsKey(key))
                        retVal[rec.RefStudentID].Add(key, 0);

                    retVal[rec.RefStudentID][key]++;
                }
            }

            return retVal;
        }

        //public static void test()
        //{
        //    SmartSchool.Customization.Data.AccessHelper ac = new SmartSchool.Customization.Data.AccessHelper();
        //    List<SmartSchool.Customization.Data.StudentRecord> listS = ac.StudentHelper.GetSelectedStudent();

        //     ac.StudentHelper.FillSemesterSubjectScore(true,listS);


        //     SmartSchool.Evaluation.WearyDogComputer computer = new SmartSchool.Evaluation.WearyDogComputer();

        //     computer.FillStudentGradCalcScore(ac, listS);
        //     Dictionary<SmartSchool.Customization.Data.StudentRecord, List<string>> errormessages = computer.FillStudentGradCheck(ac, listS);

        //     foreach (SmartSchool.Customization.Data.StudentRecord rec in listS)
        //     {
        //         foreach (SmartSchool.Customization.Data.StudentExtension.SemesterEntryScoreInfo se in rec.SemesterEntryScoreList)
        //         { 

        //         }

        //     }




        //}


        /// <summary>
        /// 透過學生編號取得學生服務學習時數 傳入學生編號、學年度,回傳：學生編號、內容、值
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, decimal> GetServiceLearningBySchoolYear(List<string> StudentIDList, int SchoolYear)
        {
            Dictionary<string, decimal> retVal = new Dictionary<string, decimal>();

            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,school_year,semester,hours from $k12.service.learning.record where ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "') and school_year=" + SchoolYear + " order by ref_student_id,school_year,semester;";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr[0].ToString();
                    decimal hr;
                    decimal.TryParse(dr[3].ToString(), out hr);

                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, 0);

                    retVal[sid] += hr;

                }
            }
            return retVal;
        }

        /// <summary>
        /// 透過學生編號取得學生服務學習時數 傳入學生編號、學年度,回傳：學生編號、內容、值
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <param name="beginDate"></param>
        /// <param name="endDate"></param>
        /// <returns></returns>
        public static Dictionary<string, decimal> GetServiceLearningBySchoolYearSemester(List<string> StudentIDList, int SchoolYear, int Semester)
        {
            Dictionary<string, decimal> retVal = new Dictionary<string, decimal>();

            if (StudentIDList.Count > 0)
            {
                QueryHelper qh = new QueryHelper();
                string query = "select ref_student_id,school_year,semester,hours from $k12.service.learning.record where ref_student_id in('" + string.Join("','", StudentIDList.ToArray()) + "') and school_year=" + SchoolYear + " and semester=" + Semester + " order by ref_student_id,school_year,semester;";
                DataTable dt = qh.Select(query);
                foreach (DataRow dr in dt.Rows)
                {
                    string sid = dr[0].ToString();
                    decimal hr;
                    decimal.TryParse(dr[3].ToString(), out hr);

                    if (!retVal.ContainsKey(sid))
                        retVal.Add(sid, 0);

                    retVal[sid] += hr;

                }
            }
            return retVal;
        }


        /// <summary>
        /// 取得缺曠對照 List,一般_曠課..
        /// </summary>
        /// <returns></returns>
        public static List<string> GetATMappingKey()
        {
            List<string> retVal = new List<string>();
            List<string> key1List = new List<string>();
            List<string> Key2List = new List<string>();
            foreach (PeriodMappingInfo data in PeriodMapping.SelectAll())
                if (!key1List.Contains(data.Type))
                    key1List.Add(data.Type);

            foreach (AbsenceMappingInfo data in AbsenceMapping.SelectAll())
                if (!Key2List.Contains(data.Name))
                    Key2List.Add(data.Name);

            // 一般_曠課
            foreach (string key1 in key1List)
                foreach (string key2 in Key2List)
                    retVal.Add(key1 + "_" + key2);

            return retVal;
        }

        /// <summary>
        /// 取得學生及格與補考標準，參數用學生IDList,回傳:key:StudentID,1_及,數字
        /// </summary>
        /// <param name="StudRecList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, decimal>> GetStudentApplyLimitDict(List<SmartSchool.Customization.Data.StudentRecord> StudRecList)
        {

            Dictionary<string, Dictionary<string, decimal>> retVal = new Dictionary<string, Dictionary<string, decimal>>();


            foreach (SmartSchool.Customization.Data.StudentRecord studRec in StudRecList)
            {
                //及格標準<年級,及格與補考標準>
                if (!retVal.ContainsKey(studRec.StudentID))
                    retVal.Add(studRec.StudentID, new Dictionary<string, decimal>());

                XmlElement scoreCalcRule = SmartSchool.Evaluation.ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(studRec.StudentID) == null ? null : SmartSchool.Evaluation.ScoreCalcRule.ScoreCalcRule.Instance.GetStudentScoreCalcRuleInfo(studRec.StudentID).ScoreCalcRuleElement;
                if (scoreCalcRule == null)
                {

                }
                else
                {

                    DSXmlHelper helper = new DSXmlHelper(scoreCalcRule);
                    decimal tryParseDecimal;
                    decimal tryParseDecimala;

                    foreach (XmlElement element in helper.GetElements("及格標準/學生類別"))
                    {
                        string cat = element.GetAttribute("類別");
                        bool useful = false;
                        //掃描學生的類別作比對
                        foreach (CategoryInfo catinfo in studRec.StudentCategorys)
                        {
                            if (catinfo.Name == cat || catinfo.FullName == cat)
                                useful = true;
                        }
                        //學生是指定的類別或類別為"預設"
                        if (cat == "預設" || useful)
                        {
                            for (int gyear = 1; gyear <= 4; gyear++)
                            {
                                switch (gyear)
                                {
                                    case 1:
                                        if (decimal.TryParse(element.GetAttribute("一年級及格標準"), out tryParseDecimal))
                                        {
                                            string k1s = gyear + "_及";

                                            if (!retVal[studRec.StudentID].ContainsKey(k1s))
                                                retVal[studRec.StudentID].Add(k1s, tryParseDecimal);

                                            if (retVal[studRec.StudentID][k1s] > tryParseDecimal)
                                                retVal[studRec.StudentID][k1s] = tryParseDecimal;
                                        }

                                        if (decimal.TryParse(element.GetAttribute("一年級補考標準"), out tryParseDecimala))
                                        {
                                            string k1a = gyear + "_補";

                                            if (!retVal[studRec.StudentID].ContainsKey(k1a))
                                                retVal[studRec.StudentID].Add(k1a, tryParseDecimala);

                                            if (retVal[studRec.StudentID][k1a] > tryParseDecimala)
                                                retVal[studRec.StudentID][k1a] = tryParseDecimala;
                                        }

                                        break;
                                    case 2:
                                        if (decimal.TryParse(element.GetAttribute("二年級及格標準"), out tryParseDecimal))
                                        {
                                            string k2s = gyear + "_及";

                                            if (!retVal[studRec.StudentID].ContainsKey(k2s))
                                                retVal[studRec.StudentID].Add(k2s, tryParseDecimal);

                                            if (retVal[studRec.StudentID][k2s] > tryParseDecimal)
                                                retVal[studRec.StudentID][k2s] = tryParseDecimal;
                                        }

                                        if (decimal.TryParse(element.GetAttribute("二年級補考標準"), out tryParseDecimala))
                                        {
                                            string k2a = gyear + "_補";

                                            if (!retVal[studRec.StudentID].ContainsKey(k2a))
                                                retVal[studRec.StudentID].Add(k2a, tryParseDecimala);

                                            if (retVal[studRec.StudentID][k2a] > tryParseDecimala)
                                                retVal[studRec.StudentID][k2a] = tryParseDecimala;

                                        }

                                        break;
                                    case 3:
                                        if (decimal.TryParse(element.GetAttribute("三年級及格標準"), out tryParseDecimal))
                                        {
                                            string k3s = gyear + "_及";

                                            if (!retVal[studRec.StudentID].ContainsKey(k3s))
                                                retVal[studRec.StudentID].Add(k3s, tryParseDecimal);

                                            if (retVal[studRec.StudentID][k3s] > tryParseDecimal)
                                                retVal[studRec.StudentID][k3s] = tryParseDecimal;
                                        }

                                        if (decimal.TryParse(element.GetAttribute("三年級補考標準"), out tryParseDecimala))
                                        {
                                            string k3a = gyear + "_補";

                                            if (!retVal[studRec.StudentID].ContainsKey(k3a))
                                                retVal[studRec.StudentID].Add(k3a, tryParseDecimala);

                                            if (retVal[studRec.StudentID][k3a] > tryParseDecimala)
                                                retVal[studRec.StudentID][k3a] = tryParseDecimala;
                                        }

                                        break;
                                    case 4:
                                        if (decimal.TryParse(element.GetAttribute("四年級及格標準"), out tryParseDecimal))
                                        {
                                            string k4s = gyear + "_及";

                                            if (!retVal[studRec.StudentID].ContainsKey(k4s))
                                                retVal[studRec.StudentID].Add(k4s, tryParseDecimal);

                                            if (retVal[studRec.StudentID][k4s] > tryParseDecimal)
                                                retVal[studRec.StudentID][k4s] = tryParseDecimal;
                                        }

                                        if (decimal.TryParse(element.GetAttribute("四年級補考標準"), out tryParseDecimala))
                                        {
                                            string k4a = gyear + "_補";

                                            if (!retVal[studRec.StudentID].ContainsKey(k4a))
                                                retVal[studRec.StudentID].Add(k4a, tryParseDecimala);

                                            if (retVal[studRec.StudentID][k4a] > tryParseDecimala)
                                                retVal[studRec.StudentID][k4a] = tryParseDecimala;
                                        }

                                        break;
                                    default:
                                        break;
                                }
                            }
                        }
                    }
                }
            }
            return retVal;
        }


        /// <summary>
        /// 取得學生評量排名、五標與組距資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, DataRow>> GetRankMatrixData(string SchoolYear, string Semester, string ExamID, List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, DataRow>> value = new Dictionary<string, Dictionary<string, DataRow>>();

            // 沒有學生不處理
            if (StudentIDList.Count == 0)
                return value;

            QueryHelper qh = new QueryHelper();
            string query = "" +
               " SELECT " +
" 	rank_matrix.id AS rank_matrix_id" +
" 	, rank_matrix.school_year" +
" 	, rank_matrix.semester" +
" 	, rank_matrix.grade_year" +
" 	, rank_matrix.item_type" +
" 	, rank_matrix.ref_exam_id AS exam_id" +
" 	, rank_matrix.item_name" +
" 	, rank_matrix.rank_type" +
" 	, rank_matrix.rank_name" +
" 	, class.class_name" +
" 	, student.seat_no" +
" 	, student.student_number" +
" 	, student.name" +
" 	, rank_detail.ref_student_id AS student_id " +
" 	, rank_detail.rank" +
"   , rank_matrix.matrix_count " +
" 	, rank_detail.pr" +
" 	, rank_detail.percentile" +
"   , rank_matrix.avg_top_25" +
"   , rank_matrix.avg_top_50" +
"   , rank_matrix.avg" +
"   , rank_matrix.avg_bottom_50" +
"   , rank_matrix.avg_bottom_25" +
" 	, rank_matrix.level_gte100" +
" 	, rank_matrix.level_90" +
" 	, rank_matrix.level_80" +
" 	, rank_matrix.level_70" +
" 	, rank_matrix.level_60" +
" 	, rank_matrix.level_50" +
" 	, rank_matrix.level_40" +
" 	, rank_matrix.level_30" +
" 	, rank_matrix.level_20" +
" 	, rank_matrix.level_10" +
" 	, rank_matrix.level_lt10" +
" FROM " +
" 	rank_matrix" +
" 	LEFT OUTER JOIN rank_detail" +
" 		ON rank_detail.ref_matrix_id = rank_matrix.id" +
" 	LEFT OUTER JOIN student" +
" 		ON student.id = rank_detail.ref_student_id" +
" 	LEFT OUTER JOIN class" +
" 		ON class.id = student.ref_class_id" +
" WHERE" +
" 	rank_matrix.is_alive = true" +
" 	AND rank_matrix.school_year = " + SchoolYear +
"     AND rank_matrix.semester = " + Semester +
" 	AND rank_matrix.item_type like '定期評量%'" +
" 	AND rank_matrix.ref_exam_id = " + ExamID +
"     AND ref_student_id IN (" + string.Join(",", StudentIDList.ToArray()) + ") " +
" ORDER BY " +
" 	rank_matrix.id" +
" 	, rank_detail.rank" +
" 	, class.grade_year" +
" 	, class.display_order" +
" 	, class.class_name" +
" 	, student.seat_no" +
" 	, student.id";

            DataTable dt = qh.Select(query);
            //dt.TableName = "d5";
            //dt.WriteXmlSchema(Application.StartupPath + "\\d5s.xml");
            //dt.WriteXml(Application.StartupPath + "\\d5d.xml");

            // student id key
            // key = item_type + item_name +  rank_name 
            foreach (DataRow dr in dt.Rows)
            {
                string sid = dr["student_id"].ToString();
                if (!value.ContainsKey(sid))
                    value.Add(sid, new Dictionary<string, DataRow>());

                string key = dr["item_type"].ToString() + "_" + dr["item_name"].ToString() + "_" + dr["rank_type"].ToString();

                if (!value[sid].ContainsKey(key))
                    value[sid].Add(key, dr);
            }

            return value;
        }


        /// <summary>
        /// 取得學生學期排名、五標與組距資料
        /// </summary>
        /// <param name="StudentIDList"></param>
        /// <returns></returns>
        public static Dictionary<string, Dictionary<string, DataRow>> GetSemsScoreRankMatrixData(string SchoolYear, string Semester,  List<string> StudentIDList)
        {
            Dictionary<string, Dictionary<string, DataRow>> value = new Dictionary<string, Dictionary<string, DataRow>>();

            // 沒有學生不處理
            if (StudentIDList.Count == 0)
                return value;

            QueryHelper qh = new QueryHelper();
            string query = "" +
               " SELECT " +
" 	rank_matrix.id AS rank_matrix_id" +
" 	, rank_matrix.school_year" +
" 	, rank_matrix.semester" +
" 	, rank_matrix.grade_year" +
" 	, rank_matrix.item_type" +
" 	, rank_matrix.ref_exam_id AS exam_id" +
" 	, rank_matrix.item_name" +
" 	, rank_matrix.rank_type" +
" 	, rank_matrix.rank_name" +
" 	, class.class_name" +
" 	, student.seat_no" +
" 	, student.student_number" +
" 	, student.name" +
" 	, rank_detail.ref_student_id AS student_id " +
" 	, rank_detail.rank" +
"   , rank_matrix.matrix_count " +
" 	, rank_detail.pr" +
" 	, rank_detail.percentile" +
"   , rank_matrix.avg_top_25" +
"   , rank_matrix.avg_top_50" +
"   , rank_matrix.avg" +
"   , rank_matrix.avg_bottom_50" +
"   , rank_matrix.avg_bottom_25" +
" 	, rank_matrix.level_gte100" +
" 	, rank_matrix.level_90" +
" 	, rank_matrix.level_80" +
" 	, rank_matrix.level_70" +
" 	, rank_matrix.level_60" +
" 	, rank_matrix.level_50" +
" 	, rank_matrix.level_40" +
" 	, rank_matrix.level_30" +
" 	, rank_matrix.level_20" +
" 	, rank_matrix.level_10" +
" 	, rank_matrix.level_lt10" +
" FROM " +
" 	rank_matrix" +
" 	LEFT OUTER JOIN rank_detail" +
" 		ON rank_detail.ref_matrix_id = rank_matrix.id" +
" 	LEFT OUTER JOIN student" +
" 		ON student.id = rank_detail.ref_student_id" +
" 	LEFT OUTER JOIN class" +
" 		ON class.id = student.ref_class_id" +
" WHERE" +
" 	rank_matrix.is_alive = true" +
" 	AND rank_matrix.school_year = " + SchoolYear +
"     AND rank_matrix.semester = " + Semester +
" 	AND rank_matrix.item_type like '學期%'" +
" 	AND rank_matrix.ref_exam_id = -1 " +
"     AND ref_student_id IN (" + string.Join(",", StudentIDList.ToArray()) + ") " +
" ORDER BY " +
" 	rank_matrix.id" +
" 	, rank_detail.rank" +
" 	, class.grade_year" +
" 	, class.display_order" +
" 	, class.class_name" +
" 	, student.seat_no" +
" 	, student.id";

            DataTable dt = qh.Select(query);
            //dt.TableName = "d5";
            //dt.WriteXmlSchema(Application.StartupPath + "\\d5s.xml");
            //dt.WriteXml(Application.StartupPath + "\\d5d.xml");

            // student id key
            // key = item_type + item_name +  rank_name 
            foreach (DataRow dr in dt.Rows)
            {
                string sid = dr["student_id"].ToString();
                if (!value.ContainsKey(sid))
                    value.Add(sid, new Dictionary<string, DataRow>());

                string key = dr["item_type"].ToString() + "_" + dr["item_name"].ToString() + "_" + dr["rank_type"].ToString();

                if (!value[sid].ContainsKey(key))
                    value[sid].Add(key, dr);
            }

            return value;
        }


    }
}
