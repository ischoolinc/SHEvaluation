using FISCA.Data;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartSchool.Evaluation.Process.Wizards.LearningHistory
{
    public class LearningHistoryDataAccess
    {

        private JSubjectInfo CreateJSubjectInfo(SubjectScoreRec108 ssr, string serialNo, string name, int SchoolYear, int Semester)
        {
            JSubjectInfo JInfo = new JSubjectInfo();
            JInfo.refStudentID = ssr.StudentID;
            JInfo.serialNo = serialNo;
            JInfo.name = name;
            JInfo.schoolYear = SchoolYear;
            JInfo.semester = Semester;
            JInfo.subject = ssr.SubjectName;
            JInfo.subject = ssr.SubjectName;
            JInfo.subjectLevel = null;
            if (!string.IsNullOrEmpty(ssr.SubjectLevel))
                JInfo.subjectLevel = int.Parse(ssr.SubjectLevel);
            JInfo.detail = new List<JSubjectDetail>();

            return JInfo;
        }

        private JSubjectInfo CreateJSubjectInfoN(SubjectYearScoreRec108N ssr, string serialNo, string name, int SchoolYear, int Semester)
        {
            JSubjectInfo JInfo = new JSubjectInfo();
            JInfo.refStudentID = ssr.StudentID;
            JInfo.serialNo = serialNo;
            JInfo.name = name;
            JInfo.schoolYear = SchoolYear;
            JInfo.semester = Semester;
            JInfo.subject = ssr.SubjectName;
            JInfo.subject = ssr.SubjectName;
            JInfo.subjectLevel = null;
            JInfo.detail = new List<JSubjectDetail>();

            return JInfo;
        }

        private JSubjectDetail CreateJSubjectDetail(string serialNo, string name, string value)
        {
            return new JSubjectDetail
            {
                serialNo = serialNo,
                name = name,
                value = value
            };
        }

        private void SaveToDatabase(string strSQL)
        {
            if (strSQL.Length < 1)
            {
                return;
            }

            var qh = new QueryHelper();
            qh.Select(strSQL);
        }

        public void SaveScores42(List<SubjectScoreRec108> scores, int SchoolYear, int Semester)
        {
            var jSubjectInfoList = new List<JSubjectInfo>();
            foreach (var ssr in scores)
            {
                if (ssr.checkPass)
                {
                    // 跳過補修
                    if (ssr.isScScore)
                        continue;
                    if (!ssr.CodePass) //跳過課程代碼不可提交
                        continue;

                    var jInfo = CreateJSubjectInfo(ssr, "4.2", "學期成績", SchoolYear, Semester);
                    jInfo.detail = new List<JSubjectDetail>
                {
                 CreateJSubjectDetail("4.2.1", "身分證號", ssr.IDNumber),
                 CreateJSubjectDetail("4.2.2", "出生日期", ssr.Birthday),
                 CreateJSubjectDetail("4.2.3", "課程代碼", ssr.CourseCode),
                 CreateJSubjectDetail("4.2.4", "科目名稱", ssr.SubjectName),
                 CreateJSubjectDetail("4.2.5", "開課年級", ssr.GradeYear),
                 CreateJSubjectDetail("4.2.6", "修課學分", ssr.Credit),
                 CreateJSubjectDetail("4.2.7", "學期學業成績", ssr.Score),
                 CreateJSubjectDetail("4.2.8", "成績及格", ssr.ScoreP),
                 CreateJSubjectDetail("4.2.9", "補考成績", ssr.ReScore),
                 CreateJSubjectDetail("4.2.10", "補考及格", ssr.ReScoreP),
                 CreateJSubjectDetail("4.2.11", "是否採計學分", ssr.useCredit),
                 CreateJSubjectDetail("4.2.12", "質性文字描述", ssr.Text),
                 CreateJSubjectDetail("4.2.13", "DataKey", ""),
                 CreateJSubjectDetail("4.2.14", "備註(學生姓名)", ssr.Name),
                 CreateJSubjectDetail("4.2.15", "備註(資料當學期班級)", ssr.HisClassName),
                 CreateJSubjectDetail("4.2.16", "備註(資料當學期座號)", ssr.HisSeatNo + ""),
                 CreateJSubjectDetail("4.2.17", "備註(資料當學期學號)", ssr.HisStudentNumber),
                 CreateJSubjectDetail("4.2.18", "備註(資料次學期班級)", ssr.ClassName),
                 CreateJSubjectDetail("4.2.19", "備註(資料次學期座號)", ssr.SeatNo),
                 CreateJSubjectDetail("4.2.20", "備註(資料次學期學號)", ssr.StudentNumber)
                };
                    jSubjectInfoList.Add(jInfo);
                }
            }
            SaveToDatabase(GetSemesterScoreSqlTemplate42(jSubjectInfoList));
        }
        public void SaveScores43(List<SubjectScoreRec108> scores, int SchoolYear, int Semester)
        {
            var jSubjectInfoList = new List<JSubjectInfo>();
            foreach (var ssr in scores)
            {
                if (ssr.checkPass)
                {
                    if (!ssr.CodePass) //跳過課程代碼不可提交
                        continue;

                    // 補修
                    if (ssr.isScScore)
                    {
                        var jInfo = CreateJSubjectInfo(ssr, "4.3", "補修成績", SchoolYear, Semester);
                        jInfo.detail = new List<JSubjectDetail>
                        {
                            CreateJSubjectDetail("4.3.1", "身分證號", ssr.IDNumber),
                            CreateJSubjectDetail("4.3.2", "出生日期", ssr.Birthday),
                            CreateJSubjectDetail("4.3.3", "應修課學年度", ssr.SchoolYear),
                            CreateJSubjectDetail("4.3.4", "應修課學期", ssr.Semester),
                            CreateJSubjectDetail("4.3.5", "課程代碼", ssr.CourseCode),
                            CreateJSubjectDetail("4.3.6", "科目名稱", ssr.SubjectName),
                            CreateJSubjectDetail("4.3.7", "開課年級", ssr.GradeYear),
                            CreateJSubjectDetail("4.3.8", "修課學分", ssr.Credit),
                            CreateJSubjectDetail("4.3.9", "補修成績", ssr.Score),
                            CreateJSubjectDetail("4.3.10", "補修及格", ssr.ScoreP),
                            CreateJSubjectDetail("4.3.11", "補考成績", ssr.ReScore),
                            CreateJSubjectDetail("4.3.12", "補考及格", ssr.ReScoreP),
                            CreateJSubjectDetail("4.3.13", "補修方式", ssr.ScScoreType),
                            CreateJSubjectDetail("4.3.14", "是否採計學分", ssr.useCredit),
                            CreateJSubjectDetail("4.3.15", "質性文字描述", ssr.Text),
                            CreateJSubjectDetail("4.3.16", "DataKey", ""),
                            CreateJSubjectDetail("4.3.17", "備註(學生姓名)", ssr.Name),
                            CreateJSubjectDetail("4.3.18", "備註(資料當學期班級)", ssr.HisClassName),
                            CreateJSubjectDetail("4.3.19", "備註(資料當學期座號)", ssr.HisSeatNo + ""),
                            CreateJSubjectDetail("4.3.20", "備註(資料當學期學號)", ssr.HisStudentNumber),
                            CreateJSubjectDetail("4.3.21", "備註(資料次學期班級)", ssr.ClassName),
                            CreateJSubjectDetail("4.3.22", "備註(資料次學期座號)", ssr.SeatNo),
                            CreateJSubjectDetail("4.3.23", "備註(資料次學期學號)", ssr.StudentNumber)
                        };
                        jSubjectInfoList.Add(jInfo);
                    }

                }
            }
            SaveToDatabase(GetSemesterScoreSqlTemplate43(jSubjectInfoList));
        }

        public void SaveScores44(List<SubjectScoreRec108> scores, int SchoolYear, int Semester)
        {
            var jSubjectInfoList = new List<JSubjectInfo>();
            foreach (var ssr in scores)
            {
                if (ssr.checkPass)
                {
                    if (!ssr.CodePass) //跳過課程代碼不可提交
                        continue;

                    var jInfo = CreateJSubjectInfo(ssr, "4.4", "轉學轉科成績", SchoolYear, Semester);
                    jInfo.detail = new List<JSubjectDetail>
                        {
                            CreateJSubjectDetail("4.4.1", "身分證號", ssr.IDNumber),
                            CreateJSubjectDetail("4.4.2", "出生日期", ssr.Birthday),
                            CreateJSubjectDetail("4.4.3", "對應學年度", ssr.SchoolYear),
                            CreateJSubjectDetail("4.4.4", "對應學期", ssr.Semester),
                            CreateJSubjectDetail("4.4.5", "課程代碼", ssr.CourseCode),
                            CreateJSubjectDetail("4.4.6", "科目名稱", ssr.SubjectName),
                            CreateJSubjectDetail("4.4.7", "對應年級", ssr.GradeYear),
                            CreateJSubjectDetail("4.4.8", "修課學分", ssr.Credit),
                            CreateJSubjectDetail("4.4.9", "身分別", ssr.StudType),
                            CreateJSubjectDetail("4.4.10", "抵免後成績", ssr.Score),
                            CreateJSubjectDetail("4.4.11", "成績及格", ssr.ScoreP),
                            CreateJSubjectDetail("4.4.12", "是否採計學分", ssr.useCredit),
                            CreateJSubjectDetail("4.4.13", "質性文字描述", ssr.Text),
                            CreateJSubjectDetail("4.4.14", "DataKey", ""),
                            CreateJSubjectDetail("4.4.15", "備註(學生姓名)", ssr.Name),
                            CreateJSubjectDetail("4.4.16", "備註(資料當學期班級)", ssr.HisClassName),
                            CreateJSubjectDetail("4.4.17", "備註(資料當學期座號)", ssr.HisSeatNo + ""),
                            CreateJSubjectDetail("4.4.18", "備註(資料當學期學號)", ssr.HisStudentNumber),
                            CreateJSubjectDetail("4.4.19", "備註(資料次學期班級)", ssr.ClassName),
                            CreateJSubjectDetail("4.4.20", "備註(資料次學期座號)", ssr.SeatNo),
                            CreateJSubjectDetail("4.4.21", "備註(資料次學期學號)", ssr.StudentNumber)
                        };
                    jSubjectInfoList.Add(jInfo);
                }
            }
            SaveToDatabase(GetSemesterScoreSqlTemplate44(jSubjectInfoList));
        }

        public void SaveScores52(List<SubjectScoreRec108> scores, int SchoolYear, int Semester)
        {
            var jSubjectInfoList = new List<JSubjectInfo>();
            foreach (var ssr in scores)
            {
                if (ssr.checkPass)
                {
                    if (!ssr.CodePass) //跳過課程代碼不可提交
                        continue;

                    var jInfo = CreateJSubjectInfo(ssr, "5.2", "重修成績", SchoolYear, Semester);
                    jInfo.detail = new List<JSubjectDetail>
                        {
                         CreateJSubjectDetail("5.2.1", "身分證號", ssr.IDNumber),
                         CreateJSubjectDetail("5.2.2", "出生日期", ssr.Birthday),
                         CreateJSubjectDetail("5.2.3", "原修課學年度", ssr.SchoolYear),
                         CreateJSubjectDetail("5.2.4", "原修課學期", ssr.Semester),
                         CreateJSubjectDetail("5.2.5", "原修課課程代碼", ssr.CourseCode),
                         CreateJSubjectDetail("5.2.6", "原修課科目名稱", ssr.SubjectName),
                         CreateJSubjectDetail("5.2.7", "原修課開課年級", ssr.GradeYear),
                         CreateJSubjectDetail("5.2.8", "原修課修課學分", ssr.Credit),
                         CreateJSubjectDetail("5.2.9", "重修方式", ssr.ReAScoreType),
                         CreateJSubjectDetail("5.2.10", "重修成績", ssr.ReAScore),
                         CreateJSubjectDetail("5.2.11", "重修及格", ssr.ReAScoreP),
                         CreateJSubjectDetail("5.2.12", "質性文字描述", ssr.Text),
                         CreateJSubjectDetail("5.2.13", "DataKey", ""),
                         CreateJSubjectDetail("5.2.14", "備註(學生姓名)", ssr.Name),
                         CreateJSubjectDetail("5.2.15", "備註(資料當學期班級)", ssr.HisClassName),
                         CreateJSubjectDetail("5.2.16", "備註(資料當學期座號)", ssr.HisSeatNo + ""),
                         CreateJSubjectDetail("5.2.17", "備註(資料當學期學號)", ssr.HisStudentNumber),
                         CreateJSubjectDetail("5.2.18", "備註(資料次學期班級)", ssr.ClassName),
                         CreateJSubjectDetail("5.2.19", "備註(資料次學期座號)", ssr.SeatNo),
                         CreateJSubjectDetail("5.2.20", "備註(資料次學期學號)", ssr.StudentNumber)
                        };
                    jSubjectInfoList.Add(jInfo);
                }
            }
            SaveToDatabase(GetSemesterScoreSqlTemplate52(jSubjectInfoList));
        }

        public void SaveScores53(List<SubjectScoreRec108> scores, int SchoolYear, int Semester)
        {
            var jSubjectInfoList = new List<JSubjectInfo>();
            foreach (var ssr in scores)
            {
                if (ssr.checkPass)
                {
                    if (!ssr.CodePass) //跳過課程代碼不可提交
                        continue;

                    var jInfo = CreateJSubjectInfo(ssr, "5.3", "重讀成績", SchoolYear, Semester);
                    jInfo.detail = new List<JSubjectDetail>
                        {
                        CreateJSubjectDetail("5.3.1", "身分證號", ssr.IDNumber),
                        CreateJSubjectDetail("5.3.2", "出生日期", ssr.Birthday),
                        CreateJSubjectDetail("5.3.3", "課程代碼", ssr.CourseCode),
                        CreateJSubjectDetail("5.3.4", "科目名稱", ssr.SubjectName),
                        CreateJSubjectDetail("5.3.5", "開課年級", ssr.GradeYear),
                        CreateJSubjectDetail("5.3.6", "修課學分", ssr.Credit),
                        CreateJSubjectDetail("5.3.7", "再次修習成績", ssr.Score),
                        CreateJSubjectDetail("5.3.8", "再次修習成績及格", ssr.ScoreP),
                        CreateJSubjectDetail("5.3.9", "補考成績", ssr.ReScore),
                        CreateJSubjectDetail("5.3.10", "補考及格", ssr.ReScoreP),
                        CreateJSubjectDetail("5.3.11", "重讀成績", ssr.RepeatScore),
                        CreateJSubjectDetail("5.3.12", "成績及格", ssr.RepeatScoreP),
                        CreateJSubjectDetail("5.3.13", "重讀註記", ssr.RepeatMemo),
                        CreateJSubjectDetail("5.3.14", "是否採計學分", ssr.useCredit),
                        CreateJSubjectDetail("5.3.15", "質性文字描述", ssr.Text),
                        CreateJSubjectDetail("5.3.16", "DataKey", ""),
                        CreateJSubjectDetail("5.3.17", "備註(學生姓名)", ssr.Name),
                        CreateJSubjectDetail("5.3.18", "備註(資料當學期班級)", ssr.HisClassName),
                        CreateJSubjectDetail("5.3.19", "備註(資料當學期座號)", ssr.HisSeatNo + ""),
                        CreateJSubjectDetail("5.3.20", "備註(資料當學期學號)", ssr.HisStudentNumber),
                        CreateJSubjectDetail("5.3.21", "備註(資料次學期班級)", ssr.ClassName),
                        CreateJSubjectDetail("5.3.22", "備註(資料次學期座號)", ssr.SeatNo),
                        CreateJSubjectDetail("5.3.23", "備註(資料次學期學號)", ssr.StudentNumber)
                        };
                    jSubjectInfoList.Add(jInfo);
                }
            }
            SaveToDatabase(GetSemesterScoreSqlTemplate53(jSubjectInfoList));
        }


        public void SaveScores62(List<SubjectScoreRec108> scores, int SchoolYear, int Semester)
        {
            var jSubjectInfoList = new List<JSubjectInfo>();
            foreach (var ssr in scores)
            {
                if (ssr.checkPass)
                {
                    // 跳過補修
                    if (ssr.isScScore)
                        continue;
                    if (!ssr.CodePass) //跳過課程代碼不可提交
                        continue;

                    var jInfo = CreateJSubjectInfo(ssr, "6.2", "學期成績", SchoolYear, Semester);
                    jInfo.detail = new List<JSubjectDetail>
                              {
                    CreateJSubjectDetail("6.2.1", "身分證號", ssr.IDNumber),
                    CreateJSubjectDetail("6.2.2", "出生日期", ssr.Birthday),
                    CreateJSubjectDetail("6.2.3", "課程代碼", ssr.CourseCode),
                    CreateJSubjectDetail("6.2.4", "科目名稱", ssr.SubjectName),
                    CreateJSubjectDetail("6.2.5", "開課年級", ssr.GradeYear),
                    CreateJSubjectDetail("6.2.6", "修課節數", ssr.Credit),
                    CreateJSubjectDetail("6.2.7", "學期學業成績", ssr.Score),
                    CreateJSubjectDetail("6.2.8", "成績及格", ssr.ScoreP),
                    CreateJSubjectDetail("6.2.9", "學年學業成績", ssr.YearScore),
                    CreateJSubjectDetail("6.2.10", "學年及格", ssr.YearScoreP),
                    CreateJSubjectDetail("6.2.11", "是否採計學時", ssr.useCredit),
                    CreateJSubjectDetail("6.2.12", "質性文字描述", ssr.Text),
                    CreateJSubjectDetail("6.2.13", "DataKey", ""),
                    CreateJSubjectDetail("6.2.14", "備註(學生姓名)", ssr.Name),
                    CreateJSubjectDetail("6.2.15", "備註(資料當學期班級)", ssr.HisClassName),
                    CreateJSubjectDetail("6.2.16", "備註(資料當學期座號)", ssr.HisSeatNo + ""),
                    CreateJSubjectDetail("6.2.17", "備註(資料當學期學號)", ssr.HisStudentNumber),
                    CreateJSubjectDetail("6.2.18", "備註(資料次學期班級)", ssr.ClassName),
                    CreateJSubjectDetail("6.2.19", "備註(資料次學期座號)", ssr.SeatNo),
                    CreateJSubjectDetail("6.2.20", "備註(資料次學期學號)", ssr.StudentNumber)
                              };
                    jSubjectInfoList.Add(jInfo);
                }
            }
            SaveToDatabase(GetSemesterScoreSqlTemplate62(jSubjectInfoList));
        }
        public void SaveScores63(List<SubjectYearScoreRec108N> scores, int SchoolYear, int Semester)
        {
            var jSubjectInfoList = new List<JSubjectInfo>();
            foreach (var ssr in scores)
            {
                if (ssr.checkPass)
                {
                    if (!ssr.CodePass) //跳過課程代碼不可提交
                        continue;


                    var jInfo = CreateJSubjectInfoN(ssr, "6.3", "補考成績", SchoolYear, Semester);
                    jInfo.detail = new List<JSubjectDetail>
                  {
                    CreateJSubjectDetail("6.3.1", "身分證號", ssr.IDNumber),
                    CreateJSubjectDetail("6.3.2", "出生日期", ssr.Birthday),
                    CreateJSubjectDetail("6.3.3", "課程代碼", ssr.CourseCode),
                    CreateJSubjectDetail("6.3.4", "科目名稱", ssr.SubjectName),
                    CreateJSubjectDetail("6.3.5", "開課年級", ssr.GradeYear),
                    CreateJSubjectDetail("6.3.6", "第一學期開設修課節數", ssr.Credit1 == null ? "-1" : ssr.Credit1),
                    CreateJSubjectDetail("6.3.7", "第二學期開設修課節數", ssr.Credit2 == null ? "-1" : ssr.Credit2),
                    CreateJSubjectDetail("6.3.8", "第一次補考成績", ssr.ReScore),
                    CreateJSubjectDetail("6.3.9", "第一次及格", ssr.ScoreP),
                    CreateJSubjectDetail("6.3.10", "第二次補考成績", "-1"),
                    CreateJSubjectDetail("6.3.11", "第二次及格", "-1"),
                    CreateJSubjectDetail("6.3.12", "DataKey", ""),
                    CreateJSubjectDetail("6.3.13", "備註(學生姓名)", ssr.Name),
                    CreateJSubjectDetail("6.3.14", "備註(資料當學期班級)", ssr.HisClassName),
                    CreateJSubjectDetail("6.3.15", "備註(資料當學期座號)", ssr.HisSeatNo + ""),
                    CreateJSubjectDetail("6.3.16", "備註(資料當學期學號)", ssr.HisStudentNumber),
                    CreateJSubjectDetail("6.3.17", "備註(資料次學期班級)", ssr.ClassName),
                    CreateJSubjectDetail("6.3.18", "備註(資料次學期座號)", ssr.SeatNo),
                    CreateJSubjectDetail("6.3.19", "備註(資料次學期學號)", ssr.StudentNumber)
                  };
                    jSubjectInfoList.Add(jInfo);
                }


            }
            SaveToDatabase(GetSemesterScoreSqlTemplate63(jSubjectInfoList));
        }

        public void SaveScores64(List<SubjectScoreRec108> scores, int SchoolYear, int Semester)
        {
            var jSubjectInfoList = new List<JSubjectInfo>();
            foreach (var ssr in scores)
            {
                if (ssr.checkPass)
                {
                    if (!ssr.CodePass) //跳過課程代碼不可提交
                        continue;

                    var jInfo = CreateJSubjectInfo(ssr, "6.4", "轉學轉科成績", SchoolYear, Semester);
                    jInfo.detail = new List<JSubjectDetail>
                  {
                    CreateJSubjectDetail("6.4.1", "身分證號", ssr.IDNumber),
                    CreateJSubjectDetail("6.4.2", "出生日期", ssr.Birthday),
                    CreateJSubjectDetail("6.4.3", "對應學年度", ssr.SchoolYear),
                    CreateJSubjectDetail("6.4.4", "對應學期", ssr.Semester),
                    CreateJSubjectDetail("6.4.5", "課程代碼", ssr.CourseCode),
                    CreateJSubjectDetail("6.4.6", "科目名稱", ssr.SubjectName),
                    CreateJSubjectDetail("6.4.7", "對應年級", ssr.GradeYear),
                    CreateJSubjectDetail("6.4.8", "修課節數", ssr.Credit),
                    CreateJSubjectDetail("6.4.9", "身分別", ssr.StudType),
                    CreateJSubjectDetail("6.4.10", "抵免後成績", ssr.Score),
                    CreateJSubjectDetail("6.4.11", "成績及格", ssr.ScoreP),
                    CreateJSubjectDetail("6.4.12", "是否採計學時", ssr.useCredit),
                    CreateJSubjectDetail("6.4.13", "質性文字描述", ssr.Text),
                    CreateJSubjectDetail("6.4.14", "DataKey", ""),
                    CreateJSubjectDetail("6.4.15", "備註(學生姓名)", ssr.Name),
                    CreateJSubjectDetail("6.4.16", "備註(資料當學期班級)", ssr.HisClassName),
                    CreateJSubjectDetail("6.4.17", "備註(資料當學期座號)", ssr.HisSeatNo + ""),
                    CreateJSubjectDetail("6.4.18", "備註(資料當學期學號)", ssr.HisStudentNumber),
                    CreateJSubjectDetail("6.4.19", "備註(資料次學期班級)", ssr.ClassName),
                    CreateJSubjectDetail("6.4.20", "備註(資料次學期座號)", ssr.SeatNo),
                    CreateJSubjectDetail("6.4.21", "備註(資料次學期學號)", ssr.StudentNumber)
                  };
                    jSubjectInfoList.Add(jInfo);
                }
            }
            SaveToDatabase(GetSemesterScoreSqlTemplate64(jSubjectInfoList));
        }


        string GetSemesterScoreSqlTemplate(List<JSubjectInfo> jSubjectInfoList, string value)
        {
            if (jSubjectInfoList.Count == 0)
            {
                return "";
            }

            return string.Format(@"
        WITH row AS (
                -- 輸入資料（JSONB 格式）
                SELECT '{0}'::JSONB AS input
            ), row_expand AS (
                -- 將輸入展開成單一筆記錄（假設輸入是一個陣列的單一物件）    
                SELECT
                    ref_student_id,
                    serial_no,
                    name,
                    school_year,
                    semester,
                    subject,
                    subj_level,
                    val_obj->>'serialNo' AS detail_serial_no,
                    val_obj->>'name' AS detail_name,
                    val_obj->>'value' AS detail_value
                FROM
                    (
                        SELECT
                            (s.stu_obj->>'refStudentID')::BIGINT AS ref_student_id,
                            (s.stu_obj->>'serialNo') AS serial_no,
                            (s.stu_obj->>'name') AS name,
                            (s.stu_obj->>'schoolYear')::INT AS school_year,
                            (s.stu_obj->>'semester')::INT AS semester,
                            (s.stu_obj->>'subject') AS subject,
                            (s.stu_obj->>'subjectLevel')::INT AS subj_level,
                            JSONB_ARRAY_ELEMENTS(stu_obj->'detail') AS val_obj
                        FROM
                            (
                                SELECT JSONB_ARRAY_ELEMENTS(row.input) AS stu_obj
                                FROM row
                            ) AS s
                    ) AS s2
            ), detail_template AS (
                -- 定義完整的 20 個 detail 項目模板
                SELECT 
                    * 
                FROM 
                        (
                            SELECT DISTINCT 
                                ref_student_id,
                                serial_no,
                                name,
                                school_year,
                                semester,
                                subject,
                                subj_level
                            FROM row_expand
                        ) AS re
                        CROSS JOIN (
                            VALUES 
                 {1}
                        ) AS t (detail_serial_no, detail_name, detail_value)
            ), current_data_expand AS (
                -- 搜尋資料庫中的現有資料並展開 detail
                SELECT 
                    student_learning_history.id AS current_id,
                    student_learning_history.ref_student_id,
                    student_learning_history.serial_no,
                    student_learning_history.name,
                    student_learning_history.school_year,
                    student_learning_history.semester,
                    student_learning_history.subject,
                    student_learning_history.subj_level,
                    jsonb_array_elements(student_learning_history.detail)->>'serialNo' AS detail_serial_no,
                    jsonb_array_elements(student_learning_history.detail)->>'name' AS detail_name,
                    jsonb_array_elements(student_learning_history.detail)->>'value' AS detail_value
                FROM 
                    (
                        SELECT DISTINCT 
                            ref_student_id,
                            serial_no,
                            school_year,
                            semester,
                            subject,
                            subj_level
                        FROM row_expand
                    ) AS re INNER JOIN student_learning_history
                        ON student_learning_history.ref_student_id = re.ref_student_id
                        AND student_learning_history.serial_no = re.serial_no
                        AND student_learning_history.school_year = re.school_year
                        AND student_learning_history.semester = re.semester
                        AND student_learning_history.subject = re.subject
                        AND (
                            student_learning_history.subj_level = re.subj_level 
                            OR (
                                student_learning_history.subj_level IS NULL AND re.subj_level IS NULL
                            )
                        )
            ), summarize_row AS (
                SELECT DISTINCT
                    ref_student_id,
                    serial_no,
                    name,
                    school_year,
                    semester,
                    subject,
                    subj_level,
                    detail_serial_no
                FROM
                    (
                        SELECT 
                            ref_student_id,
                            serial_no,
                            name,
                            school_year,
                            semester,
                            subject,
                            subj_level,
                            detail_serial_no
                        FROM row_expand
            
                        UNION ALL
            
                        SELECT 
                            ref_student_id,
                            serial_no,
                            name,
                            school_year,
                            semester,
                            subject,
                            subj_level,
                            detail_serial_no
                        FROM current_data_expand
            
                        UNION ALL
            
                        SELECT 
                            ref_student_id,
                            serial_no,
                            name,
                            school_year,
                            semester,
                            subject,
                            subj_level,
                            detail_serial_no
                        FROM detail_template
                    ) AS s
            ), detail_merge AS (
                SELECT
                    MAX(current_data_expand.current_id) AS current_id,
                    summarize_row.ref_student_id,
                    summarize_row.serial_no,
                    summarize_row.name,
                    summarize_row.school_year,
                    summarize_row.semester,
                    summarize_row.subject,
                    summarize_row.subj_level,
                    JSONB_AGG(
                        JSONB_BUILD_OBJECT(
                            'serialNo', summarize_row.detail_serial_no,
                            'name', COALESCE(
                                    row_expand.detail_name,
                                    current_data_expand.detail_name,
                                    detail_template.detail_name
                                ),
                            'value', COALESCE(
                                    row_expand.detail_value,
                                    current_data_expand.detail_value,
                                    detail_template.detail_value
                                )
                        )
                    ) AS merged_detail
                FROM
                    summarize_row
                    LEFT OUTER JOIN row_expand
                        ON row_expand.ref_student_id = summarize_row.ref_student_id
                        AND row_expand.serial_no = summarize_row.serial_no
                        AND row_expand.school_year = summarize_row.school_year
                        AND row_expand.semester = summarize_row.semester
                        AND row_expand.subject = summarize_row.subject
                        AND (
                            row_expand.subj_level = summarize_row.subj_level 
                            OR (
                                row_expand.subj_level IS NULL 
                                AND summarize_row.subj_level IS NULL
                            )
                        )
                        AND row_expand.detail_serial_no = summarize_row.detail_serial_no
                    LEFT OUTER JOIN current_data_expand
                        ON current_data_expand.ref_student_id = summarize_row.ref_student_id
                        AND current_data_expand.serial_no = summarize_row.serial_no
                        AND current_data_expand.school_year = summarize_row.school_year
                        AND current_data_expand.semester = summarize_row.semester
                        AND current_data_expand.subject = summarize_row.subject
                        AND (
                            current_data_expand.subj_level = summarize_row.subj_level 
                            OR (
                                current_data_expand.subj_level IS NULL 
                                AND summarize_row.subj_level IS NULL
                            )
                        )
                        AND current_data_expand.detail_serial_no = summarize_row.detail_serial_no
                    LEFT OUTER JOIN detail_template
                        ON detail_template.ref_student_id = summarize_row.ref_student_id
                        AND detail_template.serial_no = summarize_row.serial_no
                        AND detail_template.school_year = summarize_row.school_year
                        AND detail_template.semester = summarize_row.semester
                        AND detail_template.subject = summarize_row.subject
                        AND (
                            detail_template.subj_level = summarize_row.subj_level 
                            OR (
                                detail_template.subj_level IS NULL 
                                AND summarize_row.subj_level IS NULL
                            )
                        )
                        AND detail_template.detail_serial_no = summarize_row.detail_serial_no
                GROUP BY
                    summarize_row.ref_student_id,
                    summarize_row.serial_no,
                    summarize_row.name,
                    summarize_row.school_year,
                    summarize_row.semester,
                    summarize_row.subject,
                    summarize_row.subj_level
            ), data_update AS (
                -- 更新現有資料
                UPDATE student_learning_history
                SET detail = detail_merge.merged_detail
                FROM detail_merge
                WHERE
                    student_learning_history.id = detail_merge.current_id
                    AND detail_merge.current_id IS NOT NULL
                RETURNING student_learning_history.id, 'updated'::TEXT AS action
            ), data_insert AS (
                -- 新增未建立資料
                INSERT INTO student_learning_history (
                    ref_student_id, 
                    serial_no, 
                    name, 
                    school_year, 
                    semester, 
                subject, 
                subj_level, 
                detail
            )
            SELECT 
                detail_merge.ref_student_id,
                detail_merge.serial_no,
                detail_merge.name,
                detail_merge.school_year,
                detail_merge.semester,
                detail_merge.subject,
                detail_merge.subj_level,
                detail_merge.merged_detail
            FROM
                detail_merge
            WHERE 
                detail_merge.current_id IS NULL
            RETURNING id, 'inserted'::TEXT AS action
        )
        -- 回傳新增和更新的筆數
        SELECT 
            COUNT(CASE WHEN action = 'inserted' THEN 1 END) AS inserted_count,
            COUNT(CASE WHEN action = 'updated' THEN 1 END) AS updated_count
        FROM (
            SELECT action FROM data_update
            UNION ALL
            SELECT action FROM data_insert
        ) AS upsert_result;
        ", JsonConvert.SerializeObject(jSubjectInfoList), value);
        }


        string GetSemesterScoreSqlTemplate42(List<JSubjectInfo> jSubjectInfoList)
        {
            string value = @"
               ('4.2.1', '身分證號', ''),
                ('4.2.2', '出生日期', ''),
                ('4.2.3', '課程代碼', ''),
                ('4.2.4', '科目名稱', ''),
                ('4.2.5', '開課年級', ''),
                ('4.2.6', '修課學分', ''),
                ('4.2.7', '學期學業成績', ''),
                ('4.2.8', '成績及格', ''),
                ('4.2.9', '補考成績', ''),
                ('4.2.10', '補考及格', ''),
                ('4.2.11', '是否採計學分', ''),
                ('4.2.12', '質性文字描述', ''),
                ('4.2.13', 'DataKey', ''),
                ('4.2.14', '備註(學生姓名)', ''),
                ('4.2.15', '備註(資料當學期班級)', ''),
                ('4.2.16', '備註(資料當學期座號)', ''),
                ('4.2.17', '備註(資料當學期學號)', ''),
                ('4.2.18', '備註(資料次學期班級)', ''),
                ('4.2.19', '備註(資料次學期座號)', ''),
                ('4.2.20', '備註(資料次學期學號)', '')
";


            return GetSemesterScoreSqlTemplate(jSubjectInfoList, value);
        }

        string GetSemesterScoreSqlTemplate43(List<JSubjectInfo> jSubjectInfoList)
        {
            string value = @"
            ('4.3.1', '身分證號', ''),
			('4.3.2', '出生日期', ''),
			('4.3.3', '應修課學年度', ''),
			('4.3.4', '應修課學期', ''),
			('4.3.5', '課程代碼', ''),
			('4.3.6', '科目名稱', ''),
			('4.3.7', '開課年級', ''),
			('4.3.8', '修課學分', ''),
			('4.3.9', '補修成績', ''),
			('4.3.10', '補修及格', ''),
			('4.3.11', '補考成績', ''),
			('4.3.12', '補考及格', ''),
			('4.3.13', '補修方式', ''),
			('4.3.14', '是否採計學分', ''),
			('4.3.15', '質性文字描述', ''),
			('4.3.16', 'DataKey', ''),
			('4.3.17', '備註(學生姓名)', ''),
			('4.3.18', '備註(資料當學期班級)', ''),
			('4.3.19', '備註(資料當學期座號)', ''),
			('4.3.20', '備註(資料當學期學號)', ''),
			('4.3.21', '備註(資料次學期班級)', ''),
			('4.3.22', '備註(資料次學期座號)', '')
";

            return GetSemesterScoreSqlTemplate(jSubjectInfoList, value);
        }

        string GetSemesterScoreSqlTemplate44(List<JSubjectInfo> jSubjectInfoList)
        {
            string value = @"
            ('4.4.1', '身分證號', ''),
			('4.4.2', '出生日期', ''),
			('4.4.3', '對應學年度', ''),
			('4.4.4', '對應學期', ''),
			('4.4.5', '課程代碼', ''),
			('4.4.6', '科目名稱', ''),
			('4.4.7', '對應年級', ''),
			('4.4.8', '修課學分', ''),
			('4.4.9', '身分別', ''),
			('4.4.10', '抵免後成績', ''),
			('4.4.11', '成績及格', ''),
			('4.4.12', '是否採計學分', ''),
			('4.4.13', '質性文字描述', ''),
			('4.4.14', 'DataKey', ''),
			('4.4.15', '備註(學生姓名)', ''),
			('4.4.16', '備註(資料當學期班級)', ''),
			('4.4.17', '備註(資料當學期座號)', ''),
			('4.4.18', '備註(資料當學期學號)', ''),
			('4.4.19', '備註(資料次學期班級)', ''),
			('4.4.20', '備註(資料次學期座號)', ''),
			('4.4.21', '備註(資料次學期學號)', '')
";
            return GetSemesterScoreSqlTemplate(jSubjectInfoList, value);
        }

        string GetSemesterScoreSqlTemplate52(List<JSubjectInfo> jSubjectInfoList)
        {
            string value = @"
            ('5.2.1', '身分證號', ''),
            ('5.2.2', '出生日期', ''),
            ('5.2.3', '原修課學年度', ''),
            ('5.2.4', '原修課學期', ''),
            ('5.2.5', '原修課課程代碼', ''),
            ('5.2.6', '原修課科目名稱', ''),
            ('5.2.7', '原修課開課年級', ''),
            ('5.2.8', '原修課修課學分', ''),
            ('5.2.9', '重修方式', ''),
            ('5.2.10', '重修成績', ''),
            ('5.2.11', '重修及格', ''),
            ('5.2.12', '質性文字描述', ''),
            ('5.2.13', 'DataKey', ''),
            ('5.2.14', '備註(學生姓名)', ''),
            ('5.2.15', '備註(資料當學期班級)', ''),
            ('5.2.16', '備註(資料當學期座號)', ''),
            ('5.2.17', '備註(資料當學期學號)', ''),
            ('5.2.18', '備註(資料次學期班級)', ''),
            ('5.2.19', '備註(資料次學期座號)', ''),
            ('5.2.20', '備註(資料次學期學號)', '')
";

            return GetSemesterScoreSqlTemplate(jSubjectInfoList, value);
        }

        string GetSemesterScoreSqlTemplate53(List<JSubjectInfo> jSubjectInfoList)
        {
            string value = @"
            ('5.3.1', '身分證號', ''),
            ('5.3.2', '出生日期', ''),
            ('5.3.3', '課程代碼', ''),
            ('5.3.4', '科目名稱', ''),
            ('5.3.5', '開課年級', ''),
            ('5.3.6', '修課學分', ''),
            ('5.3.7', '再次修習成績', ''),
            ('5.3.8', '再次修習成績及格', ''),
            ('5.3.9', '補考成績', ''),
            ('5.3.10', '補考及格', ''),
            ('5.3.11', '重讀成績', ''),
            ('5.3.12', '成績及格', ''),
            ('5.3.13', '重讀註記', ''),
            ('5.3.14', '是否採計學分', ''),
            ('5.3.15', '質性文字描述', ''),
            ('5.3.16', 'DataKey', ''),
            ('5.3.17', '備註(學生姓名)', ''),
            ('5.3.18', '備註(資料當學期班級)', ''),
            ('5.3.19', '備註(資料當學期座號)', ''),
            ('5.3.20', '備註(資料當學期學號)', ''),
            ('5.3.21', '備註(資料次學期班級)', ''),
            ('5.3.22', '備註(資料次學期座號)', ''),
            ('5.3.23', '備註(資料次學期學號)', '')
";
            return GetSemesterScoreSqlTemplate(jSubjectInfoList, value);
        }

        string GetSemesterScoreSqlTemplate62(List<JSubjectInfo> jSubjectInfoList)
        {
            string value = @"
            ('6.2.1', '身分證號', ''),
            ('6.2.2', '出生日期', ''),
            ('6.2.3', '課程代碼', ''),
            ('6.2.4', '科目名稱', ''),
            ('6.2.5', '開課年級', ''),
            ('6.2.6', '修課節數', ''),
            ('6.2.7', '學期學業成績', ''),
            ('6.2.8', '成績及格', ''),
            ('6.2.9', '學年學業成績', ''),
            ('6.2.10', '學年及格', ''),
            ('6.2.11', '是否採計學時', ''),
            ('6.2.12', '質性文字描述', ''),
            ('6.2.13', 'DataKey', ''),
            ('6.2.14', '備註(學生姓名)', ''),
            ('6.2.15', '備註(資料當學期班級)', ''),
            ('6.2.16', '備註(資料當學期座號)', ''),
            ('6.2.17', '備註(資料當學期學號)', ''),
            ('6.2.18', '備註(資料次學期班級)', ''),
            ('6.2.19', '備註(資料次學期座號)', ''),
            ('6.2.20', '備註(資料次學期學號)', '')
";


            return GetSemesterScoreSqlTemplate(jSubjectInfoList, value);
        }

        string GetSemesterScoreSqlTemplate63(List<JSubjectInfo> jSubjectInfoList)
        {
            string value = @"
            ('6.3.1', '身分證號', ''),
            ('6.3.2', '出生日期', ''),
            ('6.3.3', '課程代碼', ''),
            ('6.3.4', '科目名稱', ''),
            ('6.3.5', '開課年級', ''),
            ('6.3.6', '第一學期開設修課節數', ''),
            ('6.3.7', '第二學期開設修課節數', ''),
            ('6.3.8', '第一次補考成績', ''),
            ('6.3.9', '第一次及格', ''),
            ('6.3.10', '第二次補考成績', ''),
            ('6.3.11', '第二次及格', ''),
            ('6.3.12', 'DataKey', ''),
            ('6.3.13', '備註(學生姓名)', ''),
            ('6.3.14', '備註(資料當學期班級)', ''),
            ('6.3.15', '備註(資料當學期座號)', ''),
            ('6.3.16', '備註(資料當學期學號)', ''),
            ('6.3.17', '備註(資料次學期班級)', ''),
            ('6.3.18', '備註(資料次學期座號)', ''),
            ('6.3.19', '備註(資料次學期學號)', '')
";

            return GetSemesterScoreSqlTemplate(jSubjectInfoList, value);
        }

        string GetSemesterScoreSqlTemplate64(List<JSubjectInfo> jSubjectInfoList)
        {
            string value = @"
            ('6.4.1', '身分證號', ''),
            ('6.4.2', '出生日期', ''),
            ('6.4.3', '對應學年度', ''),
            ('6.4.4', '對應學期', ''),
            ('6.4.5', '課程代碼', ''),
            ('6.4.6', '科目名稱', ''),
            ('6.4.7', '對應年級', ''),
            ('6.4.8', '修課節數', ''),
            ('6.4.9', '身分別', ''),
            ('6.4.10', '抵免後成績', ''),
            ('6.4.11', '成績及格', ''),
            ('6.4.12', '是否採計學時', ''),
            ('6.4.13', '質性文字描述', ''),
            ('6.4.14', 'DataKey', ''),
            ('6.4.15', '備註(學生姓名)', ''),
            ('6.4.16', '備註(資料當學期班級)', ''),
            ('6.4.17', '備註(資料當學期座號)', ''),
            ('6.4.18', '備註(資料當學期學號)', ''),
            ('6.4.19', '備註(資料次學期班級)', ''),
            ('6.4.20', '備註(資料次學期座號)', ''),
            ('6.4.21', '備註(資料次學期學號)', '')    
";
            return GetSemesterScoreSqlTemplate(jSubjectInfoList, value);
        }

    }
}
