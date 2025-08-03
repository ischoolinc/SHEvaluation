using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SmartSchool.Evaluation.Process.Wizards.LearningHistory
{
    public class SubjectScoreRec108
    {
        public string StudentID { get; set; }

        /// <summary>
        /// 身分證號
        /// </summary>
        public string IDNumber { get; set; }

        /// <summary>
        /// 出生日期
        /// </summary>
        public string Birthday { get; set; }

        /// <summary>
        /// 課程代碼
        /// </summary>
        public string CourseCode { get; set; }

        /// <summary>
        /// 科目名稱
        /// </summary>
        public string SubjectName { get; set; }

        /// <summary>
        /// 科目級別
        /// </summary>
        public string SubjectLevel { get; set; }

        /// <summary>
        /// 開課年級
        /// </summary>
        public string GradeYear { get; set; }

        /// <summary>
        /// 修課學分
        /// </summary>
        public string Credit { get; set; }

        /// <summary>
        /// 學期學業成績
        /// </summary>
        public string Score { get; set; }

        /// <summary>
        /// 成績及格
        /// </summary>
        public string ScoreP { get; set; }

        /// <summary>
        /// 學年學業成績
        /// </summary>
        public string YearScore { get; set; }

        /// <summary>
        /// 學年成績及格
        /// </summary>
        public string YearScoreP { get; set; }

        /// <summary>
        /// 補考成績
        /// </summary>
        public string ReScore { get; set; }

        /// <summary>
        /// 補考及格
        /// </summary>
        public string ReScoreP { get; set; }


        /// <summary>
        /// 重修成績
        /// </summary>
        public string ReAScore { get; set; }

        /// <summary>
        /// 重修及格
        /// </summary>
        public string ReAScoreP { get; set; }

        /// <summary>
        /// 重修方式
        /// </summary>
        public string ReAScoreType { get; set; }

        /// <summary>
        /// 重讀成績（5.3 專用）
        /// </summary>
        public string RepeatScore { get; set; }

        /// <summary>
        /// 重讀成績及格（1/0/-1，5.3 專用）
        /// </summary>
        public string RepeatScoreP { get; set; }

        /// <summary>
        /// 重讀註記（列抵請填1，再次修習請填2。）
        /// </summary>
        public string RepeatMemo { get; set; }

        /// <summary>
        /// 是否採計學分
        /// </summary>
        public string useCredit { get; set; }

        /// <summary>
        /// 質性文字描述
        /// </summary>
        public string Text { get; set; }
        /// <summary>
        /// 應修課學年度
        /// </summary>
        public string SchoolYear { get; set; }

        /// <summary>
        /// 應修課學期
        /// </summary>
        public string Semester { get; set; }

        /// <summary>
        /// 補修成績
        /// </summary>
        public string ScScore { get; set; }

        /// <summary>
        /// 補修及格
        /// </summary>
        public string ScScoreP { get; set; }

        /// <summary>
        /// 補修方式
        /// </summary>
        public string ScScoreType { get; set; }

        /// <summary>
        /// 對應學年度
        /// </summary>
        public string mSchoolYear { get; set; }

        /// <summary>
        /// 對應學期
        /// </summary>
        public string mSemester { get; set; }

        /// <summary>
        /// 對應年級
        /// </summary>
        public string mGradeYear { get; set; }
        /// <summary>
        /// 身分別
        /// </summary>
        public string StudType { get; set; }


        /// <summary>
        /// 是否補修成績
        /// </summary>
        public bool isScScore { get; set; }

        /// <summary>
        /// 重讀註記
        /// </summary>
        public string ReStudMark { get; set; }

        /// <summary>
        /// 檢查資料是否通過
        /// </summary>
        public bool checkPass { get; set; }

        /// <summary>
        /// 檢查課程代碼是否可提交成績
        /// </summary>
        public bool CodePass { get; set; }

        /// <summary>
        /// 備註(學生姓名)
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// 備註(資料當學期班級) 
        /// 學期歷程班級
        /// </summary>
        public string HisClassName { get; set; }

        /// <summary>
        /// 備註(資料當學期座號)
        /// 學期歷程座號
        /// </summary>
        public int? HisSeatNo { get; set; }

        /// <summary>
        /// 備註(資料當學期學號)
        /// 學期歷程學號
        /// </summary>
        public string HisStudentNumber { get; set; }

        /// <summary>
        /// 備註(資料次學期班級)
        /// 現在班級
        /// </summary>
        public string ClassName { get; set; }

        /// <summary>
        /// 備註(資料次學期座號)
        /// 現在座號
        /// </summary>
        public string SeatNo { get; set; }

        /// <summary>
        /// 備註(資料次學期學號)
        /// 現在學號
        /// </summary>
        public string StudentNumber { get; set; }

        /// <summary>
        /// 補修學年度
        /// </summary>
        public string ScScoreSchoolYear { get; set; }
    }
}
