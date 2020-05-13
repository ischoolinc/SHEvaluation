using K12.Data;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace 班級定期評量成績單_固定排名.Model
{
    /// <summary>
    /// 紀錄畫面上的參數 傳遞變數用
    /// </summary>
    class FormParam
    {
        /// <summary>
        /// 選取的classIDs
        /// </summary>
        internal  List<string>  SelectedClassIDs{ get; set; }
        /// <summary>
        /// 選取得 Class 資料 
        /// </summary>
        internal   List<ClassRecord> SelectedClasses { get; set; }
        /// <summary>
        /// 使用者選擇 的 年級
        /// </summary>
        internal List<int?> SelectGradeYears { get; set; }
    
        /// <summary>
        /// 列印試別ID 
        /// </summary>
        public int ExamID { get; set; }
        /// <summary>
        /// 參考試別ID (比較試別)
        /// </summary>
        public int CompareExamID { get; set; }
        /// <summary>
        /// 
        /// </summary>
        public Configure configure { get; set; }

        /// <summary>
        /// 建構子 
        /// </summary>
        public FormParam(string schoolYear ,string semester,int examID, List<ClassRecord> selectedClasses) 
        {
           
            this.SelectedClasses = selectedClasses;
            this.ExamID = examID;

            this.SelectedClassIDs = selectedClasses.Select(x => x.ID).ToList(); // 取得所選 班級 之IDs 
            this.SelectGradeYears = selectedClasses.Select(x => x.GradeYear).ToList(); // 取得所選 班級 之年級s

        }
    }
}
