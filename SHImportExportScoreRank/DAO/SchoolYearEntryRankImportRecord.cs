namespace SHImportExportScoreRank.DAO
{
    public class SchoolYearEntryRankImportRecord
    {
        public long RefStudentID { get; set; }
        public int SchoolYear { get; set; }
        public int GradeYear { get; set; }
        public string RankName { get; set; }
        public int Rank { get; set; }

        /// <summary>
        /// 僅供錯誤訊息顯示，不寫入 rank_override。
        /// </summary>
        public string StudentNumber { get; set; }
    }
}
