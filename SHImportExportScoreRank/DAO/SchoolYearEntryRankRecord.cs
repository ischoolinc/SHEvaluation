namespace SHImportExportScoreRank.DAO
{
    public class SchoolYearEntryRankRecord
    {
        public long ID { get; set; }
        public string StudentID { get; set; }
        public int? SchoolYear { get; set; }
        public int? GradeYear { get; set; }
        public string RankName { get; set; }
        public int? Rank { get; set; }

        public string ScoreType { get; set; }
        public string ScoreItem { get; set; }
        public string RankType { get; set; }
        public string CreateTime { get; set; }
        public string CreateMethod { get; set; }

        public string StudentNumber { get; set; }
        public string SeatNo { get; set; }
        public string DepartmentName { get; set; }
        public string StudentName { get; set; }
    }
}
