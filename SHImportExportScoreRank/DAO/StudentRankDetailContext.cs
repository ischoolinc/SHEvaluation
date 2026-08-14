namespace SHImportExportScoreRank.DAO
{
    public class StudentRankDetailContext
    {
        public string StudentId { get; set; }
        public int? SchoolYear { get; set; }
        public int? GradeYear { get; set; }
        public string ScoreType { get; set; }
        public string ScoreItem { get; set; }
        public string CreateTime { get; set; }
        public string CreateMethod { get; set; }
    }
}
