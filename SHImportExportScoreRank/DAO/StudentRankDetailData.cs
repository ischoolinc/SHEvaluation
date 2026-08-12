using System.Collections.Generic;

namespace SHImportExportScoreRank.DAO
{
    public class StudentRankDetailData
    {
        public int? SchoolYear { get; set; }
        public string ScoreType { get; set; }
        public string ScoreItem { get; set; }
        public string CreateType { get; set; }
        public string CreateTime { get; set; }
        public string BatchName { get; set; }
        public List<StudentRankDetailItem> Items { get; set; }
    }

    public class StudentRankDetailItem
    {
        public long RankOverrideId { get; set; }
        public string ScoreCategory { get; set; }
        public string RankMethod { get; set; }
        public string Score { get; set; }
        public string RankType { get; set; }
        public string RankName { get; set; }
        public int? Rank { get; set; }
        public int? MatrixCount { get; set; }
        public string PR { get; set; }
        public string Percentage { get; set; }

        public string RankDisplay
        {
            get
            {
                if (!Rank.HasValue)
                    return string.Empty;

                if (!MatrixCount.HasValue)
                    return Rank.Value.ToString();

                return Rank.Value + " / " + MatrixCount.Value;
            }
        }
    }
}
