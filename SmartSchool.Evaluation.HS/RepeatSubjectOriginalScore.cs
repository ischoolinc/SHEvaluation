using System;

namespace SmartSchool.Evaluation
{
    /// <summary>
    /// 再次修習：原始成績與 5.3 重讀成績分開計算。
    /// </summary>
    public static class RepeatSubjectOriginalScore
    {
        public static bool HasValidDesignate(string designateText)
        {
            decimal designateScore;
            return designateText != null && decimal.TryParse(designateText.Trim(), out designateScore);
        }

        public static decimal? ResolveEffectiveScore(
            string designateText,
            bool hasFinalScore,
            decimal finalScore,
            Func<decimal, decimal> round)
        {
            if (round == null)
                throw new ArgumentNullException("round");

            decimal designateScore;
            if (designateText != null && decimal.TryParse(designateText.Trim(), out designateScore))
                return round(designateScore);

            if (hasFinalScore)
                return round(finalScore);

            return null;
        }

        public static decimal? ChooseOriginalScore(
            bool hasDesignate,
            decimal? effectiveScore,
            decimal? historyScore)
        {
            return Higher(effectiveScore, historyScore);
        }

        public static decimal? ChooseRepeatScore(decimal? effectiveScore, decimal? historyScore)
        {
            return Higher(effectiveScore, historyScore);
        }

        public static string PassFlag(decimal? score, decimal passStandard)
        {
            if (!score.HasValue)
                return "0";

            return score.Value >= passStandard ? "1" : "0";
        }

        private static decimal? Higher(decimal? effectiveScore, decimal? historyScore)
        {
            if (effectiveScore.HasValue && historyScore.HasValue)
                return historyScore.Value > effectiveScore.Value ? historyScore.Value : effectiveScore.Value;

            if (effectiveScore.HasValue)
                return effectiveScore;

            return historyScore;
        }
    }
}
