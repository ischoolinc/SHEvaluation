using System;

namespace SmartSchool.Evaluation
{
    internal static class Program
    {
        private static int _failures;

        private static void Main()
        {
            CheckRequiredTable();
            CheckExtraCases();
            decimal? history = RepeatSubjectOriginalScore.ChooseRepeatScore(82m, 95m);
            Expect(Original("70", true, 99m, history), 95m, "指定優先於課程，封存最高");
            history = RepeatSubjectOriginalScore.ChooseRepeatScore(95m, 82m);
            Expect(Original("70", true, 99m, history), 95m, "既有學期最高");
            Expect(Original("", true, 99m, history), 99m, "未指定，課程最高");
            CheckPassStandardAndRounding();

            if (_failures > 0)
            {
                Console.Error.WriteLine("FAILED " + _failures);
                Environment.Exit(1);
            }

            Console.WriteLine("PASSED");
        }

        private static void CheckRequiredTable()
        {
            ExpectRow("100", true, 90m, 82m, 60m, 100m, 100m, "1", 100m, "1", "指定 100");
            ExpectRow("80", true, 90m, 82m, 60m, 82m, 80m, "1", 82m, "1", "指定 80");
            ExpectRow("70", true, 90m, 82m, 60m, 82m, 70m, "1", 82m, "1", "指定 70");
            ExpectRow("0", true, 90m, 82m, 60m, 82m, 0m, "0", 82m, "1", "指定 0");
            ExpectRow("", true, 90m, 82m, 60m, 90m, 90m, "1", 90m, "1", "指定空白且課程 90");
            ExpectRow(null, true, 90m, 82m, 60m, 90m, 90m, "1", 90m, "1", "指定未提供且課程 90");
            ExpectRow("", true, 75m, 82m, 60m, 82m, 75m, "1", 82m, "1", "指定空白且課程 75");
        }

        private static void CheckExtraCases()
        {
            ExpectRow("ABC", true, 90m, 82m, 60m, 90m, 90m, "1", 90m, "1", "指定格式錯誤");
            ExpectRow("   ", true, 90m, 82m, 60m, 90m, 90m, "1", 90m, "1", "指定只有空白");
            ExpectRow("80", true, 90m, null, 60m, 80m, 80m, "1", 80m, "1", "沒有歷史成績");
            ExpectRow("70", false, 0m, null, 60m, 70m, 70m, "1", 70m, "1", "只有指定總成績");
            ExpectRow("0", false, 0m, 82m, 60m, 82m, 0m, "0", 82m, "1", "只有指定 0 且有歷史");
            ExpectRow(null, false, 0m, 82m, 60m, 82m, null, "0", 82m, "1", "沒有本次有效成績");
            ExpectRow("", false, 0m, null, 60m, null, null, "0", null, "0", "兩種成績都沒有");

            decimal? firstOriginal = Original("80", true, 90m, 82m);
            decimal? secondOriginal = Original("80", true, 90m, 82m);
            Expect(secondOriginal, firstOriginal, "相同輸入連續計算兩次");
        }

        private static void CheckPassStandardAndRounding()
        {
            ExpectRow("50", true, 90m, 82m, 70m, 82m, 50m, "0", 82m, "1", "及格標準 70 且指定較低");
            ExpectRow("", true, 75m, 82m, 70m, 82m, 75m, "1", 82m, "1", "及格標準 70 且未指定");

            decimal? original = Original("80.126", true, 90m, 82.126m, RoundAwayFromZero2);
            decimal? repeat = Repeat("80.126", true, 90m, 82.126m, RoundAwayFromZero2);
            Expect(original, 82.126m, "指定總成績先進位再與歷史擇優");
            Expect(repeat, 82.126m, "重讀成績保留未再進位的歷史較高分");

            decimal? courseOriginal = Original("", true, 90.126m, 90m, RoundAwayFromZero2);
            Expect(courseOriginal, 90.13m, "未指定時課程總成績先進位再與歷史比較");
        }

        private static void ExpectRow(
            string designateText,
            bool hasFinalScore,
            decimal finalScore,
            decimal? history,
            decimal passStandard,
            decimal? expectedOriginal,
            decimal? expectedScore,
            string expectedScoreP,
            decimal? expectedRepeat,
            string expectedRepeatP,
            string name)
        {
            bool hasDesignate = RepeatSubjectOriginalScore.HasValidDesignate(designateText);
            decimal? score = RepeatSubjectOriginalScore.ResolveEffectiveScore(
                designateText,
                hasFinalScore,
                finalScore,
                Identity);
            decimal? original = RepeatSubjectOriginalScore.ChooseOriginalScore(hasDesignate, score, history);
            decimal? repeat = RepeatSubjectOriginalScore.ChooseRepeatScore(score, history);

            Expect(original, expectedOriginal, name + " 原始成績");
            Expect(score, expectedScore, name + " Score");
            ExpectText(RepeatSubjectOriginalScore.PassFlag(score, passStandard), expectedScoreP, name + " ScoreP");
            Expect(repeat, expectedRepeat, name + " RepeatScore");
            ExpectText(RepeatSubjectOriginalScore.PassFlag(repeat, passStandard), expectedRepeatP, name + " RepeatScoreP");
        }

        private static decimal? Original(
            string designateText,
            bool hasFinalScore,
            decimal finalScore,
            decimal? history)
        {
            return Original(designateText, hasFinalScore, finalScore, history, Identity);
        }

        private static decimal? Original(
            string designateText,
            bool hasFinalScore,
            decimal finalScore,
            decimal? history,
            Func<decimal, decimal> round)
        {
            decimal? score = RepeatSubjectOriginalScore.ResolveEffectiveScore(
                designateText,
                hasFinalScore,
                finalScore,
                round);
            return RepeatSubjectOriginalScore.ChooseOriginalScore(
                RepeatSubjectOriginalScore.HasValidDesignate(designateText),
                score,
                history);
        }

        private static decimal? Repeat(
            string designateText,
            bool hasFinalScore,
            decimal finalScore,
            decimal? history,
            Func<decimal, decimal> round)
        {
            decimal? score = RepeatSubjectOriginalScore.ResolveEffectiveScore(
                designateText,
                hasFinalScore,
                finalScore,
                round);
            return RepeatSubjectOriginalScore.ChooseRepeatScore(score, history);
        }

        private static decimal Identity(decimal score)
        {
            return score;
        }

        private static decimal RoundAwayFromZero2(decimal score)
        {
            return decimal.Round(score, 2, MidpointRounding.AwayFromZero);
        }

        private static void Expect(decimal? actual, decimal? expected, string name)
        {
            if (actual == expected)
                return;

            _failures++;
            Console.Error.WriteLine(name + " expected " + Format(expected) + " actual " + Format(actual));
        }

        private static void ExpectText(string actual, string expected, string name)
        {
            if (actual == expected)
                return;

            _failures++;
            Console.Error.WriteLine(name + " expected " + expected + " actual " + actual);
        }

        private static string Format(decimal? value)
        {
            return value.HasValue ? value.Value.ToString() : "(null)";
        }
    }
}
