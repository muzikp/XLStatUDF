/// <summary>
/// Implements the Wilcoxon signed-rank test for paired data.
/// </summary>
namespace XLStatUDF.Functions.Tests
{
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class WilcoxonPaired
    {
        [ExcelFunction(Name = "WILCOXON.PAIRED", Description = "Wilcoxonův párový znaménkový test", Category = FunctionCategories.Tests)]
        public static object[,] Run(
            [ExcelArgument(Name = "x", Description = "První měření")] object xValues,
            [ExcelArgument(Name = "y", Description = "Druhé měření")] object yValues,
            [ExcelArgument(Name = "ma_záhlaví", Description = "Volitelně: 0=autodetect, 1=má záhlaví, 2=nemá záhlaví")] object? hasHeader = null,
            [ExcelArgument(Name = "alpha", Description = "Hladina významnosti")] double alpha = 0.05,
            [ExcelArgument(Name = "smer", Description = "Volitelně: 0=two, 1=left, 2=right")] object? direction = null)
        {
            if (!TestHelper.TryParseDirection(direction, out var parsedDirection))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!ArgumentHelper.TryParseHeaderMode(hasHeader, out var parsedHeaderMode))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!TestHelper.IsValidAlpha(alpha))
            {
                return SpillError(ExcelErrors.Num);
            }

            if (!DataHelper.TryReadPairedNumericVectors(xValues, yValues, out var x, out var y, out var error, parsedHeaderMode))
            {
                return SpillError(error!);
            }

            var differences = x.Zip(y, (first, second) => first - second)
                .Where(diff => Math.Abs(diff) > 1e-12)
                .ToArray();

            if (differences.Length < 1)
            {
                return SpillError(ExcelErrors.Count);
            }

            var absDiffs = differences.Select(Math.Abs).ToArray();
            var ranks = RankHelper.MidRank(absDiffs);
            var wPlus = differences
                .Select((diff, index) => diff > 0 ? ranks[index] : 0.0)
                .Sum();
            var wMinus = differences
                .Select((diff, index) => diff < 0 ? ranks[index] : 0.0)
                .Sum();
            var w = Math.Min(wPlus, wMinus);

            var n = differences.Length;
            var meanW = n * (n + 1) / 4.0;
            var varianceW = (n * (n + 1) * ((2.0 * n) + 1) / 24.0) - ComputeTieAdjustment(absDiffs);
            if (varianceW <= 0)
            {
                return SpillError(ExcelErrors.Num);
            }

            var centered = wPlus - meanW;
            var z = centered / Math.Sqrt(varianceW);
            if (parsedDirection == "two" && Math.Abs(z) > 0)
            {
                z = Math.Sign(z) * Math.Max(0.0, Math.Abs(centered) - 0.5) / Math.Sqrt(varianceW);
            }

            var p = TestHelper.PValueFromZ(z, parsedDirection);
            var critical = TestHelper.CriticalZ(alpha, parsedDirection);
            var r = Math.Abs(z) / Math.Sqrt(n);
            var medianDiff = StatisticsHelper.Median(differences);

            var rows = new List<object[]>();
            SpillBuilder.AddRow(rows, "n", n);
            SpillBuilder.AddRow(rows, "med(d)", medianDiff);
            SpillBuilder.AddRow(rows, "α", alpha);
            SpillBuilder.AddRow(rows, "W+", wPlus);
            SpillBuilder.AddRow(rows, "W-", wMinus);
            SpillBuilder.AddRow(rows, "W", w);
            SpillBuilder.AddRow(rows, "z", z);
            SpillBuilder.AddRow(rows, CriticalValues.LabelForDirection("z", parsedDirection), critical);
            SpillBuilder.AddRow(rows, "p", p);
            SpillBuilder.AddRow(rows, "r", r);
            return SpillBuilder.Build(rows);
        }

        private static double ComputeTieAdjustment(double[] absoluteDifferences)
        {
            return absoluteDifferences
                .GroupBy(value => value)
                .Select(group => group.Count())
                .Where(count => count > 1)
                .Sum(count => (Math.Pow(count, 3) - count) / 48.0);
        }

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
