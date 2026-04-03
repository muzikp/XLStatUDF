/// <summary>
/// Implements the Mann-Whitney U test for two independent groups.
/// </summary>
namespace XLStatUDF.Functions.Tests
{
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using XLStatUDF.Helpers;

    public static class MannWhitneyTest
    {
        [ExcelFunction(Name = "MANN.WHITNEY.G", Description = "Mann-Whitneyho test nad groupovanými daty", Category = FunctionCategories.Tests)]
        public static object[,] Run(
            [ExcelArgument(Name = "kategorie", Description = "Štítky dvou skupin")] object categories,
            [ExcelArgument(Name = "hodnoty", Description = "Číselná pozorování")] object values,
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

            if (!GroupingHelper.TryReadGroups(categories, values, out var grouped, out var error, parsedHeaderMode))
            {
                return SpillError(error!);
            }

            if (grouped.Count != 2)
            {
                return SpillError(ExcelErrors.Count);
            }

            var ordered = grouped.OrderBy(pair => pair.Key, StringComparer.Ordinal).ToArray();
            var groupA = ordered[0];
            var groupB = ordered[1];

            if (groupA.Value.Count < 1 || groupB.Value.Count < 1)
            {
                return SpillError(ExcelErrors.Count);
            }

            var n1 = groupA.Value.Count;
            var n2 = groupB.Value.Count;
            var combined = groupA.Value
                .Select(value => (Group: 0, Value: value))
                .Concat(groupB.Value.Select(value => (Group: 1, Value: value)))
                .ToArray();
            var ranks = RankHelper.MidRank(combined.Select(item => item.Value).ToArray());
            var rankSumA = combined
                .Select((item, index) => (item.Group, Rank: ranks[index]))
                .Where(item => item.Group == 0)
                .Sum(item => item.Rank);

            var u1 = rankSumA - ((n1 * (n1 + 1)) / 2.0);
            var u2 = (n1 * n2) - u1;
            var uStatistic = parsedDirection switch
            {
                "left" => u1,
                "right" => u2,
                _ => Math.Min(u1, u2)
            };

            var meanU = (n1 * n2) / 2.0;
            var tieCorrection = ComputeTieCorrection(combined.Select(item => item.Value).ToArray());
            var varianceU = (n1 * n2 / 12.0) * (((n1 + n2) + 1.0) - tieCorrection);
            if (varianceU <= 0)
            {
                return SpillError(ExcelErrors.Num);
            }

            var continuity = parsedDirection == "two" ? 0.5 : 0.0;
            var signedU = parsedDirection switch
            {
                "left" => u1,
                "right" => u2,
                _ => uStatistic
            };
            var z = (signedU - meanU + continuity) / Math.Sqrt(varianceU);
            if (parsedDirection == "right")
            {
                z = -z;
            }

            var p = TestHelper.PValueFromZ(z, parsedDirection);
            var critical = TestHelper.CriticalZ(alpha, parsedDirection);
            var r = Math.Abs(z) / Math.Sqrt(n1 + n2);

            var rows = new List<object[]>
            {
                new object[] { "Skupina", "n", "x̄", "med", "sₓ", "min", "max" },
                new object[] { groupA.Key, n1, StatisticsHelper.Mean(groupA.Value), StatisticsHelper.Median(groupA.Value), StatisticsHelper.SampleStandardDeviation(groupA.Value), groupA.Value.Min(), groupA.Value.Max() },
                new object[] { groupB.Key, n2, StatisticsHelper.Mean(groupB.Value), StatisticsHelper.Median(groupB.Value), StatisticsHelper.SampleStandardDeviation(groupB.Value), groupB.Value.Min(), groupB.Value.Max() },
                new object[] { string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
                new object[] { "Výsledky testu", string.Empty, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
                new object[] { "α", alpha, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
                new object[] { "U", uStatistic, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
                new object[] { "U₁", u1, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
                new object[] { "U₂", u2, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
                new object[] { "z", z, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
                new object[] { CriticalValues.LabelForDirection("z", parsedDirection), critical, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
                new object[] { "p", p, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty },
                new object[] { "r", r, string.Empty, string.Empty, string.Empty, string.Empty, string.Empty }
            };

            return SpillBuilder.Build(rows);
        }

        private static double ComputeTieCorrection(double[] values)
        {
            var groups = values
                .GroupBy(value => value)
                .Select(group => group.Count())
                .Where(count => count > 1)
                .ToArray();

            if (groups.Length == 0)
            {
                return 0.0;
            }

            var n = values.Length;
            return groups.Sum(t => (Math.Pow(t, 3) - t)) / (n * (n - 1.0));
        }

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
