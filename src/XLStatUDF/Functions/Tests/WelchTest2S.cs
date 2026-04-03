/// <summary>
/// Implementuje Welchův dvouvýběrový t-test pro dvě nezávislé skupiny.
/// </summary>
namespace XLStatUDF.Functions.Tests
{
    using System.Globalization;
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class WelchTest2S
    {
        [ExcelFunction(Name = "WELCH.TEST.2S.G", Description = "Welchův dvouvýběrový t-test nad groupovanými daty", Category = FunctionCategories.Tests)]
        public static object[,] Run(
            [ExcelArgument(Name = "kategorie", Description = "Štítky dvou skupin")] object categories,
            [ExcelArgument(Name = "hodnoty", Description = "Číselná pozorování")] object values,
            [ExcelArgument(Name = "ma_záhlaví", Description = "Volitelně: 0=autodetect, 1=má záhlaví, 2=nemá záhlaví")] object? hasHeader = null,
            [ExcelArgument(Name = "alpha", Description = "Hladina významnosti")] double alpha = 0.05,
            [ExcelArgument(Name = "smer", Description = "Volitelně: 0=two, 1=left, 2=right")] object? direction = null)
            => RunInternal(categories, values, direction, alpha, hasHeader);

        private static object[,] RunInternal(
            object categories,
            object values,
            object? direction,
            double alpha,
            object? hasHeader)
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
            var group1Name = ordered[0].Key;
            var group2Name = ordered[1].Key;
            var group1DisplayName = group1Name.ToUpperInvariant();
            var group2DisplayName = group2Name.ToUpperInvariant();
            var group1 = ordered[0].Value;
            var group2 = ordered[1].Value;

            if (group1.Count < 2 || group2.Count < 2)
            {
                return SpillError(ExcelErrors.Count);
            }

            var n1 = group1.Count;
            var n2 = group2.Count;
            var mean1 = StatisticsHelper.Mean(group1);
            var mean2 = StatisticsHelper.Mean(group2);
            var median1 = StatisticsHelper.Median(group1);
            var median2 = StatisticsHelper.Median(group2);
            var sd1 = StatisticsHelper.SampleStandardDeviation(group1);
            var sd2 = StatisticsHelper.SampleStandardDeviation(group2);
            var min1 = group1.Min();
            var min2 = group2.Min();
            var max1 = group1.Max();
            var max2 = group2.Max();

            var variancePerN1 = (sd1 * sd1) / n1;
            var variancePerN2 = (sd2 * sd2) / n2;
            var t = (mean1 - mean2) / Math.Sqrt(variancePerN1 + variancePerN2);
            var dfNumerator = Math.Pow(variancePerN1 + variancePerN2, 2);
            var dfDenominator = (Math.Pow(variancePerN1, 2) / (n1 - 1)) + (Math.Pow(variancePerN2, 2) / (n2 - 1));
            var df = dfNumerator / dfDenominator;
            var critical = TestHelper.CriticalT(alpha, df, parsedDirection);
            var p = TestHelper.PValueFromT(t, df, parsedDirection);

            var pooledVariance = (((n1 - 1) * sd1 * sd1) + ((n2 - 1) * sd2 * sd2)) / (n1 + n2 - 2.0);
            var pooledSd = Math.Sqrt(pooledVariance);
            var d = pooledSd == 0 ? 0.0 : (mean1 - mean2) / pooledSd;
            var imbalanceRatio = Math.Max(n1, n2) / (double)Math.Min(n1, n2);
            var r = imbalanceRatio > 1.5
                ? Math.Sign(t) * Math.Sqrt((t * t) / ((t * t) + df))
                : d / Math.Sqrt((d * d) + 4.0);

            var rows = new List<object[]>
            {
                new object[] { "Skupina", "n", "x̄", "med", "sₓ", "min", "max" },
                new object[] { group1DisplayName, n1, mean1, median1, sd1, min1, max1 },
                new object[] { group2DisplayName, n2, mean2, median2, sd2, min2, max2 },
                new object[] { "", "", "", "", "", "", "" },
                new object[] { "Výsledky testu", "", "", "", "", "", "" }
            };

            rows.Add(["α", alpha, "", "", "", "", ""]);
            rows.Add(["t", t, "", "", "", "", ""]);
            rows.Add(["df", df, "", "", "", "", ""]);
            rows.Add([CriticalValues.LabelForDirection("t", parsedDirection), critical, "", "", "", "", ""]);
            rows.Add(["p", p, "", "", "", "", ""]);
            rows.Add(["d", d, "", "", "", "", ""]);
            rows.Add(["r", r, "", "", "", "", ""]);
            return SpillBuilder.Build(rows);
        }

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
