/// <summary>
/// Implementuje jednofaktorovou ANOVA včetně Leveneho testu a základních post-hoc srovnání.
/// </summary>
namespace XLStatUDF.Functions.Tests
{
    using MathNet.Numerics.Distributions;
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class Anova
    {
        [ExcelFunction(Name = "ANOVA.G", Description = "[Testy] Jednofaktorová analýza rozptylu nad groupovanými daty", Category = "XLStatUDF")]
        public static object[,] Run(
            [ExcelArgument(Name = "kategorie", Description = "Štítky skupin")] object categories,
            [ExcelArgument(Name = "hodnoty", Description = "Číselná pozorování")] object values,
            [ExcelArgument(Name = "post_hoc", Description = "Volitelne: 0=none, 1=tukey, 2=bonferroni, 3=scheffe, 4=games-howell")] object? postHoc = null,
            [ExcelArgument(Name = "alpha", Description = "Hladina významnosti")] double alpha = 0.05,
            [ExcelArgument(Name = "ma_zahlavi", Description = "Volitelne: 0=autodetect, 1=ma zahlavi, 2=nema zahlavi")] object? hasHeader = null)
            => RunInternal(categories, values, postHoc, alpha, hasHeader);

        private static object[,] RunInternal(
            object categories,
            object values,
            object? postHoc,
            double alpha,
            object? hasHeader)
        {
            if (!TestHelper.IsValidAlpha(alpha))
            {
                return SpillError(ExcelErrors.Num);
            }

            if (!ArgumentHelper.TryParseHeaderMode(hasHeader, out var parsedHeaderMode))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!TryParsePostHoc(postHoc, out var parsedPostHoc))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!GroupingHelper.TryReadGroups(categories, values, out var grouped, out var error, parsedHeaderMode))
            {
                return SpillError(error!);
            }

            if (grouped.Count < 3)
            {
                return SpillError(ExcelErrors.Count);
            }

            var orderedGroups = grouped
                .OrderBy(pair => pair.Key, StringComparer.Ordinal)
                .Select(pair => (pair.Key, pair.Value))
                .ToArray();
            if (orderedGroups.Any(pair => pair.Value.Count < 2))
            {
                return SpillError(ExcelErrors.Count);
            }

            var totalN = orderedGroups.Sum(pair => pair.Value.Count);
            var grandMean = orderedGroups.SelectMany(pair => pair.Value).Average();
            var k = orderedGroups.Length;
            var dfBetween = k - 1;
            var dfWithin = totalN - k;
            var dfTotal = totalN - 1;

            var ssBetween = orderedGroups.Sum(pair => pair.Value.Count * Math.Pow(pair.Value.Average() - grandMean, 2));
            var ssWithin = orderedGroups.Sum(pair =>
            {
                var mean = pair.Value.Average();
                return pair.Value.Sum(value => Math.Pow(value - mean, 2));
            });
            var ssTotal = ssBetween + ssWithin;
            var msBetween = ssBetween / dfBetween;
            var msWithin = ssWithin / dfWithin;
            var f = msBetween / msWithin;
            var p = 1 - FisherSnedecor.CDF(dfBetween, dfWithin, f);
            var fCrit = FisherSnedecor.InvCDF(dfBetween, dfWithin, 1 - alpha);

            var (leveneF, leveneP) = ComputeLevene(orderedGroups);
            var heterogenous = leveneP < alpha;

            var etaSquared = ssTotal == 0 ? 0.0 : ssBetween / ssTotal;
            var omegaSquared = (ssTotal + msWithin) == 0 ? 0.0 : (ssBetween - (dfBetween * msWithin)) / (ssTotal + msWithin);
            omegaSquared = Math.Max(0.0, omegaSquared);
            var cohensF = etaSquared >= 1 ? double.PositiveInfinity : Math.Sqrt(etaSquared / Math.Max(1e-12, 1 - etaSquared));

            var rows = new List<object[]>();

            rows.Add(["POPISNÉ STATISTIKY", "", "", "", "", "", ""]);
            rows.Add(["Skupina", "n", "x̄", "med", "sₓ", "min", "max"]);
            foreach (var group in orderedGroups)
            {
                rows.Add([
                    group.Key.ToUpperInvariant(),
                    group.Value.Count,
                    StatisticsHelper.Mean(group.Value),
                    StatisticsHelper.Median(group.Value),
                    StatisticsHelper.SampleStandardDeviation(group.Value),
                    group.Value.Min(),
                    group.Value.Max()
                ]);
            }
            rows.Add(["", "", "", "", "", "", ""]);

            rows.Add(["ANOVA", "", "", "", "", "", ""]);
            rows.Add(["Zdroj variability", "SS", "df", "MS", "F", "p-hodnota", ""]);
            rows.Add(["Mezi skupinami", ssBetween, dfBetween, msBetween, f, p, ""]);
            rows.Add(["Uvnitř skupin (rezidua)", ssWithin, dfWithin, msWithin, "", "", ""]);
            rows.Add(["Celkem", ssTotal, dfTotal, "", "", "", ""]);
            rows.Add(["α", alpha, "", "", "", "", ""]);
            rows.Add(["F₁₋α", fCrit, "", "", "", "", ""]);
            rows.Add(["", "", "", "", "", "", ""]);
            rows.Add(["OVĚŘENÍ PŘEDPOKLADŮ", "", "", "", "", "", ""]);
            rows.Add(["Levenův test (homogenita rozptylů)", "F", "df₁", "df₂", "p-hodnota", "⚠ heterogenní rozptyly", ""]);
            rows.Add(["", leveneF, dfBetween, dfWithin, leveneP, heterogenous, ""]);
            rows.Add(["", "", "", "", "", "", ""]);

            rows.Add(["VELIKOST ÚČINKU", "", "", "", "", "", ""]);
            rows.Add(["η²", etaSquared, "", "", "", "", ""]);
            rows.Add(["ω²", omegaSquared, "", "", "", "", ""]);
            rows.Add(["f", cohensF, "", "", "", "", ""]);
            rows.Add(["", "", "", "", "", "", ""]);

            if (parsedPostHoc != "none")
            {
                var title = parsedPostHoc == "tukey" ? "POST-HOC: TUKEY HSD (BONFERRONI FALLBACK)" : $"POST-HOC: {parsedPostHoc.ToUpperInvariant()}";
                rows.Add([title, "", "", "", "", "", ""]);
                rows.Add(["Skupina A", "Skupina B", "Δ průměrů (A−B)", "p-hodnota", "Sig.", "", ""]);

                var comparisons = BuildPostHocComparisons(orderedGroups, msWithin, dfWithin, parsedPostHoc, alpha);
                foreach (var comparison in comparisons)
                {
                    rows.Add([comparison.GroupA.ToUpperInvariant(), comparison.GroupB.ToUpperInvariant(), comparison.Difference, comparison.PValue, comparison.Significance, "", ""]);
                }
            }

            return SpillBuilder.Build(rows);
        }

        private static (double F, double P) ComputeLevene((string Key, List<double> Value)[] groups)
        {
            var transformed = new List<(string Group, double Value)>();
            foreach (var group in groups)
            {
                var median = StatisticsHelper.Median(group.Value);
                transformed.AddRange(group.Value.Select(value => (group.Key, Math.Abs(value - median))));
            }

            var grouped = transformed
                .GroupBy(item => item.Group)
                .Select(group => (group.Key, Values: group.Select(item => item.Value).ToList()))
                .ToArray();

            var totalN = grouped.Sum(group => group.Values.Count);
            var grandMean = grouped.SelectMany(group => group.Values).Average();
            var ssBetween = grouped.Sum(group => group.Values.Count * Math.Pow(group.Values.Average() - grandMean, 2));
            var ssWithin = grouped.Sum(group =>
            {
                var mean = group.Values.Average();
                return group.Values.Sum(value => Math.Pow(value - mean, 2));
            });

            var dfBetween = grouped.Length - 1;
            var dfWithin = totalN - grouped.Length;
            var msBetween = ssBetween / dfBetween;
            var msWithin = ssWithin / dfWithin;
            var f = msWithin == 0 ? 0.0 : msBetween / msWithin;
            var p = 1 - FisherSnedecor.CDF(dfBetween, dfWithin, f);
            return (f, p);
        }

        private static IEnumerable<(string GroupA, string GroupB, double Difference, double PValue, string Significance)> BuildPostHocComparisons(
            (string Key, List<double> Value)[] groups,
            double msWithin,
            int dfWithin,
            string postHoc,
            double alpha)
        {
            var m = groups.Length * (groups.Length - 1) / 2.0;

            for (var i = 0; i < groups.Length; i++)
            {
                for (var j = i + 1; j < groups.Length; j++)
                {
                    var a = groups[i];
                    var b = groups[j];
                    var meanA = a.Value.Average();
                    var meanB = b.Value.Average();
                    var diff = meanA - meanB;
                    double pValue;

                    switch (postHoc)
                    {
                        case "games-howell":
                        {
                            var varA = Math.Pow(StatisticsHelper.SampleStandardDeviation(a.Value), 2) / a.Value.Count;
                            var varB = Math.Pow(StatisticsHelper.SampleStandardDeviation(b.Value), 2) / b.Value.Count;
                            var t = Math.Abs(diff) / Math.Sqrt(varA + varB);
                            var df = Math.Pow(varA + varB, 2) /
                                     ((Math.Pow(varA, 2) / (a.Value.Count - 1)) + (Math.Pow(varB, 2) / (b.Value.Count - 1)));
                            pValue = 2 * (1 - StudentT.CDF(0, 1, df, t));
                            break;
                        }
                        case "scheffe":
                        {
                            var f = (diff * diff) / ((groups.Length - 1) * msWithin * ((1.0 / a.Value.Count) + (1.0 / b.Value.Count)));
                            pValue = 1 - FisherSnedecor.CDF(groups.Length - 1, dfWithin, f);
                            break;
                        }
                        case "tukey":
                        case "bonferroni":
                        default:
                        {
                            var se = Math.Sqrt(msWithin * ((1.0 / a.Value.Count) + (1.0 / b.Value.Count)));
                            var t = Math.Abs(diff) / se;
                            var rawP = 2 * (1 - StudentT.CDF(0, 1, dfWithin, t));
                            pValue = Math.Min(1.0, rawP * m);
                            break;
                        }
                    }

                    yield return (a.Key, b.Key, diff, pValue, ToSig(pValue, alpha));
                }
            }
        }

        private static string ToSig(double pValue, double alpha)
            => pValue < 0.001 ? "***"
                : pValue < 0.01 ? "**"
                : pValue < alpha ? "*"
                : "ns";

        private static bool TryParsePostHoc(object? input, out string postHoc)
        {
            postHoc = string.Empty;

            switch (input)
            {
                case null:
                case ExcelMissing:
                case ExcelEmpty:
                    postHoc = "none";
                    return true;
                case double number:
                    return Math.Abs(number - Math.Round(number)) < 1e-9 && TryParsePostHocCode((int)number, out postHoc);
                case int number:
                    return TryParsePostHocCode(number, out postHoc);
                default:
                    return false;
            }
        }

        private static bool TryParsePostHocCode(int code, out string postHoc)
        {
            postHoc = code switch
            {
                0 => "none",
                1 => "tukey",
                2 => "bonferroni",
                3 => "scheffe",
                4 => "games-howell",
                _ => string.Empty
            };

            return postHoc.Length > 0;
        }

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
