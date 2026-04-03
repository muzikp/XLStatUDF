/// <summary>
/// Implements one-factor repeated-measures ANOVA across columns.
/// </summary>
namespace XLStatUDF.Functions.Tests
{
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using XLStatUDF.Helpers;

    public static class AnovaRepeatedMeasures
    {
        [ExcelFunction(Name = "ANOVA.RM", Description = "Jednofaktorová ANOVA s opakovaným měřením nad sloupci", Category = FunctionCategories.Tests)]
        public static object[,] Run(
            [ExcelArgument(Name = "hodnoty", Description = "Matice hodnot: řádky=subjekty, sloupce=podmínky")] object values,
            [ExcelArgument(Name = "ma_záhlaví", Description = "Volitelně: 0=autodetect, 1=má záhlaví, 2=nemá záhlaví")] object? hasHeader = null,
            [ExcelArgument(Name = "alpha", Description = "Hladina významnosti")] double alpha = 0.05,
            [ExcelArgument(Name = "post_hoc", Description = "Volitelně: 0=none, 1=tukey, 2=bonferroni, 3=scheffe, 4=games-howell")] object? postHoc = null)
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

            if (!TryReadMatrix(values, parsedHeaderMode, out var conditionNames, out var matrix, out var error))
            {
                return SpillError(error!);
            }

            var subjectCount = matrix.Length;
            var conditionCount = conditionNames.Length;
            if (subjectCount < 2 || conditionCount < 2)
            {
                return SpillError(ExcelErrors.Count);
            }

            var allValues = matrix.SelectMany(row => row).ToArray();
            var grandMean = allValues.Average();
            var conditionMeans = Enumerable.Range(0, conditionCount)
                .Select(col => matrix.Average(row => row[col]))
                .ToArray();
            var subjectMeans = matrix.Select(row => row.Average()).ToArray();

            var ssTotal = allValues.Sum(value => Math.Pow(value - grandMean, 2));
            var ssConditions = subjectCount * conditionMeans.Sum(mean => Math.Pow(mean - grandMean, 2));
            var ssSubjects = conditionCount * subjectMeans.Sum(mean => Math.Pow(mean - grandMean, 2));
            var ssError = Math.Max(0.0, ssTotal - ssConditions - ssSubjects);

            var dfConditions = conditionCount - 1;
            var dfSubjects = subjectCount - 1;
            var dfError = dfConditions * dfSubjects;
            var dfTotal = (subjectCount * conditionCount) - 1;

            var msConditions = ssConditions / dfConditions;
            var msSubjects = ssSubjects / dfSubjects;
            var msError = ssError / dfError;
            var f = msError == 0 ? 0.0 : msConditions / msError;
            var p = 1 - FisherSnedecor.CDF(dfConditions, dfError, f);
            var fCrit = FisherSnedecor.InvCDF(dfConditions, dfError, 1 - alpha);

            var etaSquared = ssTotal <= 0 ? double.NaN : ssConditions / ssTotal;
            var etaSquaredPartial = (ssConditions + ssError) <= 0 ? double.NaN : ssConditions / (ssConditions + ssError);
            var omegaSquared = (ssTotal + msError) <= 0 ? double.NaN : Math.Max(0.0, (ssConditions - (dfConditions * msError)) / (ssTotal + msError));
            var omegaSquaredPartial = Math.Max(0.0, (dfConditions * (f - 1.0)) / Math.Max(1e-12, (dfConditions * (f - 1.0)) + subjectCount));

            var rows = new List<object[]>();

            SpillBuilder.AddHeader(rows, "POPISNE STATISTIKY", 7);
            rows.Add(["Podminka", "n", "x̄", "med", "sₓ", "min", "max"]);
            for (var col = 0; col < conditionCount; col++)
            {
                var series = matrix.Select(row => row[col]).ToArray();
                rows.Add([
                    conditionNames[col],
                    subjectCount,
                    StatisticsHelper.Mean(series),
                    StatisticsHelper.Median(series),
                    StatisticsHelper.SampleStandardDeviation(series),
                    series.Min(),
                    series.Max()
                ]);
            }

            SpillBuilder.AddSeparator(rows, 7);
            SpillBuilder.AddHeader(rows, "ANOVA S OPAKOVANYM MERENIM", 10);
            rows.Add(["Zdroj variability", "SS", "df", "MS", "F", "p", "η²", "η²p", "ω²", "ω²p"]);
            rows.Add(["Podminky", ssConditions, dfConditions, msConditions, f, p, etaSquared, etaSquaredPartial, omegaSquared, omegaSquaredPartial]);
            rows.Add(["Subjekty", ssSubjects, dfSubjects, msSubjects, "", "", "", "", "", ""]);
            rows.Add(["Reziduum", ssError, dfError, msError, "", "", "", "", "", ""]);
            rows.Add(["Celkem", ssTotal, dfTotal, "", "", "", "", "", "", ""]);
            rows.Add(["α", alpha, "", "", "", "", "", "", "", ""]);
            rows.Add(["F₁₋α", fCrit, "", "", "", "", "", "", "", ""]);

            SpillBuilder.AddSeparator(rows, 10);
            SpillBuilder.AddHeader(rows, "POZNAMKA", 2);
            SpillBuilder.AddRow(rows, "Sphericita", "Netestovana");
            SpillBuilder.AddSeparator(rows, 2);

            if (parsedPostHoc != "none")
            {
                var title = parsedPostHoc switch
                {
                    "tukey" => "POST-HOC: TUKEY (BONFERRONI FALLBACK)",
                    "scheffe" => "POST-HOC: SCHEFFE (BONFERRONI FALLBACK)",
                    "games-howell" => "POST-HOC: GAMES-HOWELL (BONFERRONI FALLBACK)",
                    _ => $"POST-HOC: {parsedPostHoc.ToUpperInvariant()}"
                };
                SpillBuilder.AddHeader(rows, title, 5);
                rows.Add(["Podminka A", "Podminka B", "Δ průměru (A-B)", "p", "Sig."]);

                foreach (var comparison in BuildPostHocComparisons(conditionNames, matrix, parsedPostHoc, alpha))
                {
                    rows.Add([comparison.ConditionA, comparison.ConditionB, comparison.Difference, comparison.PValue, comparison.Significance]);
                }
            }

            return SpillBuilder.Build(rows);
        }

        private static IEnumerable<(string ConditionA, string ConditionB, double Difference, double PValue, string Significance)> BuildPostHocComparisons(
            string[] conditionNames,
            double[][] matrix,
            string postHoc,
            double alpha)
        {
            var conditionCount = conditionNames.Length;
            var m = conditionCount * (conditionCount - 1) / 2.0;
            var df = matrix.Length - 1;

            for (var i = 0; i < conditionCount; i++)
            {
                for (var j = i + 1; j < conditionCount; j++)
                {
                    var diffs = matrix.Select(row => row[i] - row[j]).ToArray();
                    var meanDiff = diffs.Average();
                    var sdDiff = StatisticsHelper.SampleStandardDeviation(diffs);
                    var seDiff = sdDiff / Math.Sqrt(diffs.Length);
                    var t = seDiff == 0 ? 0.0 : Math.Abs(meanDiff) / seDiff;
                    var rawP = 2 * (1 - StudentT.CDF(0, 1, df, t));
                    var pValue = postHoc switch
                    {
                        "bonferroni" => Math.Min(1.0, rawP * m),
                        _ => Math.Min(1.0, rawP * m)
                    };

                    yield return (conditionNames[i], conditionNames[j], meanDiff, pValue, ToSig(pValue, alpha));
                }
            }
        }

        private static bool TryReadMatrix(
            object input,
            HeaderMode headerMode,
            out string[] conditionNames,
            out double[][] matrix,
            out object? error)
        {
            conditionNames = [];
            matrix = [];

            if (input is not object[,] rawMatrix)
            {
                error = ExcelErrors.Value;
                return false;
            }

            var rowCount = rawMatrix.GetLength(0);
            var colCount = rawMatrix.GetLength(1);
            if (rowCount < 1 || colCount < 2)
            {
                error = ExcelErrors.Count;
                return false;
            }

            var startRow = 0;
            if (headerMode == HeaderMode.HasHeader)
            {
                startRow = 1;
            }
            else if (headerMode == HeaderMode.AutoDetect && ShouldSkipHeader(rawMatrix))
            {
                startRow = 1;
            }

            conditionNames = Enumerable.Range(0, colCount)
                .Select(index =>
                {
                    if (startRow == 1)
                    {
                        var label = Convert.ToString(rawMatrix[0, index], System.Globalization.CultureInfo.InvariantCulture);
                        return string.IsNullOrWhiteSpace(label) ? $"Podminka {index + 1}" : label;
                    }

                    return $"Podminka {index + 1}";
                })
                .ToArray();

            var parsedRows = new List<double[]>();
            for (var row = startRow; row < rowCount; row++)
            {
                var values = new double[colCount];
                var hasAnyValue = false;
                var hasBlank = false;

                for (var col = 0; col < colCount; col++)
                {
                    var cell = rawMatrix[row, col];
                    if (DataHelper.IsBlank(cell))
                    {
                        hasBlank = true;
                        continue;
                    }

                    hasAnyValue = true;
                    if (!DataHelper.TryGetDouble(cell, out var parsed))
                    {
                        error = ExcelErrors.Value;
                        return false;
                    }

                    values[col] = parsed;
                }

                if (!hasAnyValue)
                {
                    continue;
                }

                if (hasBlank || values.Any(double.IsNaN))
                {
                    continue;
                }

                parsedRows.Add(values);
            }

            if (parsedRows.Count < 2)
            {
                error = ExcelErrors.Count;
                return false;
            }

            matrix = parsedRows.ToArray();
            error = null;
            return true;
        }

        private static bool ShouldSkipHeader(object[,] matrix)
        {
            if (matrix.GetLength(0) < 2)
            {
                return false;
            }

            for (var col = 0; col < matrix.GetLength(1); col++)
            {
                if (DataHelper.IsBlank(matrix[0, col]) || DataHelper.TryGetDouble(matrix[0, col], out _))
                {
                    return false;
                }

                var hasNumericBelow = false;
                for (var row = 1; row < matrix.GetLength(0); row++)
                {
                    if (DataHelper.IsBlank(matrix[row, col]))
                    {
                        continue;
                    }

                    if (!DataHelper.TryGetDouble(matrix[row, col], out _))
                    {
                        return false;
                    }

                    hasNumericBelow = true;
                }

                if (!hasNumericBelow)
                {
                    return false;
                }
            }

            return true;
        }

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

        private static string ToSig(double pValue, double alpha)
            => pValue < 0.001 ? "***"
                : pValue < 0.01 ? "**"
                : pValue < alpha ? "*"
                : "ns";

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
