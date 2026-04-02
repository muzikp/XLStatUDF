/// <summary>
/// Implements ANCOVA for grouped data with one factor and one or more covariates.
/// </summary>
namespace XLStatUDF.Functions.Tests
{
    using System.Globalization;
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using MathNet.Numerics.LinearAlgebra;
    using MathNet.Numerics.LinearAlgebra.Double;
    using XLStatUDF.Helpers;

    public static class Ancova
    {
        [ExcelFunction(Name = "ANCOVA.G", Description = "[Testy] ANCOVA nad groupovanými daty", Category = "XLStatUDF")]
        public static object[,] Run(
            [ExcelArgument(Name = "faktor", Description = "Kategorie faktoru")] object groups,
            [ExcelArgument(Name = "zavisla_promenna", Description = "Závislá proměnná")] object values,
            [ExcelArgument(Name = "kovariaty", Description = "Jedna nebo více kovariát ve sloupcích")] object covariates,
            [ExcelArgument(Name = "post_hoc", Description = "Volitelne: 0=none, 1=tukey, 2=bonferroni, 3=scheffe, 4=games-howell")] object? postHoc = null,
            [ExcelArgument(Name = "alpha", Description = "Hladina významnosti")] double alpha = 0.05,
            [ExcelArgument(Name = "ma_zahlavi", Description = "Volitelne: 0=autodetect, 1=ma zahlavi, 2=nema zahlavi")] object? hasHeader = null)
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

            if (!TryReadAncovaInputs(groups, values, covariates, parsedHeaderMode, out var data, out var error))
            {
                return SpillError(error!);
            }

            if (data.GroupLabels.Length < 2)
            {
                return SpillError(ExcelErrors.Count);
            }

            var includedCovariates = Enumerable.Range(0, data.CovariateNames.Length).ToArray();
            var fullModel = FitModel(data, includeGroups: true, includeCovariates: includedCovariates);
            if (!fullModel.Success || fullModel.ResidualDegreesOfFreedom <= 0)
            {
                return SpillError(ExcelErrors.Num);
            }

            var covariateOnlyModel = FitModel(data, includeGroups: false, includeCovariates: includedCovariates);
            if (!covariateOnlyModel.Success)
            {
                return SpillError(ExcelErrors.Num);
            }

            var factorTest = BuildNestedFTest(covariateOnlyModel, fullModel);
            var covariateTests = includedCovariates
                .Select(index =>
                {
                    var reducedCovariates = includedCovariates.Where(i => i != index).ToArray();
                    var reducedModel = FitModel(data, includeGroups: true, includeCovariates: reducedCovariates);
                    return (Name: data.CovariateNames[index], Test: BuildNestedFTest(reducedModel, fullModel));
                })
                .ToArray();

            var interactionTests = includedCovariates
                .Select(index => BuildInteractionTest(data, index))
                .ToArray();

            var totalMean = data.Y.Average();
            var ssTotal = data.Y.Sum(value => Math.Pow(value - totalMean, 2));
            var residualDf = fullModel.ResidualDegreesOfFreedom;
            var msError = fullModel.MeanSquareError;
            var criticalF = FisherSnedecor.InvCDF(Math.Max(1, factorTest.DfEffect), residualDf, 1 - alpha);
            var mainModelSumSquares = factorTest.SumOfSquares
                + covariateTests.Sum(item => item.Test.SumOfSquares)
                + fullModel.SumSquaredErrors;

            var rows = new List<object[]>();

            var descriptiveColumnCount = Math.Max(3 + data.CovariateNames.Length, 4);
            SpillBuilder.AddHeader(rows, "POPISNE STATISTIKY", descriptiveColumnCount);
            var descriptiveHeader = new object[descriptiveColumnCount];
            descriptiveHeader[0] = "Skupina";
            descriptiveHeader[1] = "n";
            descriptiveHeader[2] = "y";
            for (var i = 0; i < data.CovariateNames.Length; i++)
            {
                descriptiveHeader[3 + i] = data.CovariateNames[i];
            }

            rows.Add(descriptiveHeader);
            foreach (var summary in BuildGroupSummaries(data))
            {
                var row = new object[descriptiveColumnCount];
                row[0] = summary.Group;
                row[1] = summary.Count;
                row[2] = summary.MeanY;
                for (var i = 0; i < summary.CovariateMeans.Length; i++)
                {
                    row[3 + i] = summary.CovariateMeans[i];
                }

                rows.Add(row);
            }

            SpillBuilder.AddSeparator(rows, descriptiveColumnCount);

            SpillBuilder.AddHeader(rows, "ANCOVA", 10);
            rows.Add(["Zdroj", "SS", "df", "MS", "F", "p", "η²", "η²p", "ω²", "ω²p"]);
            rows.Add(BuildEffectRow("Faktor", factorTest, mainModelSumSquares, fullModel.SumSquaredErrors, msError, data.Y.Length));
            foreach (var covariate in covariateTests)
            {
                rows.Add(BuildEffectRow(covariate.Name, covariate.Test, mainModelSumSquares, fullModel.SumSquaredErrors, msError, data.Y.Length));
            }

            foreach (var interaction in interactionTests)
            {
                var interactionModelSumSquares = interaction.Test.SumOfSquares + fullModel.SumSquaredErrors;
                rows.Add(BuildEffectRow($"Skupina × {interaction.Name}", interaction.Test, interactionModelSumSquares, fullModel.SumSquaredErrors, msError, data.Y.Length));
            }

            rows.Add(["Reziduum", fullModel.SumSquaredErrors, residualDf, msError, "", "", "", "", "", ""]);
            rows.Add(["Celkem", ssTotal, data.Y.Length - 1, "", "", "", "", "", "", ""]);
            rows.Add(["α", alpha, "", "", "", "", "", "", "", ""]);
            rows.Add(["F₁₋α", criticalF, "", "", "", "", "", "", "", ""]);
            SpillBuilder.AddSeparator(rows, 10);

            var significantInteractions = interactionTests
                .Where(interaction => !double.IsNaN(interaction.Test.PValue) && interaction.Test.PValue < alpha)
                .ToArray();
            if (significantInteractions.Length > 0)
            {
                SpillBuilder.AddHeader(rows, "UPOZORNENI", 2);
                SpillBuilder.AddRow(rows, "Homogenita sklonu", "Porusena");
                SpillBuilder.AddRow(rows, "Detail", $"Vyznacna interakce: {string.Join("; ", significantInteractions.Select(item => $"Skupina × {item.Name}"))}");
                SpillBuilder.AddSeparator(rows, 2);
            }

            SpillBuilder.AddHeader(rows, "ADJUSTED MEANS", 6);
            rows.Add(["Skupina", "Adjusted mean", "SE", "CI dolni", "CI horni", ""]);
            var adjustedMeans = BuildAdjustedMeans(data, fullModel, alpha);
            foreach (var adjusted in adjustedMeans)
            {
                rows.Add([adjusted.Group, adjusted.Mean, adjusted.StandardError, adjusted.LowerBound, adjusted.UpperBound, ""]);
            }

            SpillBuilder.AddSeparator(rows, 6);

            if (parsedPostHoc != "none")
            {
                var title = parsedPostHoc switch
                {
                    "tukey" => "POST-HOC: TUKEY (BONFERRONI FALLBACK)",
                    "scheffe" => "POST-HOC: SCHEFFE",
                    "games-howell" => "POST-HOC: GAMES-HOWELL (BONFERRONI FALLBACK)",
                    _ => $"POST-HOC: {parsedPostHoc.ToUpperInvariant()}"
                };
                SpillBuilder.AddHeader(rows, title, 5);
                rows.Add(["Skupina A", "Skupina B", "Δ adjusted means", "p", "Sig."]);
                foreach (var comparison in BuildPostHocComparisons(adjustedMeans, fullModel, parsedPostHoc, alpha))
                {
                    rows.Add([comparison.GroupA, comparison.GroupB, comparison.Difference, comparison.PValue, comparison.Significance]);
                }
            }

            return SpillBuilder.Build(rows);
        }

        private static bool TryReadAncovaInputs(
            object groupsInput,
            object yInput,
            object covariatesInput,
            HeaderMode headerMode,
            out AncovaData data,
            out object? error)
        {
            data = null!;
            var rawGroups = DataHelper.Flatten(groupsInput);
            var rawY = DataHelper.Flatten(yInput);
            if (covariatesInput is not object[,] covariateMatrix)
            {
                error = ExcelErrors.Value;
                return false;
            }

            var rowCount = covariateMatrix.GetLength(0);
            var covariateCount = covariateMatrix.GetLength(1);
            if (covariateCount < 1)
            {
                error = ExcelErrors.Value;
                return false;
            }

            if (rawGroups.Length != rawY.Length || rawGroups.Length != rowCount)
            {
                error = ExcelErrors.Length;
                return false;
            }

            var startRow = 0;
            if (headerMode == HeaderMode.HasHeader)
            {
                startRow = 1;
            }
            else if (headerMode == HeaderMode.AutoDetect && ShouldSkipAncovaHeader(rawGroups, rawY, covariateMatrix))
            {
                startRow = 1;
            }

            var covariateNames = Enumerable.Range(0, covariateCount)
                .Select(index =>
                {
                    if (startRow == 1)
                    {
                        var label = Convert.ToString(covariateMatrix[0, index], CultureInfo.InvariantCulture);
                        return string.IsNullOrWhiteSpace(label) ? $"Kovariata {index + 1}" : label;
                    }

                    return $"Kovariata {index + 1}";
                })
                .ToArray();

            var groups = new List<string>();
            var yValues = new List<double>();
            var covariates = new List<double[]>();

            for (var row = startRow; row < rowCount; row++)
            {
                var groupValue = rawGroups[row];
                var yValue = rawY[row];
                if (DataHelper.IsBlank(groupValue) || DataHelper.IsBlank(yValue))
                {
                    continue;
                }

                if (!DataHelper.TryGetDouble(yValue, out var parsedY))
                {
                    error = ExcelErrors.Value;
                    return false;
                }

                var covariateRow = new double[covariateCount];
                var skipRow = false;
                for (var col = 0; col < covariateCount; col++)
                {
                    var cell = covariateMatrix[row, col];
                    if (DataHelper.IsBlank(cell))
                    {
                        skipRow = true;
                        break;
                    }

                    if (!DataHelper.TryGetDouble(cell, out var parsedCovariate))
                    {
                        error = ExcelErrors.Value;
                        return false;
                    }

                    covariateRow[col] = parsedCovariate;
                }

                if (skipRow)
                {
                    continue;
                }

                var label = Convert.ToString(groupValue, CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(label))
                {
                    continue;
                }

                groups.Add(label);
                yValues.Add(parsedY);
                covariates.Add(covariateRow);
            }

            if (groups.Count < 3)
            {
                error = ExcelErrors.Count;
                return false;
            }

            var distinctGroups = groups.Distinct(StringComparer.Ordinal).OrderBy(value => value, StringComparer.Ordinal).ToArray();
            if (distinctGroups.Length < 2)
            {
                error = ExcelErrors.Count;
                return false;
            }

            if (distinctGroups.Any(group => groups.Count(value => string.Equals(value, group, StringComparison.Ordinal)) < 2))
            {
                error = ExcelErrors.Count;
                return false;
            }

            data = new AncovaData(groups.ToArray(), yValues.ToArray(), covariates.ToArray(), covariateNames, distinctGroups);
            error = null;
            return true;
        }

        private static bool ShouldSkipAncovaHeader(object[] groups, object[] yValues, object[,] covariates)
        {
            if (groups.Length < 2 || covariates.GetLength(0) < 2)
            {
                return false;
            }

            if (!DataHelper.ShouldSkipLeadingHeader(groups, item => !DataHelper.IsBlank(item)))
            {
                return false;
            }

            if (!DataHelper.ShouldSkipLeadingHeader(yValues, item => DataHelper.TryGetDouble(item, out _)))
            {
                return false;
            }

            for (var col = 0; col < covariates.GetLength(1); col++)
            {
                var items = Enumerable.Range(0, covariates.GetLength(0)).Select(row => covariates[row, col]).ToArray();
                if (!DataHelper.ShouldSkipLeadingHeader(items, item => DataHelper.TryGetDouble(item, out _)))
                {
                    return false;
                }
            }

            return true;
        }

        private static AncovaModel FitModel(AncovaData data, bool includeGroups, int[] includeCovariates, int? interactionCovariate = null)
        {
            try
            {
                var designRows = new List<double[]>(data.Y.Length);
                var groupDummies = includeGroups ? data.GroupLabels.Skip(1).ToArray() : [];

                for (var i = 0; i < data.Y.Length; i++)
                {
                    var row = new List<double> { 1.0 };

                    foreach (var group in groupDummies)
                    {
                        row.Add(string.Equals(data.Groups[i], group, StringComparison.Ordinal) ? 1.0 : 0.0);
                    }

                    foreach (var covariateIndex in includeCovariates)
                    {
                        row.Add(data.Covariates[i][covariateIndex]);
                    }

                    if (interactionCovariate.HasValue)
                    {
                        foreach (var group in groupDummies)
                        {
                            var isGroup = string.Equals(data.Groups[i], group, StringComparison.Ordinal) ? 1.0 : 0.0;
                            row.Add(isGroup * data.Covariates[i][interactionCovariate.Value]);
                        }
                    }

                    designRows.Add(row.ToArray());
                }

                var x = DenseMatrix.OfRowArrays(designRows);
                var y = DenseVector.OfArray(data.Y);
                var xtx = x.TransposeThisAndMultiply(x);
                var xtxInverse = xtx.Inverse();
                var beta = xtxInverse * x.TransposeThisAndMultiply(y);
                var fitted = x * beta;
                var residuals = y - fitted;
                var sse = residuals.DotProduct(residuals);
                var dfResidual = data.Y.Length - x.ColumnCount;
                var mse = dfResidual > 0 ? sse / dfResidual : double.NaN;
                var covariance = xtxInverse * mse;

                return new AncovaModel(true, beta, covariance, sse, dfResidual, mse);
            }
            catch
            {
                return AncovaModel.Failed;
            }
        }

        private static NestedFTest BuildNestedFTest(AncovaModel reducedModel, AncovaModel fullModel)
        {
            if (!reducedModel.Success || !fullModel.Success || fullModel.ResidualDegreesOfFreedom <= 0)
            {
                return NestedFTest.Invalid;
            }

            var ssEffect = Math.Max(0.0, reducedModel.SumSquaredErrors - fullModel.SumSquaredErrors);
            var dfEffect = reducedModel.ResidualDegreesOfFreedom - fullModel.ResidualDegreesOfFreedom;
            if (dfEffect <= 0 || fullModel.MeanSquareError <= 0)
            {
                return NestedFTest.Invalid;
            }

            var msEffect = ssEffect / dfEffect;
            var f = msEffect / fullModel.MeanSquareError;
            var p = 1 - FisherSnedecor.CDF(dfEffect, fullModel.ResidualDegreesOfFreedom, f);
            return new NestedFTest(ssEffect, dfEffect, msEffect, f, p);
        }

        private static (string Name, NestedFTest Test) BuildInteractionTest(AncovaData data, int covariateIndex)
        {
            var includedCovariates = Enumerable.Range(0, data.CovariateNames.Length).ToArray();
            var reducedModel = FitModel(data, includeGroups: true, includeCovariates: includedCovariates);
            var fullModel = FitModel(data, includeGroups: true, includeCovariates: includedCovariates, interactionCovariate: covariateIndex);
            return (data.CovariateNames[covariateIndex], BuildNestedFTest(reducedModel, fullModel));
        }

        private static IEnumerable<GroupSummary> BuildGroupSummaries(AncovaData data)
        {
            foreach (var group in data.GroupLabels)
            {
                var indices = Enumerable.Range(0, data.Groups.Length)
                    .Where(index => string.Equals(data.Groups[index], group, StringComparison.Ordinal))
                    .ToArray();

                yield return new GroupSummary(
                    group,
                    indices.Length,
                    indices.Average(index => data.Y[index]),
                    Enumerable.Range(0, data.CovariateNames.Length)
                        .Select(covariate => indices.Average(index => data.Covariates[index][covariate]))
                        .ToArray());
            }
        }

        private static AdjustedMeanRow[] BuildAdjustedMeans(AncovaData data, AncovaModel fullModel, double alpha)
        {
            var covariateMeans = Enumerable.Range(0, data.CovariateNames.Length)
                .Select(index => data.Covariates.Average(row => row[index]))
                .ToArray();
            var tCritical = StudentT.InvCDF(0, 1, fullModel.ResidualDegreesOfFreedom, 1 - alpha / 2.0);

            return data.GroupLabels.Select(group =>
            {
                var design = BuildAdjustedMeanVector(data.GroupLabels, data.CovariateNames.Length, covariateMeans, group);
                var mean = design.DotProduct(fullModel.Coefficients);
                var variance = design * (fullModel.CoefficientCovariance * design);
                var se = Math.Sqrt(Math.Max(0.0, variance));
                return new AdjustedMeanRow(group, mean, se, mean - tCritical * se, mean + tCritical * se, design);
            }).ToArray();
        }

        private static IEnumerable<PostHocComparison> BuildPostHocComparisons(AdjustedMeanRow[] adjustedMeans, AncovaModel fullModel, string postHoc, double alpha)
        {
            var m = adjustedMeans.Length * (adjustedMeans.Length - 1) / 2.0;
            for (var i = 0; i < adjustedMeans.Length; i++)
            {
                for (var j = i + 1; j < adjustedMeans.Length; j++)
                {
                    var a = adjustedMeans[i];
                    var b = adjustedMeans[j];
                    var diffVector = a.DesignVector - b.DesignVector;
                    var diff = a.Mean - b.Mean;
                    var variance = diffVector * (fullModel.CoefficientCovariance * diffVector);
                    var se = Math.Sqrt(Math.Max(0.0, variance));
                    var t = se == 0 ? 0.0 : Math.Abs(diff) / se;
                    var rawP = 2 * (1 - StudentT.CDF(0, 1, fullModel.ResidualDegreesOfFreedom, t));
                    var pValue = postHoc switch
                    {
                        "scheffe" => Math.Min(1.0, 1 - FisherSnedecor.CDF(adjustedMeans.Length - 1, fullModel.ResidualDegreesOfFreedom, (t * t) / Math.Max(1, adjustedMeans.Length - 1))),
                        _ => Math.Min(1.0, rawP * m)
                    };

                    yield return new PostHocComparison(a.Group, b.Group, diff, pValue, ToSig(pValue, alpha));
                }
            }
        }

        private static Vector<double> BuildAdjustedMeanVector(string[] groupLabels, int covariateCount, double[] covariateMeans, string activeGroup)
        {
            var values = new List<double> { 1.0 };
            foreach (var group in groupLabels.Skip(1))
            {
                values.Add(string.Equals(group, activeGroup, StringComparison.Ordinal) ? 1.0 : 0.0);
            }

            values.AddRange(covariateMeans);
            return DenseVector.OfEnumerable(values);
        }

        private static object[] BuildEffectRow(string name, NestedFTest test, double ssTotal, double ssError, double msError, int sampleSize)
        {
            var etaSquared = ssTotal <= 0 ? double.NaN : test.SumOfSquares / ssTotal;
            var etaSquaredPartial = (test.SumOfSquares + ssError) <= 0 ? double.NaN : test.SumOfSquares / (test.SumOfSquares + ssError);
            var omegaSquared = (ssTotal + msError) <= 0 ? double.NaN : Math.Max(0.0, (test.SumOfSquares - (test.DfEffect * msError)) / (ssTotal + msError));
            var omegaSquaredPartial = sampleSize <= 0 || double.IsNaN(test.FValue)
                ? double.NaN
                : Math.Max(0.0, (test.DfEffect * (test.FValue - 1.0)) / ((test.DfEffect * (test.FValue - 1.0)) + sampleSize));

            return [name, test.SumOfSquares, test.DfEffect, test.MeanSquare, test.FValue, test.PValue, etaSquared, etaSquaredPartial, omegaSquared, omegaSquaredPartial];
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

        private sealed record AncovaData(string[] Groups, double[] Y, double[][] Covariates, string[] CovariateNames, string[] GroupLabels);

        private sealed record AncovaModel(bool Success, Vector<double> Coefficients, Matrix<double> CoefficientCovariance, double SumSquaredErrors, int ResidualDegreesOfFreedom, double MeanSquareError)
        {
            public static AncovaModel Failed { get; } = new(false, DenseVector.OfArray([0.0]), DenseMatrix.OfArray(new double[,] { { 0.0 } }), double.NaN, 0, double.NaN);
        }

        private sealed record NestedFTest(double SumOfSquares, int DfEffect, double MeanSquare, double FValue, double PValue)
        {
            public static NestedFTest Invalid { get; } = new(double.NaN, 0, double.NaN, double.NaN, double.NaN);
        }

        private sealed record GroupSummary(string Group, int Count, double MeanY, double[] CovariateMeans);

        private sealed record AdjustedMeanRow(string Group, double Mean, double StandardError, double LowerBound, double UpperBound, Vector<double> DesignVector);

        private sealed record PostHocComparison(string GroupA, string GroupB, double Difference, double PValue, string Significance);
    }
}
