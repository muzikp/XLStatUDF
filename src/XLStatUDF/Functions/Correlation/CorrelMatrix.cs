/// <summary>
/// Implements a correlation matrix for Pearson or Spearman correlation across multiple columns.
/// </summary>
namespace XLStatUDF.Functions.Correlation
{
    using System.Globalization;
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using XLStatUDF.Helpers;

    public static class CorrelMatrix
    {
        [ExcelFunction(Name = "CORREL.MATRIX", Description = "Korelační matice pro více sloupců", Category = FunctionCategories.Tests)]
        public static object[,] Run(
            [ExcelArgument(Name = "p_minimum", Description = "Volitelně: zobrazí jen vazby s p < zadaná hodnota")] object? pMinimum,
            [ExcelArgument(Name = "data", Description = "Vstupní data; alespoň dva sloupce")] object data,
            [ExcelArgument(Name = "metoda", Description = "Volitelně: 0=Pearson, 1=Spearman")] object? method = null,
            [ExcelArgument(Name = "vystup", Description = "Volitelně: 0=koeficienty, 1=p-hodnoty, 2=koeficient+p, 3=koeficient+p+sig, 4=koeficient+sig v jedné buňce")] object? output = null,
            [ExcelArgument(Name = "ma_záhlaví", Description = "Volitelně: 0=autodetect, 1=má záhlaví, 2=nemá záhlaví")] object? hasHeader = null)
        {
            if (!ArgumentHelper.TryParseHeaderMode(hasHeader, out var headerMode))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!TryParseMethod(method, out var parsedMethod) || !TryParseOutput(output, out var parsedOutput))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!TryParsePMinimum(pMinimum, out var parsedPMinimum))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!TryReadMatrix(data, headerMode, out var columnNames, out var matrix, out var error))
            {
                return SpillError(error!);
            }

            var n = matrix.Length;
            var p = columnNames.Length;
            if (p < 2 || n < 3)
            {
                return SpillError(ExcelErrors.Count);
            }

            var coefficients = new double[p, p];
            var pValues = new double[p, p];
            var significance = new string[p, p];

            for (var i = 0; i < p; i++)
            {
                var xi = matrix.Select(row => row[i]).ToArray();
                if (HasNoVariance(xi))
                {
                    return SpillError(ExcelErrors.Num);
                }

                for (var j = i; j < p; j++)
                {
                    if (i == j)
                    {
                        coefficients[i, j] = 1.0;
                        pValues[i, j] = 0.0;
                        significance[i, j] = "***";
                        continue;
                    }

                    var xj = matrix.Select(row => row[j]).ToArray();
                    if (HasNoVariance(xj))
                    {
                        return SpillError(ExcelErrors.Num);
                    }

                    var coefficient = parsedMethod == CorrelationMethod.Pearson
                        ? StatisticsHelper.PearsonCorrelation(xi, xj)
                        : StatisticsHelper.PearsonCorrelation(RankHelper.MidRank(xi), RankHelper.MidRank(xj));

                    var pValue = TwoTailPValue(coefficient, n);
                    var sig = ToSig(pValue);

                    coefficients[i, j] = coefficients[j, i] = coefficient;
                    pValues[i, j] = pValues[j, i] = pValue;
                    significance[i, j] = significance[j, i] = sig;
                }
            }

            var visible = BuildVisibilityMask(pValues, parsedPMinimum);

            return parsedOutput switch
            {
                CorrelationMatrixOutput.Coefficients => BuildSingleMatrix(columnNames, coefficients, visible),
                CorrelationMatrixOutput.PValues => BuildSingleMatrix(columnNames, pValues, visible),
                CorrelationMatrixOutput.CoefficientsAndPValues => BuildStackedMatrix(columnNames, coefficients, pValues, visible),
                CorrelationMatrixOutput.CoefficientsPValuesAndSignificance => BuildStackedMatrix(columnNames, coefficients, pValues, visible, significance),
                CorrelationMatrixOutput.CoefficientsWithSignificance => BuildCoefficientStringMatrix(columnNames, coefficients, significance, visible),
                _ => SpillError(ExcelErrors.Value)
            };
        }

        private static object[,] BuildSingleMatrix(string[] columnNames, double[,] values, bool[,] visible)
        {
            var size = columnNames.Length;
            var result = new object[size + 1, size + 1];
            result[0, 0] = string.Empty;

            for (var i = 0; i < size; i++)
            {
                result[0, i + 1] = columnNames[i];
                result[i + 1, 0] = columnNames[i];
                for (var j = 0; j < size; j++)
                {
                    result[i + 1, j + 1] = visible[i, j] ? values[i, j] : string.Empty;
                }
            }

            return result;
        }

        private static object[,] BuildStackedMatrix(string[] columnNames, double[,] coefficients, double[,] pValues, bool[,] visible, string[,]? significance = null)
        {
            var size = columnNames.Length;
            var blockHeight = significance is null ? 2 : 3;
            var result = new object[(size * blockHeight) + 1, size + 1];
            result[0, 0] = string.Empty;

            for (var i = 0; i < size; i++)
            {
                result[0, i + 1] = columnNames[i];
            }

            for (var row = 0; row < size; row++)
            {
                var baseRow = 1 + (row * blockHeight);
                result[baseRow, 0] = columnNames[row];
                result[baseRow + 1, 0] = $"p ({columnNames[row]})";
                if (significance is not null)
                {
                    result[baseRow + 2, 0] = $"sig. ({columnNames[row]})";
                }

                for (var col = 0; col < size; col++)
                {
                    result[baseRow, col + 1] = visible[row, col] ? coefficients[row, col] : string.Empty;
                    result[baseRow + 1, col + 1] = visible[row, col] ? pValues[row, col] : string.Empty;
                    if (significance is not null)
                    {
                        result[baseRow + 2, col + 1] = visible[row, col] ? significance[row, col] : string.Empty;
                    }
                }
            }

            return result;
        }

        private static object[,] BuildCoefficientStringMatrix(string[] columnNames, double[,] coefficients, string[,] significance, bool[,] visible)
        {
            var size = columnNames.Length;
            var result = new object[size + 1, size + 1];
            result[0, 0] = string.Empty;

            for (var i = 0; i < size; i++)
            {
                result[0, i + 1] = columnNames[i];
                result[i + 1, 0] = columnNames[i];
                for (var j = 0; j < size; j++)
                {
                    result[i + 1, j + 1] = visible[i, j]
                        ? coefficients[i, j].ToString("0.#####", CultureInfo.CurrentCulture) + significance[i, j]
                        : string.Empty;
                }
            }

            return result;
        }

        private static bool[,] BuildVisibilityMask(double[,] pValues, double? pMinimum)
        {
            var size = pValues.GetLength(0);
            var visible = new bool[size, size];

            for (var i = 0; i < size; i++)
            {
                for (var j = 0; j < size; j++)
                {
                    visible[i, j] = pMinimum is null
                        ? true
                        : i != j && pValues[i, j] < pMinimum.Value;
                }
            }

            return visible;
        }

        private static bool TryReadMatrix(object input, HeaderMode headerMode, out string[] columnNames, out double[][] matrix, out object? error)
        {
            columnNames = [];
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

            columnNames = Enumerable.Range(0, colCount)
                .Select(index =>
                {
                    if (startRow == 1)
                    {
                        var label = Convert.ToString(rawMatrix[0, index], CultureInfo.InvariantCulture);
                        return string.IsNullOrWhiteSpace(label) ? $"Proměnná {index + 1}" : label;
                    }

                    return $"Proměnná {index + 1}";
                })
                .ToArray();

            var rows = new List<double[]>();
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

                if (hasBlank)
                {
                    continue;
                }

                rows.Add(values);
            }

            if (rows.Count < 3)
            {
                error = ExcelErrors.Count;
                return false;
            }

            matrix = rows.ToArray();
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

        private static double TwoTailPValue(double correlation, int sampleSize)
        {
            if (Math.Abs(1 - Math.Abs(correlation)) < 1e-12)
            {
                return 0.0;
            }

            var df = sampleSize - 2;
            var t = correlation * Math.Sqrt(df) / Math.Sqrt(1 - (correlation * correlation));
            return 2 * (1 - StudentT.CDF(0, 1, df, Math.Abs(t)));
        }

        private static bool HasNoVariance(IReadOnlyList<double> values)
            => values.All(value => Math.Abs(value - values[0]) < 1e-12);

        private static bool TryParseMethod(object? input, out CorrelationMethod method)
        {
            method = CorrelationMethod.Pearson;
            switch (input)
            {
                case null:
                case ExcelMissing:
                case ExcelEmpty:
                    return true;
                case double number when Math.Abs(number - Math.Round(number)) < 1e-9 && number is 0 or 1:
                    method = (CorrelationMethod)(int)Math.Round(number);
                    return true;
                case int number when number is 0 or 1:
                    method = (CorrelationMethod)number;
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryParseOutput(object? input, out CorrelationMatrixOutput output)
        {
            output = CorrelationMatrixOutput.Coefficients;
            switch (input)
            {
                case null:
                case ExcelMissing:
                case ExcelEmpty:
                    return true;
                case double number when Math.Abs(number - Math.Round(number)) < 1e-9 && number >= 0 && number <= 4:
                    output = (CorrelationMatrixOutput)(int)Math.Round(number);
                    return true;
                case int number when number >= 0 && number <= 4:
                    output = (CorrelationMatrixOutput)number;
                    return true;
                default:
                    return false;
            }
        }

        private static bool TryParsePMinimum(object? input, out double? pMinimum)
        {
            switch (input)
            {
                case null:
                case ExcelMissing:
                case ExcelEmpty:
                    pMinimum = null;
                    return true;
                default:
                    if (!DataHelper.TryGetDouble(input, out var parsed) || parsed < 0 || parsed > 1)
                    {
                        pMinimum = null;
                        return false;
                    }

                    pMinimum = parsed;
                    return true;
            }
        }

        private static string ToSig(double pValue)
            => pValue < 0.001 ? "***"
                : pValue < 0.01 ? "**"
                : pValue < 0.05 ? "*"
                : string.Empty;

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };

        private enum CorrelationMethod
        {
            Pearson = 0,
            Spearman = 1
        }

        private enum CorrelationMatrixOutput
        {
            Coefficients = 0,
            PValues = 1,
            CoefficientsAndPValues = 2,
            CoefficientsPValuesAndSignificance = 3,
            CoefficientsWithSignificance = 4
        }
    }
}
