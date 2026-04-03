/// <summary>
/// Implements pivot-style descriptive functions.
/// </summary>
namespace XLStatUDF.Functions.Descriptive
{
    using System;
    using System.Collections.Generic;
    using System.Globalization;
    using System.Linq;
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using XLStatUDF.Helpers;

    public static class Pivot
    {
        [ExcelFunction(Name = "PIVOT.COUNT", Description = "Sestaví pivot a spočítá počet neprázdných hodnot.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotCount(
            [ExcelArgument(Name = "řádky", Description = "Jeden nebo více sloupců s kategoriemi řádků včetně záhlaví.")] object rows,
            [ExcelArgument(Name = "sloupce", Description = "Volitelně: jeden nebo více sloupců s kategoriemi sloupců včetně záhlaví.")] object columns,
            [ExcelArgument(Name = "hodnoty", Description = "Jeden sloupec hodnot včetně záhlaví.")] object values)
            => RunSingle(rows, columns, values, PivotCalculation.Count);

        [ExcelFunction(Name = "PIVOT.SUM", Description = "Sestaví pivot a spočítá součet.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotSum(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.Sum);

        [ExcelFunction(Name = "PIVOT.AVERAGE", Description = "Sestaví pivot a spočítá aritmetický průměr.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotAverage(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.Average);

        [ExcelFunction(Name = "PIVOT.MIN", Description = "Sestaví pivot a spočítá minimum.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotMin(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.Min);

        [ExcelFunction(Name = "PIVOT.MAX", Description = "Sestaví pivot a spočítá maximum.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotMax(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.Max);

        [ExcelFunction(Name = "PIVOT.MEDIAN", Description = "Sestaví pivot a spočítá medián.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotMedian(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.Median);

        [ExcelFunction(Name = "PIVOT.PERCENTILE", Description = "Sestaví pivot a spočítá percentil.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotPercentile(object rows, object columns, object values,
            [ExcelArgument(Name = "kvantil", Description = "Hledaný kvantil v intervalu (0;1).")] double quantile)
            => RunSingle(rows, columns, values, PivotCalculation.Percentile, quantile);

        [ExcelFunction(Name = "PIVOT.STDEV.S", Description = "Sestaví pivot a spočítá výběrovou směrodatnou odchylku.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotStdevSample(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.StdevSample);

        [ExcelFunction(Name = "PIVOT.STDEV.P", Description = "Sestaví pivot a spočítá populační směrodatnou odchylku.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotStdevPopulation(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.StdevPopulation);

        [ExcelFunction(Name = "PIVOT.VAR.S", Description = "Sestaví pivot a spočítá výběrový rozptyl.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotVarianceSample(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.VarianceSample);

        [ExcelFunction(Name = "PIVOT.VAR.P", Description = "Sestaví pivot a spočítá populační rozptyl.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotVariancePopulation(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.VariancePopulation);

        [ExcelFunction(Name = "PIVOT.VARCOEF.S", Description = "Sestaví pivot a spočítá výběrový variační koeficient.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotVarCoefSample(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.VarCoefSample);

        [ExcelFunction(Name = "PIVOT.VARCOEF.P", Description = "Sestaví pivot a spočítá populační variační koeficient.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotVarCoefPopulation(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.VarCoefPopulation);

        [ExcelFunction(Name = "PIVOT.CONF.T", Description = "Sestaví pivot a spočítá poloviční šířku intervalu spolehlivosti pro t-rozdělení.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotConfidenceT(object rows, object columns, object values,
            [ExcelArgument(Name = "alfa", Description = "Hladina alfa v intervalu (0;1).")] double alpha,
            [ExcelArgument(Name = "smer", Description = "Volitelně: 0 = oboustranný, -1 = levostranný, 1 = pravostranný.")] double direction = 0.0)
            => RunSingle(rows, columns, values, PivotCalculation.ConfidenceT, alpha, direction);

        [ExcelFunction(Name = "PIVOT.CONF.NORM", Description = "Sestaví pivot a spočítá poloviční šířku intervalu spolehlivosti pro normální aproximaci.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotConfidenceNorm(object rows, object columns, object values,
            [ExcelArgument(Name = "alfa", Description = "Hladina alfa v intervalu (0;1).")] double alpha,
            [ExcelArgument(Name = "smer", Description = "Volitelně: 0 = oboustranný, -1 = levostranný, 1 = pravostranný.")] double direction = 0.0)
            => RunSingle(rows, columns, values, PivotCalculation.ConfidenceNorm, alpha, direction);

        [ExcelFunction(Name = "PIVOT.MAD", Description = "Sestaví pivot a spočítá medián absolutních odchylek od mediánu.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotMad(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.Mad);

        [ExcelFunction(Name = "PIVOT.IQR", Description = "Sestaví pivot a spočítá mezikvartilové rozpětí.", Category = FunctionCategories.Descriptive)]
        public static object[,] PivotIqr(object rows, object columns, object values)
            => RunSingle(rows, columns, values, PivotCalculation.Iqr);

        private static object[,] RunSingle(object rows, object columns, object values, PivotCalculation calculation, double? detail = null, double direction = 0.0)
        {
            if (!ValidateDetail(calculation, detail) || !ValidateDirection(calculation, direction))
            {
                return SpillError(ExcelErrors.Value);
            }

            var valueMatrix = NormalizeValueColumn(values);
            if (valueMatrix is null)
            {
                return SpillError(ExcelErrors.Value);
            }

            var spec = new MetricSpec(valueMatrix, calculation, detail, direction);
            if (!TryReadPivotData(rows, columns, spec, out var data, out var error))
            {
                return SpillError(error!);
            }

            if (data.RowKeys.Count == 0)
            {
                return SpillError(ExcelErrors.Count);
            }

            return BuildPivotResult(data, spec);
        }

        private static object[,] BuildPivotResult(PivotData data, MetricSpec spec)
        {
            var headerRows = data.ColumnFieldNames.Length;
            var rowFieldCount = data.RowFieldNames.Length;
            var valueStartColumn = rowFieldCount;
            var visibleRows = data.RowKeys
                .Where(rowKey => data.ColumnKeys.Any(columnKey =>
                {
                    var aggregateKey = BuildAggregateKey(rowKey.Key, columnKey.Key);
                    return data.Aggregates.TryGetValue(aggregateKey, out var aggregate) && !DataHelper.IsBlank(aggregate);
                }))
                .ToList();

            var result = new object[headerRows + visibleRows.Count, rowFieldCount + data.ColumnKeys.Count];

            for (var headerIndex = 0; headerIndex < data.ColumnFieldNames.Length; headerIndex++)
            {
                for (var columnIndex = 0; columnIndex < data.ColumnKeys.Count; columnIndex++)
                {
                    result[headerIndex, valueStartColumn + columnIndex] = data.ColumnKeys[columnIndex].Labels[headerIndex];
                }
            }

            for (var i = 0; i < rowFieldCount; i++)
            {
                result[headerRows - 1, i] = data.RowFieldNames[i];
            }

            for (var rowIndex = 0; rowIndex < visibleRows.Count; rowIndex++)
            {
                var rowKey = visibleRows[rowIndex];
                var targetRow = headerRows + rowIndex;
                for (var labelIndex = 0; labelIndex < rowFieldCount; labelIndex++)
                {
                    result[targetRow, labelIndex] = rowKey.Labels[labelIndex];
                }

                for (var columnIndex = 0; columnIndex < data.ColumnKeys.Count; columnIndex++)
                {
                    var aggregateKey = BuildAggregateKey(rowKey.Key, data.ColumnKeys[columnIndex].Key);
                    result[targetRow, valueStartColumn + columnIndex] = data.Aggregates.TryGetValue(aggregateKey, out var aggregate)
                        ? aggregate
                        : string.Empty;
                }
            }

            return result;
        }

        private static bool TryReadPivotData(object rowsInput, object columnsInput, MetricSpec spec, out PivotData data, out object? error)
        {
            data = null!;

            if (!TryReadCategoryMatrix(rowsInput, allowBlankInput: false, out var rowMatrix, out error))
            {
                return false;
            }

            if (!TryReadCategoryMatrix(columnsInput, allowBlankInput: true, out var columnMatrix, out error))
            {
                return false;
            }

            var rowCount = rowMatrix.GetLength(0);
            if (rowCount < 2 || (columnMatrix.GetLength(1) > 0 && columnMatrix.GetLength(0) != rowCount) || spec.RawValues.GetLength(0) != rowCount)
            {
                error = ExcelErrors.Length;
                return false;
            }

            var rowFieldNames = GetFieldNames(rowMatrix, GetLocalizedRowPrefix());
            var columnFieldNames = columnMatrix.GetLength(1) == 0
                ? new[] { GetLocalizedTotalLabel() }
                : GetFieldNames(columnMatrix, GetLocalizedColumnPrefix());

            spec.SourceName = GetValueFieldName(spec.RawValues);
            spec.DisplayLabel = BuildMetricDisplayLabel(spec);

            var rowKeys = new List<KeyLabel>();
            var rowSeen = new HashSet<string>();
            var columnKeys = new List<KeyLabel>();
            var columnSeen = new HashSet<string>();
            var buckets = new Dictionary<string, Bucket>();
            var aggregates = new Dictionary<string, object>();

            if (columnMatrix.GetLength(1) == 0)
            {
                var total = new KeyLabel(string.Empty, new[] { GetLocalizedTotalLabel() });
                columnKeys.Add(total);
                columnSeen.Add(total.Key);
            }

            for (var row = 1; row < rowCount; row++)
            {
                var rowLabels = GetLabels(rowMatrix, row);
                var rowKey = SerializeKey(rowLabels);
                if (rowSeen.Add(rowKey))
                {
                    rowKeys.Add(new KeyLabel(rowKey, rowLabels));
                }

                var columnLabels = columnMatrix.GetLength(1) == 0 ? new[] { GetLocalizedTotalLabel() } : GetLabels(columnMatrix, row);
                var columnKey = columnMatrix.GetLength(1) == 0 ? string.Empty : SerializeKey(columnLabels);
                if (columnSeen.Add(columnKey))
                {
                    columnKeys.Add(new KeyLabel(columnKey, columnLabels));
                }

                var aggregateKey = BuildAggregateKey(rowKey, columnKey);
                if (!buckets.TryGetValue(aggregateKey, out var bucket))
                {
                    bucket = new Bucket();
                    buckets.Add(aggregateKey, bucket);
                }

                var cell = spec.RawValues[row, 0];
                if (DataHelper.IsBlank(cell))
                {
                    continue;
                }

                if (spec.Calculation == PivotCalculation.Count)
                {
                    bucket.Count++;
                    continue;
                }

                if (!DataHelper.TryGetDouble(cell, out var numericValue))
                {
                    error = ExcelErrors.Value;
                    return false;
                }

                bucket.NumericValues.Add(numericValue);
            }

            SortKeyLabels(rowKeys);
            SortKeyLabels(columnKeys);

            foreach (var rowKey in rowKeys)
            {
                foreach (var columnKey in columnKeys)
                {
                    var aggregateKey = BuildAggregateKey(rowKey.Key, columnKey.Key);
                    if (!buckets.TryGetValue(aggregateKey, out var bucket))
                    {
                        continue;
                    }

                    var aggregate = ComputeAggregate(spec, bucket);
                    if (aggregate is ErrorValue errorValue)
                    {
                        error = errorValue.Value;
                        return false;
                    }

                    aggregates[aggregateKey] = aggregate!;
                }
            }

            if (rowKeys.Count == 0)
            {
                error = ExcelErrors.Count;
                return false;
            }

            data = new PivotData(rowFieldNames, columnFieldNames, rowKeys, columnKeys, aggregates);
            error = null;
            return true;
        }

        private static object? ComputeAggregate(MetricSpec spec, Bucket bucket)
        {
            if (spec.Calculation == PivotCalculation.Count)
            {
                return (double)bucket.Count;
            }

            var values = bucket.NumericValues;
            if (values.Count == 0)
            {
                return string.Empty;
            }

            return spec.Calculation switch
            {
                PivotCalculation.Sum => values.Sum(),
                PivotCalculation.Average => values.Average(),
                PivotCalculation.Min => values.Min(),
                PivotCalculation.Max => values.Max(),
                PivotCalculation.Median => StatisticsHelper.Median(values),
                PivotCalculation.Percentile => ComputeInclusivePercentile(values, spec.Detail!.Value),
                PivotCalculation.StdevSample => values.Count < 2 ? new ErrorValue(ExcelErrors.Count) : StatisticsHelper.SampleStandardDeviation(values),
                PivotCalculation.StdevPopulation => PopulationStandardDeviation(values),
                PivotCalculation.VarianceSample => values.Count < 2 ? new ErrorValue(ExcelErrors.Count) : StatisticsHelper.SampleVariance(values),
                PivotCalculation.VariancePopulation => PopulationVariance(values),
                PivotCalculation.VarCoefSample => ComputeVarCoef(values, sample: true),
                PivotCalculation.VarCoefPopulation => ComputeVarCoef(values, sample: false),
                PivotCalculation.ConfidenceT => ComputeConfidenceT(values, spec.Detail!.Value, spec.Direction),
                PivotCalculation.ConfidenceNorm => ComputeConfidenceNorm(values, spec.Detail!.Value, spec.Direction),
                PivotCalculation.Mad => ComputeMad(values),
                PivotCalculation.Iqr => ComputeInclusivePercentile(values, 0.75) - ComputeInclusivePercentile(values, 0.25),
                _ => string.Empty
            };
        }

        private static object ComputeVarCoef(IReadOnlyList<double> values, bool sample)
        {
            if (sample && values.Count < 2)
            {
                return new ErrorValue(ExcelErrors.Count);
            }

            var mean = values.Average();
            if (Math.Abs(mean) < 1e-12)
            {
                return new ErrorValue(ExcelErrors.DivZero);
            }

            var variance = sample ? StatisticsHelper.SampleVariance(values) : PopulationVariance(values);
            return Math.Sqrt(variance) / mean;
        }

        private static object ComputeConfidenceT(IReadOnlyList<double> values, double alpha, double direction)
        {
            if (values.Count < 2)
            {
                return new ErrorValue(ExcelErrors.Count);
            }

            var sd = StatisticsHelper.SampleStandardDeviation(values);
            var probability = Math.Abs(direction) < 1e-12 ? 1 - (alpha / 2.0) : 1 - alpha;
            var crit = StudentT.InvCDF(0, 1, values.Count - 1, probability);
            return crit * sd / Math.Sqrt(values.Count);
        }

        private static object ComputeConfidenceNorm(IReadOnlyList<double> values, double alpha, double direction)
        {
            if (values.Count < 2)
            {
                return new ErrorValue(ExcelErrors.Count);
            }

            var sd = StatisticsHelper.SampleStandardDeviation(values);
            var probability = Math.Abs(direction) < 1e-12 ? 1 - (alpha / 2.0) : 1 - alpha;
            var crit = Normal.InvCDF(0, 1, probability);
            return crit * sd / Math.Sqrt(values.Count);
        }

        private static double ComputeMad(IReadOnlyList<double> values)
        {
            var median = StatisticsHelper.Median(values);
            var deviations = values.Select(value => Math.Abs(value - median)).ToArray();
            return StatisticsHelper.Median(deviations);
        }

        private static double ComputeInclusivePercentile(IReadOnlyList<double> values, double quantile)
        {
            var sorted = values.OrderBy(value => value).ToArray();
            if (sorted.Length == 1)
            {
                return sorted[0];
            }

            var position = quantile * (sorted.Length - 1);
            var lowerIndex = (int)Math.Floor(position);
            var upperIndex = (int)Math.Ceiling(position);
            if (lowerIndex == upperIndex)
            {
                return sorted[lowerIndex];
            }

            var fraction = position - lowerIndex;
            return sorted[lowerIndex] + (fraction * (sorted[upperIndex] - sorted[lowerIndex]));
        }

        private static double PopulationVariance(IReadOnlyList<double> values)
        {
            var mean = values.Average();
            return values.Sum(value => Math.Pow(value - mean, 2)) / values.Count;
        }

        private static double PopulationStandardDeviation(IReadOnlyList<double> values)
            => Math.Sqrt(PopulationVariance(values));

        private static object[,]? NormalizeValueColumn(object values)
        {
            if (values is object[,] rawMatrix)
            {
                return rawMatrix.GetLength(1) == 1 ? rawMatrix : null;
            }

            var flattened = DataHelper.Flatten(values);
            if (flattened.Length == 0)
            {
                return null;
            }

            var matrix = new object[flattened.Length, 1];
            for (var row = 0; row < flattened.Length; row++)
            {
                matrix[row, 0] = flattened[row];
            }

            return matrix;
        }

        private static bool TryReadCategoryMatrix(object? input, bool allowBlankInput, out object[,] matrix, out object? error)
        {
            if (allowBlankInput && input is null or ExcelMissing or ExcelEmpty)
            {
                matrix = new object[1, 0];
                error = null;
                return true;
            }

            if (allowBlankInput && input is string text && string.IsNullOrWhiteSpace(text))
            {
                matrix = new object[1, 0];
                error = null;
                return true;
            }

            if (input is object[,] rawMatrix)
            {
                if (!allowBlankInput && rawMatrix.GetLength(1) < 1)
                {
                    matrix = new object[0, 0];
                    error = ExcelErrors.Count;
                    return false;
                }

                matrix = rawMatrix;
                error = null;
                return true;
            }

            var flattened = DataHelper.Flatten(input);
            if (allowBlankInput && flattened.All(DataHelper.IsBlank))
            {
                matrix = new object[Math.Max(1, flattened.Length), 0];
                error = null;
                return true;
            }

            if (!allowBlankInput && flattened.Length == 0)
            {
                matrix = new object[0, 0];
                error = ExcelErrors.Count;
                return false;
            }

            matrix = new object[flattened.Length, 1];
            for (var row = 0; row < flattened.Length; row++)
            {
                matrix[row, 0] = flattened[row];
            }

            error = null;
            return true;
        }

        private static string[] GetFieldNames(object[,] matrix, string prefix)
        {
            var names = new string[matrix.GetLength(1)];
            for (var index = 0; index < matrix.GetLength(1); index++)
            {
                var value = Convert.ToString(matrix[0, index], CultureInfo.CurrentCulture);
                names[index] = string.IsNullOrWhiteSpace(value) ? $"{prefix} {index + 1}" : value!;
            }

            return names;
        }

        private static string GetValueFieldName(object[,] matrix)
        {
            var label = Convert.ToString(matrix[0, 0], CultureInfo.CurrentCulture);
            return string.IsNullOrWhiteSpace(label) ? GetLocalizedValueLabel(1) : label!;
        }

        private static string[] GetLabels(object[,] matrix, int row)
        {
            var labels = new string[matrix.GetLength(1)];
            for (var col = 0; col < matrix.GetLength(1); col++)
            {
                var cell = matrix[row, col];
                labels[col] = DataHelper.IsBlank(cell)
                    ? GetLocalizedBlankLabel()
                    : Convert.ToString(cell, CultureInfo.CurrentCulture) ?? GetLocalizedBlankLabel();
            }

            return labels;
        }

        private static string BuildMetricDisplayLabel(MetricSpec spec)
        {
            var localizedName = GetLocalizedCalculationLabel(spec.Calculation);
            var formattedDetail = spec.Detail?.ToString("0.###", CultureInfo.CurrentCulture);

            return spec.Calculation switch
            {
                PivotCalculation.Percentile => $"{spec.SourceName} (p{formattedDetail})",
                PivotCalculation.ConfidenceT or PivotCalculation.ConfidenceNorm => $"{spec.SourceName} ({GetLocalizedAlphaLabel()}={formattedDetail}; {GetLocalizedDirectionLabel(spec.Direction)})",
                _ => $"{spec.SourceName} ({localizedName})"
            };
        }

        private static bool ValidateDetail(PivotCalculation calculation, double? detail)
        {
            if (calculation is not (PivotCalculation.Percentile or PivotCalculation.ConfidenceT or PivotCalculation.ConfidenceNorm))
            {
                return true;
            }

            return detail.HasValue && detail.Value > 0.0 && detail.Value < 1.0;
        }

        private static bool ValidateDirection(PivotCalculation calculation, double direction)
        {
            if (calculation is not (PivotCalculation.ConfidenceT or PivotCalculation.ConfidenceNorm))
            {
                return Math.Abs(direction) < 1e-12;
            }

            return Math.Abs(direction) < 1e-12 || Math.Abs(direction - 1.0) < 1e-12 || Math.Abs(direction + 1.0) < 1e-12;
        }

        private static void SortKeyLabels(List<KeyLabel> labels)
            => labels.Sort(CompareKeyLabels);

        private static int CompareKeyLabels(KeyLabel? left, KeyLabel? right)
        {
            if (left is null && right is null)
            {
                return 0;
            }

            if (left is null)
            {
                return -1;
            }

            if (right is null)
            {
                return 1;
            }

            var compareInfo = CultureInfo.CurrentCulture.CompareInfo;
            var sharedLength = Math.Min(left.Labels.Length, right.Labels.Length);
            for (var i = 0; i < sharedLength; i++)
            {
                var comparison = compareInfo.Compare(left.Labels[i], right.Labels[i], CompareOptions.StringSort);
                if (comparison != 0)
                {
                    return comparison;
                }
            }

            return left.Labels.Length.CompareTo(right.Labels.Length);
        }

        private static string BuildAggregateKey(string rowKey, string columnKey)
            => $"{rowKey}\u001E{columnKey}";

        private static string SerializeKey(IReadOnlyList<string> labels)
            => string.Join("\u001F", labels);

        private static string GetLocalizedCalculationLabel(PivotCalculation calculation)
        {
            if (IsCzechCulture())
            {
                return calculation switch
                {
                    PivotCalculation.Count => "počet",
                    PivotCalculation.Sum => "součet",
                    PivotCalculation.Average => "průměr",
                    PivotCalculation.Min => "minimum",
                    PivotCalculation.Max => "maximum",
                    PivotCalculation.Median => "median",
                    PivotCalculation.Percentile => "percentil",
                    PivotCalculation.StdevSample => "výběrová směrodatná odchylka",
                    PivotCalculation.StdevPopulation => "populační směrodatná odchylka",
                    PivotCalculation.VarianceSample => "výběrový rozptyl",
                    PivotCalculation.VariancePopulation => "populační rozptyl",
                    PivotCalculation.VarCoefSample => "výběrový variační koeficient",
                    PivotCalculation.VarCoefPopulation => "populační variační koeficient",
                    PivotCalculation.ConfidenceT => "interval spolehlivosti t",
                    PivotCalculation.ConfidenceNorm => "interval spolehlivosti norm",
                    PivotCalculation.Mad => "MAD",
                    PivotCalculation.Iqr => "IQR",
                    _ => "počet"
                };
            }

            return calculation switch
            {
                PivotCalculation.Count => "count",
                PivotCalculation.Sum => "sum",
                PivotCalculation.Average => "average",
                PivotCalculation.Min => "minimum",
                PivotCalculation.Max => "maximum",
                PivotCalculation.Median => "median",
                PivotCalculation.Percentile => "percentile",
                PivotCalculation.StdevSample => "sample standard deviation",
                PivotCalculation.StdevPopulation => "population standard deviation",
                PivotCalculation.VarianceSample => "sample variance",
                PivotCalculation.VariancePopulation => "population variance",
                PivotCalculation.VarCoefSample => "sample coefficient of variation",
                PivotCalculation.VarCoefPopulation => "population coefficient of variation",
                PivotCalculation.ConfidenceT => "t confidence interval",
                PivotCalculation.ConfidenceNorm => "normal confidence interval",
                PivotCalculation.Mad => "MAD",
                PivotCalculation.Iqr => "IQR",
                _ => "count"
            };
        }

        private static string GetLocalizedDirectionLabel(double direction)
        {
            if (Math.Abs(direction) < 1e-12)
            {
                return IsCzechCulture() ? "oboustranný" : "two-sided";
            }

            if (direction < 0)
            {
                return IsCzechCulture() ? "levostranný" : "left-sided";
            }

            return IsCzechCulture() ? "pravostranný" : "right-sided";
        }

        private static string GetLocalizedRowPrefix() => IsCzechCulture() ? "Radek" : "Row";
        private static string GetLocalizedColumnPrefix() => IsCzechCulture() ? "Sloupec" : "Column";
        private static string GetLocalizedBlankLabel() => IsCzechCulture() ? "(prázdné)" : "(blank)";
        private static string GetLocalizedValueLabel(int ordinal) => IsCzechCulture() ? $"Hodnota {ordinal}" : $"Value {ordinal}";
        private static string GetLocalizedAlphaLabel() => IsCzechCulture() ? "alfa" : "alpha";
        private static string GetLocalizedTotalLabel() => IsCzechCulture() ? "Celkem" : "Total";
        private static bool IsCzechCulture() => string.Equals(CultureInfo.CurrentCulture.TwoLetterISOLanguageName, "cs", StringComparison.OrdinalIgnoreCase);
        private static object[,] SpillError(object error) => new object[,] { { error, string.Empty } };

        private sealed class MetricSpec
        {
            public MetricSpec(object[,] rawValues, PivotCalculation calculation, double? detail, double direction)
            {
                RawValues = rawValues;
                Calculation = calculation;
                Detail = detail;
                Direction = direction;
            }

            public object[,] RawValues { get; }
            public PivotCalculation Calculation { get; }
            public double? Detail { get; }
            public double Direction { get; }
            public string SourceName { get; set; } = string.Empty;
            public string DisplayLabel { get; set; } = string.Empty;
        }

        private sealed class PivotData
        {
            public PivotData(string[] rowFieldNames, string[] columnFieldNames, List<KeyLabel> rowKeys, List<KeyLabel> columnKeys, Dictionary<string, object> aggregates)
            {
                RowFieldNames = rowFieldNames;
                ColumnFieldNames = columnFieldNames;
                RowKeys = rowKeys;
                ColumnKeys = columnKeys;
                Aggregates = aggregates;
            }

            public string[] RowFieldNames { get; }
            public string[] ColumnFieldNames { get; }
            public List<KeyLabel> RowKeys { get; }
            public List<KeyLabel> ColumnKeys { get; }
            public Dictionary<string, object> Aggregates { get; }
        }

        private sealed class KeyLabel
        {
            public KeyLabel(string key, string[] labels)
            {
                Key = key;
                Labels = labels;
            }

            public string Key { get; }
            public string[] Labels { get; }
        }

        private sealed class Bucket
        {
            public int Count { get; set; }
            public List<double> NumericValues { get; } = new List<double>();
        }

        private sealed class ErrorValue
        {
            public ErrorValue(object value) => Value = value;
            public object Value { get; }
        }

        private enum PivotCalculation
        {
            Count = 0,
            Sum = 1,
            Average = 2,
            Min = 3,
            Max = 4,
            Median = 5,
            Percentile = 6,
            StdevSample = 7,
            StdevPopulation = 8,
            VarianceSample = 9,
            VariancePopulation = 10,
            VarCoefSample = 11,
            VarCoefPopulation = 12,
            ConfidenceT = 13,
            ConfidenceNorm = 14,
            Mad = 15,
            Iqr = 16
        }
    }
}
