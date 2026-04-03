/// <summary>
/// Implementuje analýzu kontingenčních tabulek jak z hotové tabulky, tak z groupovaných dat.
/// </summary>
namespace XLStatUDF.Functions.Tests
{
    using System.Globalization;
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using XLStatUDF.Helpers;

    public static class Contingency
    {
        [ExcelFunction(Name = "CONTINGENCY.T", Description = "Analýza kontingenční tabulky zadané přímo jako matice", Category = FunctionCategories.Tests)]
        public static object[,] RunTable(
            [ExcelArgument(Name = "tabulka", Description = "Kontingenční tabulka; volitelně i s horním řádkem a levým sloupcem popisků")] object table,
            [ExcelArgument(Name = "ma_záhlaví", Description = "Volitelně: 0=autodetect, 1=horni radek i levy sloupec jsou popisky, 2=bez popisku")] object? hasHeader = null,
            [ExcelArgument(Name = "alpha", Description = "Hladina významnosti")] double alpha = 0.05)
        {
            if (!TestHelper.IsValidAlpha(alpha))
            {
                return SpillError(ExcelErrors.Num);
            }

            if (!ArgumentHelper.TryParseHeaderMode(hasHeader, out var parsedHeaderMode))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!TryReadObservedTable(table, parsedHeaderMode, out var observed, out var rowLabels, out var columnLabels, out var error))
            {
                return SpillError(error!);
            }

            return BuildReport(observed, rowLabels, columnLabels, alpha);
        }

        [ExcelFunction(Name = "CONTINGENCY.G", Description = "Analýza kontingenční tabulky z groupovaných sloupců", Category = FunctionCategories.Tests)]
        public static object[,] RunGrouped(
            [ExcelArgument(Name = "sloupce", Description = "Kategorie sloupců")] object columnCategories,
            [ExcelArgument(Name = "řádky", Description = "Kategorie řádků")] object rowCategories,
            [ExcelArgument(Name = "počet", Description = "Volitelně: cetnosti dvojic; kdyz chybi, kazdy zaznam ma vahu 1")] object? counts = null,
            [ExcelArgument(Name = "alpha", Description = "Hladina významnosti")] double alpha = 0.05,
            [ExcelArgument(Name = "ma_záhlaví", Description = "Volitelně: 0=autodetect, 1=má záhlaví, 2=nemá záhlaví")] object? hasHeader = null)
        {
            if (!TestHelper.IsValidAlpha(alpha))
            {
                return SpillError(ExcelErrors.Num);
            }

            if (!ArgumentHelper.TryParseHeaderMode(hasHeader, out var parsedHeaderMode))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!TryReadGroupedTable(columnCategories, rowCategories, counts, parsedHeaderMode, out var observed, out var rowLabels, out var columnLabels, out var error))
            {
                return SpillError(error!);
            }

            return BuildReport(observed, rowLabels, columnLabels, alpha);
        }

        private static bool TryReadObservedTable(
            object input,
            HeaderMode headerMode,
            out double[,] observed,
            out string[] rowLabels,
            out string[] columnLabels,
            out object? error)
        {
            if (input is not object[,] matrix)
            {
                observed = new double[0, 0];
                rowLabels = [];
                columnLabels = [];
                error = ExcelErrors.Value;
                return false;
            }

            var rows = matrix.GetLength(0);
            var cols = matrix.GetLength(1);
            if (rows < 1 || cols < 1)
            {
                observed = new double[0, 0];
                rowLabels = [];
                columnLabels = [];
                error = ExcelErrors.Value;
                return false;
            }

            var actualHeaderMode = headerMode == HeaderMode.AutoDetect && HasTableHeaders(matrix)
                ? HeaderMode.HasHeader
                : headerMode;

            var dataRowOffset = actualHeaderMode == HeaderMode.HasHeader ? 1 : 0;
            var dataColOffset = actualHeaderMode == HeaderMode.HasHeader ? 1 : 0;
            var dataRows = rows - dataRowOffset;
            var dataCols = cols - dataColOffset;

            if (dataRows < 2 || dataCols < 2)
            {
                observed = new double[0, 0];
                rowLabels = [];
                columnLabels = [];
                error = ExcelErrors.Count;
                return false;
            }

            observed = new double[dataRows, dataCols];
            for (var r = 0; r < dataRows; r++)
            {
                for (var c = 0; c < dataCols; c++)
                {
                    var value = matrix[r + dataRowOffset, c + dataColOffset];
                    if (!DataHelper.TryGetDouble(value, out var parsed)
                        || parsed < 0
                        || Math.Abs(parsed - Math.Round(parsed)) > 1e-9)
                    {
                        rowLabels = [];
                        columnLabels = [];
                        error = ExcelErrors.Value;
                        return false;
                    }

                    observed[r, c] = parsed;
                }
            }

            rowLabels = Enumerable.Range(1, dataRows)
                .Select(index => index.ToString(CultureInfo.InvariantCulture))
                .ToArray();
            columnLabels = Enumerable.Range(1, dataCols)
                .Select(index => index.ToString(CultureInfo.InvariantCulture))
                .ToArray();

            if (actualHeaderMode == HeaderMode.HasHeader)
            {
                rowLabels = Enumerable.Range(0, dataRows)
                    .Select(index => Convert.ToString(matrix[index + 1, 0], CultureInfo.InvariantCulture) ?? (index + 1).ToString(CultureInfo.InvariantCulture))
                    .ToArray();
                columnLabels = Enumerable.Range(0, dataCols)
                    .Select(index => Convert.ToString(matrix[0, index + 1], CultureInfo.InvariantCulture) ?? (index + 1).ToString(CultureInfo.InvariantCulture))
                    .ToArray();
            }

            error = null;
            return true;
        }

        private static bool HasTableHeaders(object[,] matrix)
        {
            var rows = matrix.GetLength(0);
            var cols = matrix.GetLength(1);
            if (rows < 3 || cols < 3)
            {
                return false;
            }

            if (DataHelper.TryGetDouble(matrix[0, 0], out _))
            {
                return false;
            }

            var topHeaderLike = false;
            for (var c = 1; c < cols; c++)
            {
                if (!DataHelper.IsBlank(matrix[0, c]) && !DataHelper.TryGetDouble(matrix[0, c], out _))
                {
                    topHeaderLike = true;
                    break;
                }
            }

            var leftHeaderLike = false;
            for (var r = 1; r < rows; r++)
            {
                if (!DataHelper.IsBlank(matrix[r, 0]) && !DataHelper.TryGetDouble(matrix[r, 0], out _))
                {
                    leftHeaderLike = true;
                    break;
                }
            }

            if (!topHeaderLike || !leftHeaderLike)
            {
                return false;
            }

            for (var r = 1; r < rows; r++)
            {
                for (var c = 1; c < cols; c++)
                {
                    if (!DataHelper.TryGetDouble(matrix[r, c], out _))
                    {
                        return false;
                    }
                }
            }

            return true;
        }

        private static bool TryReadGroupedTable(
            object columnsInput,
            object rowsInput,
            object? countsInput,
            HeaderMode headerMode,
            out double[,] observed,
            out string[] rowLabels,
            out string[] columnLabels,
            out object? error)
        {
            var rawColumns = DataHelper.Flatten(columnsInput);
            var rawRows = DataHelper.Flatten(rowsInput);
            var rawCounts = countsInput is null or ExcelMissing or ExcelEmpty
                ? null
                : DataHelper.Flatten(countsInput);

            if (rawColumns.Length != rawRows.Length || (rawCounts is not null && rawColumns.Length != rawCounts.Length))
            {
                observed = new double[0, 0];
                rowLabels = [];
                columnLabels = [];
                error = ExcelErrors.Length;
                return false;
            }

            if (headerMode == HeaderMode.HasHeader && rawColumns.Length > 0)
            {
                rawColumns = rawColumns.Skip(1).ToArray();
                rawRows = rawRows.Skip(1).ToArray();
                if (rawCounts is not null)
                {
                    rawCounts = rawCounts.Skip(1).ToArray();
                }
            }
            else if (headerMode == HeaderMode.AutoDetect && ShouldSkipGroupedHeader(rawColumns, rawRows, rawCounts))
            {
                rawColumns = rawColumns.Skip(1).ToArray();
                rawRows = rawRows.Skip(1).ToArray();
                if (rawCounts is not null)
                {
                    rawCounts = rawCounts.Skip(1).ToArray();
                }
            }

            var rowList = new List<string>();
            var columnList = new List<string>();
            var countsByPair = new Dictionary<(string Row, string Column), double>();

            for (var i = 0; i < rawColumns.Length; i++)
            {
                if (DataHelper.IsBlank(rawColumns[i]) || DataHelper.IsBlank(rawRows[i]) || (rawCounts is not null && DataHelper.IsBlank(rawCounts[i])))
                {
                    continue;
                }

                var columnLabel = Convert.ToString(rawColumns[i], CultureInfo.InvariantCulture);
                var rowLabel = Convert.ToString(rawRows[i], CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(columnLabel) || string.IsNullOrWhiteSpace(rowLabel))
                {
                    continue;
                }

                double count = 1.0;
                if (rawCounts is not null)
                {
                    if (!DataHelper.TryGetDouble(rawCounts[i], out count)
                        || count < 0
                        || Math.Abs(count - Math.Round(count)) > 1e-9)
                    {
                        observed = new double[0, 0];
                        rowLabels = [];
                        columnLabels = [];
                        error = ExcelErrors.Value;
                        return false;
                    }
                }

                if (!rowList.Contains(rowLabel, StringComparer.Ordinal))
                {
                    rowList.Add(rowLabel);
                }

                if (!columnList.Contains(columnLabel, StringComparer.Ordinal))
                {
                    columnList.Add(columnLabel);
                }

                var key = (rowLabel, columnLabel);
                countsByPair[key] = countsByPair.TryGetValue(key, out var existing) ? existing + count : count;
            }

            if (rowList.Count < 2 || columnList.Count < 2)
            {
                observed = new double[0, 0];
                rowLabels = [];
                columnLabels = [];
                error = ExcelErrors.Count;
                return false;
            }

            rowLabels = rowList.OrderBy(value => value, StringComparer.Ordinal).ToArray();
            columnLabels = columnList.OrderBy(value => value, StringComparer.Ordinal).ToArray();

            var sortedRowIndex = rowLabels.Select((label, index) => (label, index)).ToDictionary(x => x.label, x => x.index, StringComparer.Ordinal);
            var sortedColumnIndex = columnLabels.Select((label, index) => (label, index)).ToDictionary(x => x.label, x => x.index, StringComparer.Ordinal);
            observed = new double[rowLabels.Length, columnLabels.Length];

            foreach (var item in countsByPair)
            {
                observed[sortedRowIndex[item.Key.Row], sortedColumnIndex[item.Key.Column]] += item.Value;
            }

            error = null;
            return true;
        }

        private static bool ShouldSkipGroupedHeader(object[] columns, object[] rows, object[]? counts)
        {
            if (!DataHelper.ShouldSkipLeadingHeader(columns, item => !DataHelper.IsBlank(item))
                || !DataHelper.ShouldSkipLeadingHeader(rows, item => !DataHelper.IsBlank(item)))
            {
                return false;
            }

            return counts is null || DataHelper.ShouldSkipLeadingHeader(counts, item => DataHelper.TryGetDouble(item, out _));
        }

        private static object[,] BuildReport(double[,] observed, string[] rowLabels, string[] columnLabels, double alpha)
        {
            var rowCount = observed.GetLength(0);
            var colCount = observed.GetLength(1);
            var rowTotals = new double[rowCount];
            var colTotals = new double[colCount];
            double grandTotal = 0.0;

            for (var r = 0; r < rowCount; r++)
            {
                for (var c = 0; c < colCount; c++)
                {
                    rowTotals[r] += observed[r, c];
                    colTotals[c] += observed[r, c];
                    grandTotal += observed[r, c];
                }
            }

            if (grandTotal <= 0)
            {
                return SpillError(ExcelErrors.Num);
            }

            var expected = new double[rowCount, colCount];
            var chiSquare = 0.0;
            for (var r = 0; r < rowCount; r++)
            {
                for (var c = 0; c < colCount; c++)
                {
                    expected[r, c] = rowTotals[r] * colTotals[c] / grandTotal;
                    if (expected[r, c] > 0)
                    {
                        chiSquare += Math.Pow(observed[r, c] - expected[r, c], 2) / expected[r, c];
                    }
                }
            }

            var df = (rowCount - 1) * (colCount - 1);
            var critical = ChiSquared.InvCDF(df, 1 - alpha);
            var p = 1 - ChiSquared.CDF(df, chiSquare);
            var pearsonC = Math.Sqrt(chiSquare / (chiSquare + grandTotal));
            var cramerV = Math.Sqrt(chiSquare / (grandTotal * Math.Min(rowCount - 1, colCount - 1)));
            var phi = rowCount == 2 && colCount == 2
                ? Math.Sqrt(chiSquare / grandTotal)
                : double.NaN;

            var rows = new List<object[]>();

            SpillBuilder.AddHeader(rows, "POZOROVANÁ KONTINGENČNÍ TABULKA", colCount + 2);
            rows.Add(BuildTableHeader(columnLabels));
            for (var r = 0; r < rowCount; r++)
            {
                rows.Add(BuildObservedRow(rowLabels[r], observed, r, rowTotals[r]));
            }
            rows.Add(BuildTotalsRow(colTotals, grandTotal));
            SpillBuilder.AddSeparator(rows, colCount + 2);

            SpillBuilder.AddHeader(rows, "OČEKÁVANÉ ČETNOSTI", colCount + 2);
            rows.Add(BuildTableHeader(columnLabels));
            for (var r = 0; r < rowCount; r++)
            {
                rows.Add(BuildExpectedRow(rowLabels[r], expected, r, rowTotals[r]));
            }
            rows.Add(BuildTotalsRow(colTotals, grandTotal));
            SpillBuilder.AddSeparator(rows, colCount + 2);

            SpillBuilder.AddHeader(rows, "SOUHRN TESTU", 2);
            SpillBuilder.AddRow(rows, "n", grandTotal);
            SpillBuilder.AddRow(rows, "df", df);
            SpillBuilder.AddRow(rows, "α", alpha);
            SpillBuilder.AddRow(rows, "χ²", chiSquare);
            SpillBuilder.AddRow(rows, "χ²₁₋α", critical);
            SpillBuilder.AddRow(rows, "p", p);
            SpillBuilder.AddSeparator(rows, 2);

            SpillBuilder.AddHeader(rows, "MÍRY ASOCIACE", 2);
            SpillBuilder.AddRow(rows, "Pearson C", pearsonC);
            SpillBuilder.AddRow(rows, "Cramér V", cramerV);
            SpillBuilder.AddRow(rows, "phi", double.IsNaN(phi) ? string.Empty : phi);

            return SpillBuilder.Build(rows);
        }

        private static object[] BuildTableHeader(string[] columnLabels)
        {
            var row = new object[columnLabels.Length + 2];
            row[0] = string.Empty;
            for (var i = 0; i < columnLabels.Length; i++)
            {
                row[i + 1] = columnLabels[i];
            }

            row[^1] = "Σ";
            return row;
        }

        private static object[] BuildObservedRow(string rowLabel, double[,] observed, int rowIndex, double rowTotal)
        {
            var colCount = observed.GetLength(1);
            var row = new object[colCount + 2];
            row[0] = rowLabel;
            for (var c = 0; c < colCount; c++)
            {
                row[c + 1] = observed[rowIndex, c];
            }

            row[^1] = rowTotal;
            return row;
        }

        private static object[] BuildExpectedRow(string rowLabel, double[,] expected, int rowIndex, double rowTotal)
        {
            var colCount = expected.GetLength(1);
            var row = new object[colCount + 2];
            row[0] = rowLabel;
            for (var c = 0; c < colCount; c++)
            {
                row[c + 1] = expected[rowIndex, c];
            }

            row[^1] = rowTotal;
            return row;
        }

        private static object[] BuildTotalsRow(double[] colTotals, double grandTotal)
        {
            var row = new object[colTotals.Length + 2];
            row[0] = "Σ";
            for (var i = 0; i < colTotals.Length; i++)
            {
                row[i + 1] = colTotals[i];
            }

            row[^1] = grandTotal;
            return row;
        }

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
