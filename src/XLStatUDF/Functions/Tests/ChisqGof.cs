/// <summary>
/// Implementuje chí-kvadrát test dobré shody včetně příspěvků jednotlivých kategorií.
/// </summary>
namespace XLStatUDF.Functions.Tests
{
    using System.Globalization;
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using XLStatUDF.Helpers;

    public static class ChisqGof
    {
        [ExcelFunction(Name = "CHISQ.GOF", Description = "Chí-kvadrát test dobré shody", Category = FunctionCategories.Tests)]
        public static object[,] Run(
            [ExcelArgument(Name = "pozorované", Description = "Pozorované četnosti kategorií")] object observed,
            [ExcelArgument(Name = "očekávané", Description = "Očekávané četnosti nebo pravděpodobnosti")] object expected,
            [ExcelArgument(Name = "kategorie", Description = "Volitelné názvy kategorií")] object? categories = null,
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

            if (!DataHelper.TryReadNumericVector(observed, out var observedValues, out var error, parsedHeaderMode)
                || !DataHelper.TryReadNumericVector(expected, out var expectedValues, out error, parsedHeaderMode))
            {
                return SpillError(error!);
            }

            if (observedValues.Length != expectedValues.Length)
            {
                return SpillError(ExcelErrors.Length);
            }

            if (observedValues.Length < 2)
            {
                return SpillError(ExcelErrors.Count);
            }

            if (observedValues.Any(value => value < 0 || Math.Abs(value - Math.Round(value)) > 1e-9))
            {
                return SpillError(ExcelErrors.Num);
            }

            if (expectedValues.Any(value => value <= 0))
            {
                return SpillError(ExcelErrors.Num);
            }

            var categoryLabels = ResolveCategories(categories, observedValues.Length, parsedHeaderMode, out error);
            if (error is not null)
            {
                return SpillError(error);
            }

            var totalObserved = observedValues.Sum();
            var totalExpected = expectedValues.Sum();
            var expectedCounts = Math.Abs(totalExpected - 1.0) < 1e-9
                ? expectedValues.Select(value => value * totalObserved).ToArray()
                : expectedValues.ToArray();

            if (expectedCounts.Any(value => value <= 0))
            {
                return SpillError(ExcelErrors.Num);
            }

            var contributions = observedValues.Zip(expectedCounts, (o, e) => Math.Pow(o - e, 2) / e).ToArray();
            var chiSquare = contributions.Sum();
            var df = observedValues.Length - 1;
            var critical = ChiSquared.InvCDF(df, 1 - alpha);
            var p = 1 - ChiSquared.CDF(df, chiSquare);

            var rows = new List<object[]>
            {
                new object[] { "χ² GOF", "", "", "" },
                new object[] { "χ²", chiSquare, "", "" },
                new object[] { "df", df, "", "" },
                new object[] { "α", alpha, "", "" },
                new object[] { "χ²α", critical, "", "" },
                new object[] { "p", p, "", "" },
                new object[] { "", "", "", "" },
                new object[] { "KATEGORIE", "", "", "" },
                new object[] { "Kategorie", "O", "E", "(O−E)²/E" }
            };

            for (var i = 0; i < observedValues.Length; i++)
            {
                rows.Add(new object[] { categoryLabels[i], observedValues[i], expectedCounts[i], contributions[i] });
            }

            return SpillBuilder.Build(rows);
        }

        private static string[] ResolveCategories(object? categories, int expectedLength, HeaderMode headerMode, out object? error)
        {
            if (categories is null || categories is ExcelMissing || categories is ExcelEmpty)
            {
                error = null;
                return Enumerable.Range(1, expectedLength).Select(index => index.ToString(CultureInfo.InvariantCulture)).ToArray();
            }

            var raw = DataHelper.Flatten(categories);
            if (headerMode == HeaderMode.HasHeader && raw.Length > 0)
            {
                raw = raw.Skip(1).ToArray();
            }
            else if (headerMode == HeaderMode.AutoDetect && raw.Length == expectedLength + 1)
            {
                raw = raw.Skip(1).ToArray();
            }
            if (raw.Length != expectedLength)
            {
                error = ExcelErrors.Length;
                return [];
            }

            error = null;
            return raw.Select((item, index) =>
                    DataHelper.IsBlank(item)
                        ? (index + 1).ToString(CultureInfo.InvariantCulture)
                        : Convert.ToString(item, CultureInfo.InvariantCulture) ?? (index + 1).ToString(CultureInfo.InvariantCulture))
                .ToArray();
        }

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
