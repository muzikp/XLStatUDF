/// <summary>
/// Implementuje jednovýběrový t-test s podporou jednostranné i oboustranné alternativy.
/// </summary>
namespace XLStatUDF.Functions.Tests
{
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class TTest1S
    {
        [ExcelFunction(Name = "T.TEST.1S", Description = "Jednovýběrový t-test", Category = FunctionCategories.Tests)]
        public static object[,] Run(
            [ExcelArgument(Name = "hodnoty", Description = "Výběrová data; prázdné buňky jsou ignorovány")] object values,
            [ExcelArgument(Name = "mu_0", Description = "Hypotetická střední hodnota")] double mu0,
            [ExcelArgument(Name = "smer", Description = "Směr testu: 0=two, 1=left, 2=right")] object? direction = null,
            [ExcelArgument(Name = "alpha", Description = "Hladina významnosti")] double alpha = 0.05,
            [ExcelArgument(Name = "ma_záhlaví", Description = "Volitelně: 0=autodetect, 1=má záhlaví, 2=nemá záhlaví")] object? hasHeader = null)
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

            if (!DataHelper.TryReadNumericVector(values, out var sample, out var error, parsedHeaderMode))
            {
                return SpillError(error!);
            }

            if (sample.Length < 2)
            {
                return SpillError(ExcelErrors.Count);
            }

            var mean = StatisticsHelper.Mean(sample);
            var s = StatisticsHelper.SampleStandardDeviation(sample);
            var n = sample.Length;
            var df = n - 1;
            var t = (mean - mu0) / (s / Math.Sqrt(n));
            var critical = TestHelper.CriticalT(alpha, df, parsedDirection);
            var p = TestHelper.PValueFromT(t, df, parsedDirection);

            var rows = new List<object[]>();
            SpillBuilder.AddRow(rows, "x̄", mean);
            SpillBuilder.AddRow(rows, "μ₀", mu0);
            SpillBuilder.AddRow(rows, "sₓ", s);
            SpillBuilder.AddRow(rows, "n", n);
            SpillBuilder.AddRow(rows, "α", alpha);
            SpillBuilder.AddRow(rows, "t", t);
            SpillBuilder.AddRow(rows, "df", df);
            SpillBuilder.AddRow(rows, CriticalValues.LabelForDirection("t", parsedDirection), critical);
            SpillBuilder.AddRow(rows, "p", p);
            return SpillBuilder.Build(rows);
        }

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
