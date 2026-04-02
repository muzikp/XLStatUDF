/// <summary>
/// Implementuje Spearmanův korelační koeficient a jeho test významnosti přes t-aproximaci.
/// </summary>
namespace XLStatUDF.Functions.Correlation
{
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class SpearmanCorrel
    {
        [ExcelFunction(Name = "CORREL.SPEARMAN", Description = "[Korelace] Spearmanův korelační koeficient", Category = "XLStatUDF")]
        public static object[,] Run(
            [ExcelArgument(Name = "rozsah_x", Description = "První proměnná")] object xValues,
            [ExcelArgument(Name = "rozsah_y", Description = "Druhá proměnná")] object yValues,
            [ExcelArgument(Name = "smer", Description = "Směr testu: 0=two, 1=left, 2=right")] object? direction = null,
            [ExcelArgument(Name = "alpha", Description = "Hladina významnosti")] double alpha = 0.05,
            [ExcelArgument(Name = "ma_zahlavi", Description = "Volitelne: 0=autodetect, 1=ma zahlavi, 2=nema zahlavi")] object? hasHeader = null)
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

            if (!DataHelper.TryReadPairedNumericVectors(xValues, yValues, out var x, out var y, out var error, parsedHeaderMode))
            {
                return SpillError(error!);
            }

            if (x.Length < 3)
            {
                return SpillError(ExcelErrors.Count);
            }

            var xRanks = RankHelper.MidRank(x);
            var yRanks = RankHelper.MidRank(y);

            var varianceX = xRanks.Sum(rank => Math.Pow(rank - xRanks.Average(), 2));
            var varianceY = yRanks.Sum(rank => Math.Pow(rank - yRanks.Average(), 2));
            if (varianceX == 0 || varianceY == 0)
            {
                return SpillError(ExcelErrors.Num);
            }

            var rho = StatisticsHelper.PearsonCorrelation(xRanks, yRanks);
            var df = x.Length - 2;
            var t = Math.Abs(1 - Math.Abs(rho)) < 1e-12
                ? Math.CopySign(double.PositiveInfinity, rho)
                : rho * Math.Sqrt(df) / Math.Sqrt(1 - (rho * rho));
            var critical = TestHelper.CriticalT(alpha, df, parsedDirection);
            var p = TestHelper.PValueFromT(t, df, parsedDirection);

            var rows = new List<object[]>();
            SpillBuilder.AddRow(rows, "ρ", rho);
            SpillBuilder.AddRow(rows, "n", x.Length);
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
