/// <summary>
/// Implementuje jednovýběrový z-test pro populační podíl nad binárními daty.
/// </summary>
namespace XLStatUDF.Functions.Tests
{
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class PropTest1S
    {
        [ExcelFunction(Name = "PROP.TEST.1S", Description = "Jednovýběrový test podílu", Category = FunctionCategories.Tests)]
        public static object[,] Run(
            [ExcelArgument(Name = "hodnoty", Description = "Binární data 0/1 nebo TRUE/FALSE")] object values,
            [ExcelArgument(Name = "pi_0", Description = "Hypotetický populační podíl")] double pi0,
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

            if (!TestHelper.IsValidAlpha(alpha) || pi0 <= 0 || pi0 >= 1)
            {
                return SpillError(ExcelErrors.Num);
            }

            if (!DataHelper.TryReadBinaryVector(values, out var successes, out var count, out var error, parsedHeaderMode))
            {
                return SpillError(error!);
            }

            if (count < 1)
            {
                return SpillError(ExcelErrors.Count);
            }

            var pHat = successes / (double)count;
            var z = (pHat - pi0) / Math.Sqrt(pi0 * (1 - pi0) / count);
            var critical = TestHelper.CriticalZ(alpha, parsedDirection);
            var p = TestHelper.PValueFromZ(z, parsedDirection);

            var rows = new List<object[]>();
            SpillBuilder.AddRow(rows, "p̂", pHat);
            SpillBuilder.AddRow(rows, "π₀", pi0);
            SpillBuilder.AddRow(rows, "x", successes);
            SpillBuilder.AddRow(rows, "n", count);
            SpillBuilder.AddRow(rows, "α", alpha);
            SpillBuilder.AddRow(rows, "z", z);
            SpillBuilder.AddRow(rows, CriticalValues.LabelForDirection("z", parsedDirection), critical);
            SpillBuilder.AddRow(rows, "p", p);
            return SpillBuilder.Build(rows);
        }

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
