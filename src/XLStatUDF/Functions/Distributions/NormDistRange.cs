/// <summary>
/// Implementuje funkci NORM.DIST.RANGE pro pravděpodobnost intervalu normálního rozdělení.
/// </summary>
namespace XLStatUDF.Functions.Distributions
{
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using XLStatUDF.Helpers;

    public static class NormDistRange
    {
        [ExcelFunction(
            Name = "NORM.DIST.RANGE",
            Description = "Pravděpodobnost intervalu normálního rozdělení",
            Category = FunctionCategories.Tests)]
        public static object NormalDistributionRange(
            [ExcelArgument(Name = "střední_hodnota", Description = "Střední hodnota normálního rozdělení")] object mean,
            [ExcelArgument(Name = "směrodatná_odchylka", Description = "Směrodatná odchylka; musí být kladná")] object standardDeviation,
            [ExcelArgument(Name = "dolní_hranice", Description = "Dolní mez intervalu; prázdná buňka znamená minus nekonečno")] object lowerBound,
            [ExcelArgument(Name = "horní_hranice", Description = "Horní mez intervalu; prázdná buňka znamená plus nekonečno")] object upperBound)
        {
            if (!DataHelper.TryGetDouble(mean, out var mu)
                || !DataHelper.TryGetDouble(standardDeviation, out var sigma))
            {
                return ExcelErrors.Value;
            }

            if (!TryResolveBound(lowerBound, double.NegativeInfinity, out var lower)
                || !TryResolveBound(upperBound, double.PositiveInfinity, out var upper))
            {
                return ExcelErrors.Value;
            }

            if (sigma <= 0 || lower > upper)
            {
                return ExcelErrors.Num;
            }

            return Normal.CDF(mu, sigma, upper) - Normal.CDF(mu, sigma, lower);
        }

        private static bool TryResolveBound(object input, double defaultValue, out double value)
        {
            if (DataHelper.IsBlank(input))
            {
                value = defaultValue;
                return true;
            }

            return DataHelper.TryGetDouble(input, out value);
        }
    }
}
