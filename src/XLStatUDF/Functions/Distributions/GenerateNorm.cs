/// <summary>
/// Generuje jednu náhodnou hodnotu z normálního rozdělení s volitelnou perturbací.
/// </summary>
namespace XLStatUDF.Functions.Distributions
{
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using XLStatUDF.Helpers;

    public static class GenerateNorm
    {
        [ExcelFunction(
            Name = "GENERATE.NORM",
            Description = "Vygeneruje jednu náhodnou hodnotu z normálního rozdělení s volitelným šumem",
            Category = FunctionCategories.General,
            IsVolatile = true)]
        public static object GenerateNormalSample(
            [ExcelArgument(Name = "stredni_hodnota", Description = "Střední hodnota normálního rozdělení")] object mean,
            [ExcelArgument(Name = "směrodatná_odchylka", Description = "Kladná směrodatná odchylka")] object standardDeviation,
            [ExcelArgument(Name = "outlier_rate", Description = "Volitelně: pravděpodobnost dodatečné náhodné perturbace v intervalu <0;1>; výchozí 0")] object? outlierRate = null)
        {
            if (!DataHelper.TryGetDouble(mean, out var mu)
                || !DataHelper.TryGetDouble(standardDeviation, out var sigma)
                || !TryGetOptionalProbability(outlierRate, 0.0, out var rate))
            {
                return ExcelErrors.Value;
            }

            if (sigma <= 0.0)
            {
                return ExcelErrors.Num;
            }

            var value = Normal.Sample(Random.Shared, mu, sigma);
            if (rate > 0.0 && Random.Shared.NextDouble() < rate)
            {
                value += Normal.Sample(Random.Shared, 0.0, sigma * 3.0);
            }

            return value;
        }

        private static bool TryGetOptionalProbability(object? input, double defaultValue, out double result)
        {
            if (input is null or ExcelMissing or ExcelEmpty)
            {
                result = defaultValue;
                return true;
            }

            if (!DataHelper.TryGetDouble(input, out result))
            {
                return false;
            }

            return result >= 0.0 && result <= 1.0;
        }
    }
}
