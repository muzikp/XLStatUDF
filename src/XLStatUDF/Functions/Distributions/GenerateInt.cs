/// <summary>
/// Generuje jedno náhodné celé číslo s volitelnou perturbací.
/// </summary>
namespace XLStatUDF.Functions.Distributions
{
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class GenerateInt
    {
        private const int DefaultMinimum = int.MinValue;
        private const int DefaultMaximum = int.MaxValue;

        [ExcelFunction(
            Name = "GENERATE.INT",
            Description = "Vygeneruje jedno náhodné celé číslo s volitelným šumem",
            Category = FunctionCategories.General,
            IsVolatile = true)]
        public static object GenerateIntegerSample(
            [ExcelArgument(Name = "minimum", Description = "Volitelně: dolní mez intervalu; výchozí int.MinValue")] object? minimum = null,
            [ExcelArgument(Name = "maximum", Description = "Volitelně: horní mez intervalu; výchozí int.MaxValue")] object? maximum = null,
            [ExcelArgument(Name = "outlier_rate", Description = "Volitelně: pravděpodobnost dodatečné náhodné perturbace v intervalu <0;1>; výchozí 0")] object? outlierRate = null)
        {
            if (!TryGetOptionalInteger(minimum, DefaultMinimum, out var minValue)
                || !TryGetOptionalInteger(maximum, DefaultMaximum, out var maxValue)
                || !TryGetOptionalProbability(outlierRate, 0.0, out var rate))
            {
                return ExcelErrors.Value;
            }

            if (minValue > maxValue)
            {
                return ExcelErrors.Num;
            }

            var value = Random.Shared.NextInt64(minValue, (long)maxValue + 1);
            if (rate > 0.0 && Random.Shared.NextDouble() < rate)
            {
                var width = Math.Max(1L, (long)maxValue - minValue);
                var offset = Random.Shared.NextInt64(-width, width + 1);
                value = Math.Clamp(value + offset, int.MinValue, int.MaxValue);
            }

            return value;
        }

        private static bool TryGetOptionalInteger(object? input, int defaultValue, out int result)
        {
            if (input is null or ExcelMissing or ExcelEmpty)
            {
                result = defaultValue;
                return true;
            }

            if (!DataHelper.TryGetDouble(input, out var parsed))
            {
                result = 0;
                return false;
            }

            if (Math.Abs(parsed - Math.Round(parsed)) > 1e-9
                || parsed < int.MinValue
                || parsed > int.MaxValue)
            {
                result = 0;
                return false;
            }

            result = (int)Math.Round(parsed);
            return true;
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
