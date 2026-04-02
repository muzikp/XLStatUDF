/// <summary>
/// Validuje a normalizuje váhy pro vážené statistiky.
/// </summary>
namespace XLStatUDF.Helpers
{
    public static class WeightHelper
    {
        public static object? Validate(double[] values, double[] weights)
        {
            if (values.Length != weights.Length)
            {
                return ExcelErrors.Length;
            }

            if (weights.Any(weight => weight < 0))
            {
                return ExcelErrors.Num;
            }

            if (weights.Sum() <= 0)
            {
                return ExcelErrors.Num;
            }

            return null;
        }

        public static double[] Normalize(double[] weights)
        {
            var normalizedWeights = weights.Select(weight => double.IsNaN(weight) ? 0.0 : weight).ToArray();
            var sum = normalizedWeights.Sum();
            return normalizedWeights.Select(weight => weight / sum).ToArray();
        }
    }
}
