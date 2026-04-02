/// <summary>
/// Implementuje vážené rozptyly a směrodatné odchylky pro populační i výběrovou variantu.
/// </summary>
namespace XLStatUDF.Functions.Descriptive
{
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class WeightedVariance
    {
        [ExcelFunction(Name = "VAR.P.W", Description = "[Vážené] Vážený populační rozptyl", Category = "XLStatUDF")]
        public static object VarPW(
            [ExcelArgument(Name = "hodnoty", Description = "Pozorování")] object values,
            [ExcelArgument(Name = "váhy", Description = "Nezáporné váhy")] object weights)
            => TryComputeWeightedVariance(values, weights, sample: false, out var result) ? result : result;

        [ExcelFunction(Name = "VAR.S.W", Description = "[Vážené] Vážený výběrový rozptyl", Category = "XLStatUDF")]
        public static object VarSW(
            [ExcelArgument(Name = "hodnoty", Description = "Pozorování")] object values,
            [ExcelArgument(Name = "váhy", Description = "Nezáporné váhy")] object weights)
            => TryComputeWeightedVariance(values, weights, sample: true, out var result) ? result : result;

        [ExcelFunction(Name = "STDEV.P.W", Description = "[Vážené] Vážená populační směrodatná odchylka", Category = "XLStatUDF")]
        public static object StdevPW(
            [ExcelArgument(Name = "hodnoty", Description = "Pozorování")] object values,
            [ExcelArgument(Name = "váhy", Description = "Nezáporné váhy")] object weights)
        {
            if (!TryComputeWeightedVariance(values, weights, sample: false, out var variance))
            {
                return variance;
            }

            return Math.Sqrt((double)variance);
        }

        [ExcelFunction(Name = "STDEV.S.W", Description = "[Vážené] Vážená výběrová směrodatná odchylka", Category = "XLStatUDF")]
        public static object StdevSW(
            [ExcelArgument(Name = "hodnoty", Description = "Pozorování")] object values,
            [ExcelArgument(Name = "váhy", Description = "Nezáporné váhy")] object weights)
        {
            if (!TryComputeWeightedVariance(values, weights, sample: true, out var variance))
            {
                return variance;
            }

            return Math.Sqrt((double)variance);
        }

        internal static bool TryComputeWeightedVariance(object valuesInput, object weightsInput, bool sample, out object result)
        {
            if (!DataHelper.TryReadPairedNumericWeights(valuesInput, weightsInput, out var values, out var weights, out var error))
            {
                result = error!;
                return false;
            }

            var validation = WeightHelper.Validate(values, weights);
            if (validation is not null)
            {
                result = validation;
                return false;
            }

            var sumWeights = weights.Sum();
            if (sample && sumWeights <= 1)
            {
                result = ExcelErrors.Count;
                return false;
            }

            var mean = values.Zip(weights, (value, weight) => value * weight).Sum() / sumWeights;
            var weightedSumSquares = values.Zip(weights, (value, weight) => weight * Math.Pow(value - mean, 2)).Sum();
            result = weightedSumSquares / (sample ? sumWeights - 1 : sumWeights);
            return true;
        }
    }
}
