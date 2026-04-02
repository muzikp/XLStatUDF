/// <summary>
/// Implementuje vážené aritmetické, harmonické a geometrické průměry.
/// </summary>
namespace XLStatUDF.Functions.Descriptive
{
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class WeightedMeans
    {
        [ExcelFunction(Name = "AVERAGE.W", Description = "[Vážené] Vážený aritmetický průměr", Category = "XLStatUDF")]
        public static object AverageW(
            [ExcelArgument(Name = "hodnoty", Description = "Pozorování")] object values,
            [ExcelArgument(Name = "váhy", Description = "Váhy; prázdné buňky jsou brány jako 0")] object weights)
        {
            if (!DataHelper.TryReadPairedNumericWeights(values, weights, out var parsedValues, out var parsedWeights, out var error))
            {
                return error!;
            }

            var validation = WeightHelper.Validate(parsedValues, parsedWeights);
            if (validation is not null)
            {
                return validation;
            }

            var sumWeights = parsedWeights.Sum();
            return WeightedSum(parsedValues, parsedWeights) / sumWeights;
        }

        [ExcelFunction(Name = "HARMEAN.W", Description = "[Vážené] Vážený harmonický průměr", Category = "XLStatUDF")]
        public static object HarMeanW(
            [ExcelArgument(Name = "hodnoty", Description = "Kladná pozorování")] object values,
            [ExcelArgument(Name = "váhy", Description = "Váhy; prázdné buňky jsou brány jako 0")] object weights)
        {
            if (!DataHelper.TryReadPairedNumericWeights(values, weights, out var parsedValues, out var parsedWeights, out var error))
            {
                return error!;
            }

            var validation = WeightHelper.Validate(parsedValues, parsedWeights);
            if (validation is not null)
            {
                return validation;
            }

            if (parsedValues.Zip(parsedWeights).Any(pair => pair.Second > 0 && pair.First <= 0))
            {
                return ExcelErrors.Num;
            }

            var sumWeights = parsedWeights.Sum();
            var denominator = parsedValues.Zip(parsedWeights, (value, weight) => weight == 0 ? 0.0 : weight / value).Sum();
            return sumWeights / denominator;
        }

        [ExcelFunction(Name = "GEOMEAN.W", Description = "[Vážené] Vážený geometrický průměr", Category = "XLStatUDF")]
        public static object GeoMeanW(
            [ExcelArgument(Name = "hodnoty", Description = "Kladná pozorování")] object values,
            [ExcelArgument(Name = "váhy", Description = "Váhy; prázdné buňky jsou brány jako 0")] object weights)
        {
            if (!DataHelper.TryReadPairedNumericWeights(values, weights, out var parsedValues, out var parsedWeights, out var error))
            {
                return error!;
            }

            var validation = WeightHelper.Validate(parsedValues, parsedWeights);
            if (validation is not null)
            {
                return validation;
            }

            if (parsedValues.Zip(parsedWeights).Any(pair => pair.Second > 0 && pair.First <= 0))
            {
                return ExcelErrors.Num;
            }

            var sumWeights = parsedWeights.Sum();
            var logWeightedSum = parsedValues.Zip(parsedWeights, (value, weight) => weight == 0 ? 0.0 : weight * Math.Log(value)).Sum();
            return Math.Exp(logWeightedSum / sumWeights);
        }

        private static double WeightedSum(IReadOnlyList<double> values, IReadOnlyList<double> weights)
        {
            var sum = 0.0;
            for (var i = 0; i < values.Count; i++)
            {
                sum += values[i] * weights[i];
            }

            return sum;
        }
    }
}
