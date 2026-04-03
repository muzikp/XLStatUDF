/// <summary>
/// Implementuje populační, výběrový i vážený variační koeficient.
/// </summary>
namespace XLStatUDF.Functions.Descriptive
{
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class VarCoef
    {
        [ExcelFunction(Name = "VARCOEF", Description = "Populační variační koeficient", Category = FunctionCategories.Descriptive)]
        public static object VarCoefP(
            [ExcelArgument(Name = "hodnoty", Description = "Pozorování; prázdné buňky jsou ignorovány")] object values)
        {
            if (!DataHelper.TryReadNumericVector(values, out var parsedValues, out var error))
            {
                return error!;
            }

            if (parsedValues.Length < 1)
            {
                return ExcelErrors.Count;
            }

            var mean = parsedValues.Average();
            if (mean == 0)
            {
                return ExcelErrors.DivZero;
            }

            var variance = parsedValues.Sum(value => Math.Pow(value - mean, 2)) / parsedValues.Length;
            return Math.Sqrt(variance) / mean;
        }

        [ExcelFunction(Name = "VARCOEF.S", Description = "Výběrový variační koeficient", Category = FunctionCategories.Descriptive)]
        public static object VarCoefS(
            [ExcelArgument(Name = "hodnoty", Description = "Pozorování; prázdné buňky jsou ignorovány")] object values)
        {
            if (!DataHelper.TryReadNumericVector(values, out var parsedValues, out var error))
            {
                return error!;
            }

            if (parsedValues.Length < 2)
            {
                return ExcelErrors.Count;
            }

            var mean = parsedValues.Average();
            if (mean == 0)
            {
                return ExcelErrors.DivZero;
            }

            var variance = parsedValues.Sum(value => Math.Pow(value - mean, 2)) / (parsedValues.Length - 1);
            return Math.Sqrt(variance) / mean;
        }

        [ExcelFunction(Name = "VARCOEF.W", Description = "Vážený populační variační koeficient", Category = FunctionCategories.Descriptive)]
        public static object VarCoefW(
            [ExcelArgument(Name = "hodnoty", Description = "Pozorování")] object values,
            [ExcelArgument(Name = "váhy", Description = "Nezáporné váhy")] object weights)
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
            var mean = parsedValues.Zip(parsedWeights, (value, weight) => value * weight).Sum() / sumWeights;
            if (mean == 0)
            {
                return ExcelErrors.DivZero;
            }

            var variance = parsedValues.Zip(parsedWeights, (value, weight) => weight * Math.Pow(value - mean, 2)).Sum() / sumWeights;
            return Math.Sqrt(variance) / mean;
        }

        [ExcelFunction(Name = "VARCOEF.S.W", Description = "Vážený výběrový variační koeficient", Category = FunctionCategories.Descriptive)]
        public static object VarCoefSW(
            [ExcelArgument(Name = "hodnoty", Description = "Pozorování")] object values,
            [ExcelArgument(Name = "váhy", Description = "Nezáporné váhy")] object weights)
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
            if (sumWeights <= 1)
            {
                return ExcelErrors.Count;
            }

            var mean = parsedValues.Zip(parsedWeights, (value, weight) => value * weight).Sum() / sumWeights;
            if (mean == 0)
            {
                return ExcelErrors.DivZero;
            }

            var variance = parsedValues.Zip(parsedWeights, (value, weight) => weight * Math.Pow(value - mean, 2)).Sum() / (sumWeights - 1);
            return Math.Sqrt(variance) / mean;
        }
    }
}
