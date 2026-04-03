/// <summary>
/// Implementuje percentily s volitelným filtrováním ve stylu SUMIFS.
/// </summary>
namespace XLStatUDF.Functions.Percentiles
{
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class PercentileIfs
    {
        [ExcelFunction(Name = "PERCENTILE.INC.IFS", Description = "Percentil s filtrováním (INC)", Category = FunctionCategories.Descriptive)]
        public static object PercentileIncIfs(
            [ExcelArgument(Name = "hodnoty", Description = "Zdrojová data")] object values,
            [ExcelArgument(Name = "kvantil", Description = "Požadovaný kvantil z intervalu <0;1>")] double quantile,
            params object[] criteriaArgs)
        {
            if (quantile < 0 || quantile > 1)
            {
                return ExcelErrors.Num;
            }

            if (!FilterHelper.TryApplyFilters(values, criteriaArgs, out var filteredValues, out var error))
            {
                return error!;
            }

            Array.Sort(filteredValues);
            return ComputeInclusivePercentile(filteredValues, quantile);
        }

        [ExcelFunction(Name = "PERCENTILE.EXC.IFS", Description = "Percentil s filtrováním (EXC)", Category = FunctionCategories.Descriptive)]
        public static object PercentileExcIfs(
            [ExcelArgument(Name = "hodnoty", Description = "Zdrojová data")] object values,
            [ExcelArgument(Name = "kvantil", Description = "Požadovaný kvantil z intervalu (0;1)")] double quantile,
            params object[] criteriaArgs)
        {
            if (quantile <= 0 || quantile >= 1)
            {
                return ExcelErrors.Num;
            }

            if (!FilterHelper.TryApplyFilters(values, criteriaArgs, out var filteredValues, out var error))
            {
                return error!;
            }

            Array.Sort(filteredValues);

            var n = filteredValues.Length;
            var rank = quantile * (n + 1);
            if (rank < 1 || rank > n)
            {
                return ExcelErrors.Num;
            }

            return ComputeExclusivePercentile(filteredValues, quantile);
        }

        private static double ComputeInclusivePercentile(IReadOnlyList<double> sortedValues, double quantile)
        {
            if (sortedValues.Count == 1)
            {
                return sortedValues[0];
            }

            var position = quantile * (sortedValues.Count - 1);
            var lowerIndex = (int)Math.Floor(position);
            var upperIndex = (int)Math.Ceiling(position);

            if (lowerIndex == upperIndex)
            {
                return sortedValues[lowerIndex];
            }

            var fraction = position - lowerIndex;
            return sortedValues[lowerIndex] + fraction * (sortedValues[upperIndex] - sortedValues[lowerIndex]);
        }

        private static double ComputeExclusivePercentile(IReadOnlyList<double> sortedValues, double quantile)
        {
            var position = quantile * (sortedValues.Count + 1) - 1;
            var lowerIndex = (int)Math.Floor(position);
            var upperIndex = (int)Math.Ceiling(position);

            if (lowerIndex == upperIndex)
            {
                return sortedValues[lowerIndex];
            }

            var fraction = position - lowerIndex;
            return sortedValues[lowerIndex] + fraction * (sortedValues[upperIndex] - sortedValues[lowerIndex]);
        }
    }
}
