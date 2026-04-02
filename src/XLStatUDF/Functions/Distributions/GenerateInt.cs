/// <summary>
/// Generuje spill sloupec náhodných celých čísel ze zadaného uzavřeného intervalu.
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
            Description = "[Rozdělení] Vygeneruje náhodná celá čísla do spill sloupce",
            Category = "XLStatUDF",
            IsVolatile = true)]
        public static object GenerateIntegerSample(
            [ExcelArgument(Name = "pocet", Description = "Volitelne: pocet generovanych hodnot; cele cislo >= 1; vychozi 1")] object? count = null,
            [ExcelArgument(Name = "minimum", Description = "Volitelne: dolni mez intervalu; vychozi prakticky strop int.MinValue")] object? minimum = null,
            [ExcelArgument(Name = "maximum", Description = "Volitelne: horni mez intervalu; vychozi prakticky strop int.MaxValue")] object? maximum = null)
        {
            if (!TryGetOptionalInteger(count, 1, out var sampleCount))
            {
                return ExcelErrors.Value;
            }

            if (sampleCount < 1)
            {
                return ExcelErrors.Num;
            }

            if (!TryGetOptionalInteger(minimum, DefaultMinimum, out var minValue)
                || !TryGetOptionalInteger(maximum, DefaultMaximum, out var maxValue))
            {
                return ExcelErrors.Value;
            }

            if (minValue > maxValue)
            {
                return ExcelErrors.Num;
            }

            var result = new object[sampleCount, 1];
            for (var i = 0; i < sampleCount; i++)
            {
                result[i, 0] = Random.Shared.NextInt64(minValue, (long)maxValue + 1);
            }

            return result;
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
    }
}
