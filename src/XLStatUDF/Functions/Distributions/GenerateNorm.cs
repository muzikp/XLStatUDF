/// <summary>
/// Generuje spill sloupec náhodných hodnot z normálního rozdělení se zadaným průměrem a směrodatnou odchylkou.
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
            Description = "[Rozdělení] Vygeneruje náhodná čísla z normálního rozdělení do spill sloupce",
            Category = "XLStatUDF",
            IsVolatile = true)]
        public static object GenerateNormalSample(
            [ExcelArgument(Name = "x", Description = "Požadovaný průměr normálního rozdělení")] object mean,
            [ExcelArgument(Name = "stdev", Description = "Směrodatná odchylka; musí být kladná")] object standardDeviation,
            [ExcelArgument(Name = "count", Description = "Počet generovaných hodnot; celé číslo >= 1")] object count)
        {
            if (!DataHelper.TryGetDouble(mean, out var mu)
                || !DataHelper.TryGetDouble(standardDeviation, out var sigma)
                || !DataHelper.TryGetDouble(count, out var sampleCountDouble))
            {
                return ExcelErrors.Value;
            }

            if (sigma <= 0)
            {
                return ExcelErrors.Num;
            }

            if (Math.Abs(sampleCountDouble - Math.Round(sampleCountDouble)) > 1e-9)
            {
                return ExcelErrors.Num;
            }

            var sampleCount = (int)Math.Round(sampleCountDouble);
            if (sampleCount < 1)
            {
                return ExcelErrors.Num;
            }

            var result = new object[sampleCount, 1];
            for (var i = 0; i < sampleCount; i++)
            {
                result[i, 0] = Normal.Sample(Random.Shared, mu, sigma);
            }

            return result;
        }
    }
}
