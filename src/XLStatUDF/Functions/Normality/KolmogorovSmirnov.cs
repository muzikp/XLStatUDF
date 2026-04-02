/// <summary>
/// Implementuje jednovýběrový Kolmogorov-Smirnovův test pro několik běžných rozdělení.
/// </summary>
namespace XLStatUDF.Functions.Normality
{
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using XLStatUDF.Helpers;

    public static class KolmogorovSmirnov
    {
        [ExcelFunction(Name = "KOLMOGOROV.SMIRNOV", Description = "[Normalita] Kolmogorov-Smirnovův test dobré shody", Category = "XLStatUDF")]
        public static object[,] Run(
            [ExcelArgument(Name = "rozsah_hodnot", Description = "Číselná data; prázdné buňky jsou ignorovány")] object values,
            [ExcelArgument(Name = "typ_rozdeleni", Description = "Volitelne: 0=normal (normalni), 1=lognormal, 2=exponential, 3=uniform, 4=weibull")] object? distribution = null,
            [ExcelArgument(Name = "ma_zahlavi", Description = "Volitelne: 0=autodetect, 1=ma zahlavi, 2=nema zahlavi")] object? hasHeader = null)
        {
            if (!ArgumentHelper.TryParseHeaderMode(hasHeader, out var parsedHeaderMode))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!DataHelper.TryReadNumericVector(values, out var sample, out var error, parsedHeaderMode))
            {
                return SpillError(error!);
            }

            if (sample.Length < 5)
            {
                return SpillError(ExcelErrors.Count);
            }

            if (!TryParseDistribution(distribution, out var parsedDistribution))
            {
                return SpillError(ExcelErrors.Value);
            }

            Array.Sort(sample);

            if (!TryBuildCdf(sample, parsedDistribution, out var cdf, out error))
            {
                return SpillError(error!);
            }

            var d = ComputeStatistic(sample, cdf);
            var p = ComputePValue(d, sample.Length, parsedDistribution);

            var rows = new List<object[]>();
            SpillBuilder.AddRow(rows, "D", d);
            SpillBuilder.AddRow(rows, "p", p);
            return SpillBuilder.Build(rows);
        }

        internal static double ComputeStatistic(IReadOnlyList<double> sortedSample, Func<double, double> cdf)
        {
            var n = sortedSample.Count;
            var d = 0.0;

            for (var i = 0; i < n; i++)
            {
                var f = Math.Clamp(cdf(sortedSample[i]), 0.0, 1.0);
                var empiricalUpper = (i + 1) / (double)n;
                var empiricalLower = i / (double)n;
                d = Math.Max(d, Math.Abs(empiricalUpper - f));
                d = Math.Max(d, Math.Abs(f - empiricalLower));
            }

            return d;
        }

        internal static double ComputePValue(double d, int n, string distribution)
        {
            double lambda;

            if (distribution == "normal")
            {
                lambda = (Math.Sqrt(n) - 0.01 + (0.85 / Math.Sqrt(n))) * d;
            }
            else
            {
                lambda = (Math.Sqrt(n) + 0.12 + (0.11 / Math.Sqrt(n))) * d;
            }

            return DistributionHelper.KolmogorovComplementaryCdf(lambda);
        }

        private static bool TryParseDistribution(object? input, out string distribution)
        {
            switch (input)
            {
                case null:
                case ExcelMissing:
                case ExcelEmpty:
                    distribution = "normal";
                    return true;
                case double number:
                    return TryParseDistributionCode((int)number, out distribution) && Math.Abs(number - Math.Round(number)) < 1e-9;
                case int number:
                    return TryParseDistributionCode(number, out distribution);
                default:
                    distribution = string.Empty;
                    return false;
            }
        }

        private static bool TryParseDistributionCode(int code, out string distribution)
        {
            distribution = code switch
            {
                0 => "normal",
                1 => "lognormal",
                2 => "exponential",
                3 => "uniform",
                4 => "weibull",
                _ => string.Empty
            };

            return distribution.Length > 0;
        }

        private static bool TryBuildCdf(IReadOnlyList<double> sample, string distribution, out Func<double, double> cdf, out object? error)
        {
            switch (distribution)
            {
                case "normal":
                {
                    var mean = StatisticsHelper.Mean(sample);
                    var sigma = StatisticsHelper.SampleStandardDeviation(sample);
                    if (sigma <= 0)
                    {
                        cdf = _ => 0.0;
                        error = ExcelErrors.Num;
                        return false;
                    }

                    cdf = x => Normal.CDF(mean, sigma, x);
                    error = null;
                    return true;
                }
                case "lognormal":
                {
                    if (sample.Any(value => value <= 0))
                    {
                        cdf = _ => 0.0;
                        error = ExcelErrors.Num;
                        return false;
                    }

                    var logs = sample.Select(value => Math.Log(value)).ToArray();
                    var mean = StatisticsHelper.Mean(logs);
                    var sigma = StatisticsHelper.SampleStandardDeviation(logs);
                    cdf = x => x <= 0 ? 0.0 : LogNormal.CDF(mean, sigma, x);
                    error = null;
                    return true;
                }
                case "exponential":
                {
                    if (sample.Any(value => value < 0))
                    {
                        cdf = _ => 0.0;
                        error = ExcelErrors.Num;
                        return false;
                    }

                    var mean = StatisticsHelper.Mean(sample);
                    if (mean <= 0)
                    {
                        cdf = _ => 0.0;
                        error = ExcelErrors.Num;
                        return false;
                    }

                    var lambda = 1.0 / mean;
                    cdf = x => x < 0 ? 0.0 : 1.0 - Math.Exp(-lambda * x);
                    error = null;
                    return true;
                }
                case "uniform":
                {
                    var min = sample.Min();
                    var max = sample.Max();
                    if (min == max)
                    {
                        cdf = _ => 0.0;
                        error = ExcelErrors.Num;
                        return false;
                    }

                    cdf = x => x <= min ? 0.0 : x >= max ? 1.0 : (x - min) / (max - min);
                    error = null;
                    return true;
                }
                case "weibull":
                {
                    if (sample.Any(value => value <= 0))
                    {
                        cdf = _ => 0.0;
                        error = ExcelErrors.Num;
                        return false;
                    }

                    var shape = DistributionHelper.WeibullShapeEstimate(sample);
                    var scale = DistributionHelper.WeibullScaleEstimate(sample, shape);
                    cdf = x => x <= 0 ? 0.0 : 1.0 - Math.Exp(-Math.Pow(x / scale, shape));
                    error = null;
                    return true;
                }
                default:
                    cdf = _ => 0.0;
                    error = ExcelErrors.Value;
                    return false;
            }
        }

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
