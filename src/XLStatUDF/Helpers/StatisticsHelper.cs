/// <summary>
/// Poskytuje sdílené výpočty základních statistik pro inferenční funkce.
/// </summary>
namespace XLStatUDF.Helpers
{
    public static class StatisticsHelper
    {
        public static double Mean(IReadOnlyList<double> values) => values.Average();

        public static double SampleVariance(IReadOnlyList<double> values)
        {
            var mean = Mean(values);
            return values.Sum(value => Math.Pow(value - mean, 2)) / (values.Count - 1);
        }

        public static double SampleStandardDeviation(IReadOnlyList<double> values)
            => Math.Sqrt(SampleVariance(values));

        public static double StandardDeviationOfLogs(IReadOnlyList<double> values)
        {
            var logs = values.Select(value => Math.Log(value)).ToArray();
            return SampleStandardDeviation(logs);
        }

        public static double Median(IReadOnlyList<double> values)
        {
            var sorted = values.OrderBy(value => value).ToArray();
            var middle = sorted.Length / 2;
            return sorted.Length % 2 == 0
                ? (sorted[middle - 1] + sorted[middle]) / 2.0
                : sorted[middle];
        }

        public static double PearsonCorrelation(IReadOnlyList<double> x, IReadOnlyList<double> y)
        {
            var meanX = Mean(x);
            var meanY = Mean(y);

            var sumXY = 0.0;
            var sumXX = 0.0;
            var sumYY = 0.0;

            for (var i = 0; i < x.Count; i++)
            {
                var dx = x[i] - meanX;
                var dy = y[i] - meanY;
                sumXY += dx * dy;
                sumXX += dx * dx;
                sumYY += dy * dy;
            }

            return sumXY / Math.Sqrt(sumXX * sumYY);
        }
    }
}
