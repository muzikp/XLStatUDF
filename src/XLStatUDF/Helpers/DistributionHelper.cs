/// <summary>
/// Pomocné aproximace pro distribuční testy a odhady parametrů rozdělení.
/// </summary>
namespace XLStatUDF.Helpers
{
    public static class DistributionHelper
    {
        public static double KolmogorovComplementaryCdf(double lambda)
        {
            if (lambda <= 0)
            {
                return 1.0;
            }

            var sum = 0.0;
            for (var j = 1; j < 100; j++)
            {
                var term = Math.Exp(-2.0 * j * j * lambda * lambda);
                sum += (j % 2 == 1 ? 1.0 : -1.0) * term;
                if (term < 1e-12)
                {
                    break;
                }
            }

            return Math.Clamp(2.0 * sum, 0.0, 1.0);
        }

        public static double WeibullShapeEstimate(IReadOnlyList<double> values)
        {
            var logMean = values.Average(Math.Log);
            var shape = 1.2 / StatisticsHelper.StandardDeviationOfLogs(values);
            shape = double.IsFinite(shape) && shape > 0 ? shape : 1.0;

            for (var iteration = 0; iteration < 100; iteration++)
            {
                var sumXk = 0.0;
                var sumXkLogX = 0.0;
                var sumXkLogXSq = 0.0;

                foreach (var value in values)
                {
                    var xk = Math.Pow(value, shape);
                    var logX = Math.Log(value);
                    sumXk += xk;
                    sumXkLogX += xk * logX;
                    sumXkLogXSq += xk * logX * logX;
                }

                var weightedLogMean = sumXkLogX / sumXk;
                var function = (1.0 / shape) + logMean - weightedLogMean;
                var derivative = (-1.0 / (shape * shape)) - ((sumXkLogXSq / sumXk) - (weightedLogMean * weightedLogMean));

                var next = shape - (function / derivative);
                if (!double.IsFinite(next) || next <= 0)
                {
                    next = shape / 2.0;
                }

                if (Math.Abs(next - shape) < 1e-10)
                {
                    return next;
                }

                shape = next;
            }

            return shape;
        }

        public static double WeibullScaleEstimate(IReadOnlyList<double> values, double shape)
            => Math.Pow(values.Average(value => Math.Pow(value, shape)), 1.0 / shape);
    }
}
