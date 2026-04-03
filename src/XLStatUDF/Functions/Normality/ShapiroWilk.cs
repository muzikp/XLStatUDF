/// <summary>
/// Implementuje Shapiro-Wilkův test normality s Roystonovou aproximací p-hodnoty.
/// </summary>
namespace XLStatUDF.Functions.Normality
{
    using ExcelDna.Integration;
    using MathNet.Numerics.Distributions;
    using XLStatUDF.Helpers;

    public static class ShapiroWilk
    {
        [ExcelFunction(Name = "SHAPIRO.WILK", Description = "Shapiro-Wilkův test", Category = FunctionCategories.Tests)]
        public static object[,] Run(
            [ExcelArgument(Name = "rozsah_hodnot", Description = "Číselná data; prázdné buňky jsou ignorovány")] object values,
            [ExcelArgument(Name = "ma_záhlaví", Description = "Volitelně: 0=autodetect, 1=má záhlaví, 2=nemá záhlaví")] object? hasHeader = null)
        {
            if (!ArgumentHelper.TryParseHeaderMode(hasHeader, out var parsedHeaderMode))
            {
                return SpillError(ExcelErrors.Value);
            }

            if (!DataHelper.TryReadNumericVector(values, out var sample, out var error, parsedHeaderMode))
            {
                return SpillError(error!);
            }

            if (sample.Length < 3 || sample.Length > 5000)
            {
                return SpillError(ExcelErrors.Count);
            }

            Array.Sort(sample);
            var w = ComputeStatistic(sample);
            var p = ComputePValue(w, sample.Length);

            var rows = new List<object[]>();
            SpillBuilder.AddRow(rows, "W", w);
            SpillBuilder.AddRow(rows, "p", p);
            return SpillBuilder.Build(rows);
        }

        internal static double ComputeStatistic(IReadOnlyList<double> sortedSample)
        {
            var n = sortedSample.Count;
            if (n == 3)
            {
                var range = sortedSample[2] - sortedSample[0];
                var middleDelta = sortedSample[1] - sortedSample[0];
                return range == 0 ? 1.0 : 0.75 * Math.Pow(range - 2.0 * middleDelta, 2) / Math.Pow(range, 2);
            }

            var half = n / 2;
            var m = new double[n];
            for (var i = 0; i < n; i++)
            {
                var p = ((i + 1) - 0.375) / (n + 0.25);
                m[i] = Normal.InvCDF(0, 1, p);
            }

            var mSumSquares = m.Sum(value => value * value);
            var u = 1.0 / Math.Sqrt(n);
            var weights = new double[half];
            weights[0] = EvaluatePolynomial([0.221157, -0.147981, -2.07119, 4.434685, -2.706056], u);
            if (half > 1)
            {
                weights[1] = EvaluatePolynomial([0.042981, -0.293762, -1.752461, 5.682633, -3.582633], u);
            }

            var upperExpected = Enumerable.Range(0, half).Select(index => m[n - 1 - index]).ToArray();
            var remainderDenominator = 1.0 - (2.0 * weights[0] * weights[0]) - (half > 1 ? 2.0 * weights[1] * weights[1] : 0.0);
            var remainderNumerator = mSumSquares - (2.0 * upperExpected[0] * upperExpected[0]) - (half > 1 ? 2.0 * upperExpected[1] * upperExpected[1] : 0.0);
            var scale = Math.Sqrt(remainderNumerator / remainderDenominator);

            for (var i = 2; i < half; i++)
            {
                weights[i] = upperExpected[i] / scale;
            }

            var numerator = 0.0;
            for (var i = 0; i < half; i++)
            {
                numerator += weights[i] * (sortedSample[n - 1 - i] - sortedSample[i]);
            }

            var mean = StatisticsHelper.Mean(sortedSample);
            var denominator = sortedSample.Sum(value => Math.Pow(value - mean, 2));
            if (denominator == 0)
            {
                return 1.0;
            }

            return Math.Clamp((numerator * numerator) / denominator, 0.0, 1.0);
        }

        internal static double ComputePValue(double w, int n)
        {
            w = Math.Clamp(w, 1e-12, 1 - 1e-12);
            if (n == 3)
            {
                var exact = 1.0 - ((6.0 / Math.PI) * Math.Acos(Math.Sqrt(w)));
                return Math.Clamp(exact, 0.0, 1.0);
            }

            var y = Math.Log(1.0 - w);
            double mu;
            double sigma;

            if (n <= 11)
            {
                var gamma = -2.273 + (0.459 * n);
                if (y >= gamma)
                {
                    return 1e-12;
                }

                y = -Math.Log(gamma - y);
                mu = EvaluatePolynomial([0.5440, -0.39978, 0.025054, -0.0006714], n);
                sigma = Math.Exp(EvaluatePolynomial([1.3822, -0.77857, 0.062767, -0.0020322], n));
            }
            else
            {
                var logN = Math.Log(n);
                mu = EvaluatePolynomial([-1.5861, -0.31082, -0.083751, 0.0038915], logN);
                sigma = Math.Exp(EvaluatePolynomial([-0.4803, -0.082676, 0.0030302], logN));
            }

            var z = (y - mu) / sigma;
            return Math.Clamp(1.0 - Normal.CDF(0, 1, z), 0.0, 1.0);
        }

        private static double EvaluatePolynomial(IReadOnlyList<double> coefficients, double x)
        {
            var value = 0.0;
            var power = 1.0;
            foreach (var coefficient in coefficients)
            {
                value += coefficient * power;
                power *= x;
            }

            return value;
        }

        private static object[,] SpillError(object error)
            => new object[,] { { error, string.Empty } };
    }
}
