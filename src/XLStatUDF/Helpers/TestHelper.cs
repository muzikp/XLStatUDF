/// <summary>
/// Validuje společné argumenty statistických testů a počítá směrové p-hodnoty a kritické hodnoty.
/// </summary>
namespace XLStatUDF.Helpers
{
    using MathNet.Numerics.Distributions;

    public static class TestHelper
    {
        public static bool TryParseDirection(object? input, out string direction)
        {
            switch (input)
            {
                case null:
                case ExcelDna.Integration.ExcelMissing:
                case ExcelDna.Integration.ExcelEmpty:
                    direction = "two";
                    return true;
                case double number:
                    return TryParseDirectionCode(number, out direction);
                case int number:
                    return TryParseDirectionCode(number, out direction);
                default:
                    direction = string.Empty;
                    return false;
            }
        }

        public static bool IsValidDirection(string direction)
            => direction is "two" or "left" or "right";

        public static bool IsValidAlpha(double alpha)
            => alpha > 0 && alpha < 1;

        public static double CriticalT(double alpha, double df, string direction)
            => Math.Abs(StudentT.InvCDF(0, 1, df, direction == "two" ? 1 - (alpha / 2.0) : 1 - alpha));

        public static double CriticalZ(double alpha, string direction)
            => Math.Abs(Normal.InvCDF(0, 1, direction == "two" ? 1 - (alpha / 2.0) : 1 - alpha));

        public static double PValueFromT(double statistic, double df, string direction)
        {
            var cdf = StudentT.CDF(0, 1, df, statistic);
            return direction switch
            {
                "left" => cdf,
                "right" => 1 - cdf,
                _ => 2 * (1 - StudentT.CDF(0, 1, df, Math.Abs(statistic)))
            };
        }

        public static double PValueFromZ(double statistic, string direction)
        {
            var cdf = Normal.CDF(0, 1, statistic);
            return direction switch
            {
                "left" => cdf,
                "right" => 1 - cdf,
                _ => 2 * (1 - Normal.CDF(0, 1, Math.Abs(statistic)))
            };
        }

        private static bool TryParseDirectionCode(double code, out string direction)
        {
            direction = code switch
            {
                0 => "two",
                1 => "left",
                2 => "right",
                _ => string.Empty
            };

            return direction.Length > 0;
        }
    }
}
