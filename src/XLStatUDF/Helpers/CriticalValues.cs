/// <summary>
/// Vytváří konzistentní české názvy řádků s kritickými hodnotami rozdělení.
/// </summary>
namespace XLStatUDF.Helpers
{
    public static class CriticalValues
    {
        private const string SubOne = "\u2081";
        private const string SubTwo = "\u2082";
        private const string SubMinus = "\u208B";
        private const string Alpha = "\u03B1";
        private const string FractionSlash = "\u2044";

        public static string LabelForDirection(string distribution, string direction)
        {
            var symbol = distribution switch
            {
                "t" => "t",
                "z" => "z",
                "F" => "F",
                "chi2" => "χ²",
                _ => distribution
            };

            return direction == "two"
                ? $"{symbol}{SubOne}{SubMinus}{Alpha}{FractionSlash}{SubTwo}"
                : $"{symbol}{SubOne}{SubMinus}{Alpha}";
        }
    }
}
