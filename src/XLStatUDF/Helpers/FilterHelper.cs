/// <summary>
/// Implementuje základní SUMIFS-style filtrování včetně relačních operátorů a wildcard pravidel.
/// </summary>
namespace XLStatUDF.Helpers
{
    using System.Globalization;
    using System.Text.RegularExpressions;

    public static class FilterHelper
    {
        public static bool TryApplyFilters(object? valuesInput, object[] criteriaArgs, out double[] values, out object? error)
        {
            var rawValues = DataHelper.Flatten(valuesInput);

            if (criteriaArgs.Length % 2 != 0)
            {
                values = [];
                error = ExcelErrors.Value;
                return false;
            }

            var include = Enumerable.Repeat(true, rawValues.Length).ToArray();

            for (var i = 0; i < criteriaArgs.Length; i += 2)
            {
                var criteriaRange = DataHelper.Flatten(criteriaArgs[i]);
                if (criteriaRange.Length != rawValues.Length)
                {
                    values = [];
                    error = ExcelErrors.Length;
                    return false;
                }

                var criterion = criteriaArgs[i + 1];
                for (var j = 0; j < criteriaRange.Length; j++)
                {
                    include[j] &= Matches(criteriaRange[j], criterion);
                }
            }

            var filtered = new List<double>();

            for (var i = 0; i < rawValues.Length; i++)
            {
                if (!include[i] || DataHelper.IsBlank(rawValues[i]))
                {
                    continue;
                }

                if (!DataHelper.TryGetDouble(rawValues[i], out var parsed))
                {
                    values = [];
                    error = ExcelErrors.Value;
                    return false;
                }

                filtered.Add(parsed);
            }

            values = filtered.ToArray();
            error = filtered.Count == 0 ? ExcelErrors.NA : null;
            return filtered.Count > 0;
        }

        private static bool Matches(object? actual, object? criterion)
        {
            if (criterion is null)
            {
                return actual is null;
            }

            if (criterion is string textCriterion)
            {
                var (op, operand) = ParseCriterion(textCriterion);
                return Compare(actual, op, operand);
            }

            return Compare(actual, "=", criterion);
        }

        private static (string Operator, object Operand) ParseCriterion(string criterion)
        {
            var operators = new[] { ">=", "<=", "<>", ">", "<", "=" };

            foreach (var op in operators)
            {
                if (criterion.StartsWith(op, StringComparison.Ordinal))
                {
                    return (op, criterion[op.Length..]);
                }
            }

            return ("=", criterion);
        }

        private static bool Compare(object? actual, string comparisonOperator, object operand)
        {
            if (TryAsNumeric(actual, out var actualNumber) && TryAsNumeric(operand, out var operandNumber))
            {
                return comparisonOperator switch
                {
                    "=" => actualNumber.Equals(operandNumber),
                    "<>" => !actualNumber.Equals(operandNumber),
                    ">" => actualNumber > operandNumber,
                    "<" => actualNumber < operandNumber,
                    ">=" => actualNumber >= operandNumber,
                    "<=" => actualNumber <= operandNumber,
                    _ => false
                };
            }

            var actualText = DataHelper.IsBlank(actual) ? string.Empty : Convert.ToString(actual, CultureInfo.InvariantCulture) ?? string.Empty;
            var operandText = Convert.ToString(operand, CultureInfo.InvariantCulture) ?? string.Empty;

            if (operandText.Contains('*') || operandText.Contains('?'))
            {
                var regex = "^" + Regex.Escape(operandText).Replace("\\*", ".*").Replace("\\?", ".") + "$";
                var isMatch = Regex.IsMatch(actualText, regex, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
                return comparisonOperator == "<>" ? !isMatch : isMatch;
            }

            var comparison = string.Compare(actualText, operandText, StringComparison.OrdinalIgnoreCase);
            return comparisonOperator switch
            {
                "=" => comparison == 0,
                "<>" => comparison != 0,
                ">" => comparison > 0,
                "<" => comparison < 0,
                ">=" => comparison >= 0,
                "<=" => comparison <= 0,
                _ => false
            };
        }

        private static bool TryAsNumeric(object? value, out double number)
        {
            if (value is string text)
            {
                return double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out number)
                    || double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out number);
            }

            return DataHelper.TryGetDouble(value, out number);
        }
    }
}
