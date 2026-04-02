/// <summary>
/// Pomáhá s načtením kategorií a číselných hodnot do skupin pro testy a ANOVA.
/// </summary>
namespace XLStatUDF.Helpers
{
    using System.Globalization;

    public static class GroupingHelper
    {
        public static bool TryReadGroups(
            object categoriesInput,
            object valuesInput,
            out Dictionary<string, List<double>> groups,
            out object? error,
            HeaderMode headerMode = HeaderMode.AutoDetect)
        {
            var rawCategories = DataHelper.Flatten(categoriesInput);
            var rawValues = DataHelper.Flatten(valuesInput);

            if (rawCategories.Length != rawValues.Length)
            {
                groups = new Dictionary<string, List<double>>();
                error = ExcelErrors.Length;
                return false;
            }

            if (headerMode == HeaderMode.HasHeader && rawCategories.Length > 0 && rawValues.Length > 0)
            {
                rawCategories = rawCategories.Skip(1).ToArray();
                rawValues = rawValues.Skip(1).ToArray();
            }
            else if (headerMode == HeaderMode.AutoDetect
                && !DataHelper.IsBlank(rawCategories[0])
                && !DataHelper.TryGetDouble(rawValues[0], out _)
                && DataHelper.ShouldSkipLeadingHeader(rawValues, item => DataHelper.TryGetDouble(item, out _)))
            {
                rawCategories = rawCategories.Skip(1).ToArray();
                rawValues = rawValues.Skip(1).ToArray();
            }

            groups = new Dictionary<string, List<double>>(StringComparer.Ordinal);

            for (var i = 0; i < rawCategories.Length; i++)
            {
                if (DataHelper.IsBlank(rawCategories[i]) || DataHelper.IsBlank(rawValues[i]))
                {
                    continue;
                }

                if (!DataHelper.TryGetDouble(rawValues[i], out var value))
                {
                    error = ExcelErrors.Value;
                    return false;
                }

                var label = Convert.ToString(rawCategories[i], CultureInfo.InvariantCulture);
                if (string.IsNullOrWhiteSpace(label))
                {
                    continue;
                }

                if (!groups.TryGetValue(label, out var list))
                {
                    list = new List<double>();
                    groups[label] = list;
                }

                list.Add(value);
            }

            error = null;
            return true;
        }
    }
}
