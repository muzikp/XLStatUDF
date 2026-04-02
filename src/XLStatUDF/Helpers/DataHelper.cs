/// <summary>
/// Převádí excelové vstupy na plochá pole a bezpečně parsuje číselná data s podporou prázdných buněk.
/// </summary>
namespace XLStatUDF.Helpers
{
    using System.Globalization;
    using ExcelDna.Integration;

    public static class DataHelper
    {
        private const int HeaderCandidateIndex = 0;

        public static object[] Flatten(object? input)
        {
            return input switch
            {
                null => [null!],
                object[] array => array,
                double[] doubles => doubles.Cast<object>().ToArray(),
                int[] ints => ints.Cast<object>().ToArray(),
                bool[] bools => bools.Cast<object>().ToArray(),
                object[,] matrix => FlattenMatrix(matrix),
                _ => [input]
            };
        }

        public static bool IsBlank(object? value)
        {
            return value is null
                || value is ExcelEmpty
                || value is ExcelMissing
                || value is string text && string.IsNullOrWhiteSpace(text);
        }

        public static bool TryGetDouble(object? value, out double result)
        {
            switch (value)
            {
                case null:
                case ExcelEmpty:
                case ExcelMissing:
                    result = double.NaN;
                    return false;
                case double number:
                    result = number;
                    return true;
                case float number:
                    result = number;
                    return true;
                case decimal number:
                    result = (double)number;
                    return true;
                case byte number:
                    result = number;
                    return true;
                case sbyte number:
                    result = number;
                    return true;
                case short number:
                    result = number;
                    return true;
                case ushort number:
                    result = number;
                    return true;
                case int number:
                    result = number;
                    return true;
                case uint number:
                    result = number;
                    return true;
                case long number:
                    result = number;
                    return true;
                case ulong number:
                    result = number;
                    return true;
                case string text:
                    if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.InvariantCulture, out var invariantParsed))
                    {
                        result = invariantParsed;
                        return true;
                    }

                    if (double.TryParse(text, NumberStyles.Float | NumberStyles.AllowThousands, CultureInfo.CurrentCulture, out var currentParsed))
                    {
                        result = currentParsed;
                        return true;
                    }

                    result = double.NaN;
                    return false;
                default:
                    result = double.NaN;
                    return false;
            }
        }

        public static bool TryReadNumericVector(object? input, out double[] values, out object? error, HeaderMode headerMode = HeaderMode.AutoDetect)
        {
            var items = Flatten(input);
            if (headerMode == HeaderMode.HasHeader && items.Length > 0)
            {
                items = items.Skip(1).ToArray();
            }
            else if (headerMode == HeaderMode.AutoDetect)
            {
                items = SkipLeadingHeader(items, item => TryGetDouble(item, out _));
            }
            var result = new List<double>(items.Length);

            foreach (var item in items)
            {
                if (IsBlank(item))
                {
                    continue;
                }

                if (!TryGetDouble(item, out var parsed))
                {
                    values = [];
                    error = ExcelErrors.Value;
                    return false;
                }

                result.Add(parsed);
            }

            values = result.ToArray();
            error = null;
            return true;
        }

        public static bool TryReadPairedNumericWeights(
            object? valuesInput,
            object? weightsInput,
            out double[] values,
            out double[] weights,
            out object? error,
            HeaderMode headerMode = HeaderMode.AutoDetect)
        {
            var rawValues = Flatten(valuesInput);
            var rawWeights = Flatten(weightsInput);

            if (rawValues.Length != rawWeights.Length)
            {
                values = [];
                weights = [];
                error = ExcelErrors.Length;
                return false;
            }

            if (headerMode == HeaderMode.HasHeader && rawValues.Length > 0 && rawWeights.Length > 0)
            {
                rawValues = rawValues.Skip(1).ToArray();
                rawWeights = rawWeights.Skip(1).ToArray();
            }
            else if (headerMode == HeaderMode.AutoDetect
                && ShouldSkipLeadingHeader(rawValues, item => TryGetDouble(item, out _))
                && ShouldSkipLeadingHeader(rawWeights, item => IsBlank(item) || TryGetDouble(item, out _)))
            {
                rawValues = rawValues.Skip(1).ToArray();
                rawWeights = rawWeights.Skip(1).ToArray();
            }

            var parsedValues = new List<double>(rawValues.Length);
            var parsedWeights = new List<double>(rawWeights.Length);

            for (var i = 0; i < rawValues.Length; i++)
            {
                if (IsBlank(rawValues[i]))
                {
                    continue;
                }

                if (!TryGetDouble(rawValues[i], out var value))
                {
                    values = [];
                    weights = [];
                    error = ExcelErrors.Value;
                    return false;
                }

                double weight = 0.0;
                if (!IsBlank(rawWeights[i]))
                {
                    if (!TryGetDouble(rawWeights[i], out weight))
                    {
                        values = [];
                        weights = [];
                        error = ExcelErrors.Value;
                        return false;
                    }
                }

                parsedValues.Add(value);
                parsedWeights.Add(weight);
            }

            values = parsedValues.ToArray();
            weights = parsedWeights.ToArray();
            error = null;
            return true;
        }

        public static bool TryReadPairedNumericVectors(
            object? firstInput,
            object? secondInput,
            out double[] first,
            out double[] second,
            out object? error,
            HeaderMode headerMode = HeaderMode.AutoDetect)
        {
            var rawFirst = Flatten(firstInput);
            var rawSecond = Flatten(secondInput);

            if (rawFirst.Length != rawSecond.Length)
            {
                first = [];
                second = [];
                error = ExcelErrors.Length;
                return false;
            }

            if (headerMode == HeaderMode.HasHeader && rawFirst.Length > 0 && rawSecond.Length > 0)
            {
                rawFirst = rawFirst.Skip(1).ToArray();
                rawSecond = rawSecond.Skip(1).ToArray();
            }
            else if (headerMode == HeaderMode.AutoDetect
                && ShouldSkipLeadingHeader(rawFirst, item => TryGetDouble(item, out _))
                && ShouldSkipLeadingHeader(rawSecond, item => TryGetDouble(item, out _)))
            {
                rawFirst = rawFirst.Skip(1).ToArray();
                rawSecond = rawSecond.Skip(1).ToArray();
            }

            var parsedFirst = new List<double>(rawFirst.Length);
            var parsedSecond = new List<double>(rawSecond.Length);

            for (var i = 0; i < rawFirst.Length; i++)
            {
                if (IsBlank(rawFirst[i]) || IsBlank(rawSecond[i]))
                {
                    continue;
                }

                if (!TryGetDouble(rawFirst[i], out var firstValue) || !TryGetDouble(rawSecond[i], out var secondValue))
                {
                    first = [];
                    second = [];
                    error = ExcelErrors.Value;
                    return false;
                }

                parsedFirst.Add(firstValue);
                parsedSecond.Add(secondValue);
            }

            first = parsedFirst.ToArray();
            second = parsedSecond.ToArray();
            error = null;
            return true;
        }

        public static bool TryReadBinaryVector(object? input, out int successes, out int count, out object? error, HeaderMode headerMode = HeaderMode.AutoDetect)
        {
            successes = 0;
            count = 0;

            var items = Flatten(input);
            items = headerMode switch
            {
                HeaderMode.HasHeader when items.Length > 0 => items.Skip(1).ToArray(),
                HeaderMode.AutoDetect => SkipLeadingHeader(items, IsBinaryLike),
                _ => items
            };

            foreach (var item in items)
            {
                if (IsBlank(item))
                {
                    continue;
                }

                switch (item)
                {
                    case bool flag:
                        count++;
                        successes += flag ? 1 : 0;
                        break;
                    default:
                        if (!TryGetDouble(item, out var parsed) || (parsed != 0.0 && parsed != 1.0))
                        {
                            error = ExcelErrors.Value;
                            return false;
                        }

                        count++;
                        successes += parsed == 1.0 ? 1 : 0;
                        break;
                }
            }

            error = null;
            return true;
        }

        public static object[] SkipLeadingHeader(object[] items, Func<object?, bool> dataPredicate)
            => ShouldSkipLeadingHeader(items, dataPredicate) ? items.Skip(1).ToArray() : items;

        public static bool ShouldSkipLeadingHeader(object[] items, Func<object?, bool> dataPredicate)
        {
            if (items.Length < 2)
            {
                return false;
            }

            var first = items[HeaderCandidateIndex];
            if (IsBlank(first) || dataPredicate(first))
            {
                return false;
            }

            var hasDataAfterFirst = false;
            for (var i = 1; i < items.Length; i++)
            {
                if (IsBlank(items[i]))
                {
                    continue;
                }

                hasDataAfterFirst = true;
                if (!dataPredicate(items[i]))
                {
                    return false;
                }
            }

            return hasDataAfterFirst;
        }

        private static bool IsBinaryLike(object? value)
        {
            return value switch
            {
                bool => true,
                _ when TryGetDouble(value, out var parsed) => parsed == 0.0 || parsed == 1.0,
                _ => false
            };
        }

        private static object[] FlattenMatrix(object[,] matrix)
        {
            var result = new object[matrix.Length];
            var index = 0;

            for (var row = 0; row < matrix.GetLength(0); row++)
            {
                for (var col = 0; col < matrix.GetLength(1); col++)
                {
                    result[index++] = matrix[row, col];
                }
            }

            return result;
        }
    }
}
