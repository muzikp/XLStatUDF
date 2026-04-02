/// <summary>
/// Opakuje jednu hodnotu nebo opakovaně vyhodnocuje zadaný vzorec do spill sloupce.
/// </summary>
namespace XLStatUDF.Functions.Distributions
{
    using ExcelDna.Integration;
    using XLStatUDF.Helpers;

    public static class Fill
    {
        private delegate bool ValueFactory(out object value, out object? error);

        [ExcelFunction(
            Name = "FILL",
            Description = "[Rozdělení] Zkopíruje jednu nebo více hodnot nebo vzorců do spill sloupce",
            Category = "XLStatUDF",
            IsVolatile = true)]
        public static object FillDown(
            [ExcelArgument(Name = "co", Description = "Hodnota nebo textovy vzorec zacinatejici =")] object what,
            [ExcelArgument(Name = "pocet", Description = "Pocet opakovani; cele cislo >= 1")] object count,
            [ExcelArgument(Name = "dalsi_pary", Description = "Volitelne dalsi dvojice co+pocet")] params object[] morePairs)
            => BuildFillResult(what, count, morePairs, shuffle: false);

        [ExcelFunction(
            Name = "FILL.RANDOM",
            Description = "[Rozdělení] Zkopíruje jednu nebo vice hodnot nebo vzorcu do spill sloupce a nahodne je promicha",
            Category = "XLStatUDF",
            IsVolatile = true)]
        public static object FillDownRandom(
            [ExcelArgument(Name = "co", Description = "Hodnota nebo textovy vzorec zacinatejici =")] object what,
            [ExcelArgument(Name = "pocet", Description = "Pocet opakovani; cele cislo >= 1")] object count,
            [ExcelArgument(Name = "dalsi_pary", Description = "Volitelne dalsi dvojice co+pocet")] params object[] morePairs)
            => BuildFillResult(what, count, morePairs, shuffle: true);

        private static object BuildFillResult(
            object what,
            object count,
            object[] morePairs,
            bool shuffle)
        {
            if (morePairs.Length % 2 != 0)
            {
                return ExcelErrors.Value;
            }

            var pairs = new List<(object What, object Count)> { (what, count) };
            for (var i = 0; i < morePairs.Length; i += 2)
            {
                pairs.Add((morePairs[i], morePairs[i + 1]));
            }

            var segments = new List<(ValueFactory Factory, int Count)>(pairs.Count);
            var totalCount = 0;

            foreach (var pair in pairs)
            {
                if (!TryParseCount(pair.Count, out var repeatCount, out var error))
                {
                    return error!;
                }

                if (!TryCreateValueFactory(pair.What, out var valueFactory, out error))
                {
                    return error!;
                }

                segments.Add((valueFactory, repeatCount));
                totalCount += repeatCount;
            }

            var values = new List<object>(totalCount);
            foreach (var segment in segments)
            {
                for (var i = 0; i < segment.Count; i++)
                {
                    if (!segment.Factory(out var value, out var error))
                    {
                        return error!;
                    }

                    values.Add(value);
                }
            }

            if (shuffle)
            {
                Shuffle(values);
            }

            var result = new object[values.Count, 1];
            for (var i = 0; i < values.Count; i++)
            {
                result[i, 0] = values[i];
            }

            return result;
        }

        private static void Shuffle(List<object> values)
        {
            for (var i = values.Count - 1; i > 0; i--)
            {
                var swapIndex = Random.Shared.Next(i + 1);
                (values[i], values[swapIndex]) = (values[swapIndex], values[i]);
            }
        }

        private static bool TryParseCount(object input, out int repeatCount, out object? error)
        {
            if (!DataHelper.TryGetDouble(input, out var countDouble))
            {
                repeatCount = 0;
                error = ExcelErrors.Value;
                return false;
            }

            if (Math.Abs(countDouble - Math.Round(countDouble)) > 1e-9)
            {
                repeatCount = 0;
                error = ExcelErrors.Num;
                return false;
            }

            repeatCount = (int)Math.Round(countDouble);
            if (repeatCount < 1)
            {
                error = ExcelErrors.Num;
                return false;
            }

            error = null;
            return true;
        }

        private static bool TryCreateValueFactory(
            object what,
            out ValueFactory valueFactory,
            out object? error)
        {
            if (what is string text && LooksLikeFormula(text))
            {
                var formula = text.StartsWith('=') ? text : "=" + text;
                valueFactory = (out object value, out object? localError) => TryEvaluateScalarFormula(formula, out value, out localError);
                error = null;
                return true;
            }

            if (!TryGetScalarValue(what, out var scalarValue))
            {
                valueFactory = null!;
                error = ExcelErrors.Value;
                return false;
            }

            valueFactory = (out object value, out object? localError) =>
            {
                value = scalarValue;
                localError = null;
                return true;
            };
            error = null;
            return true;
        }

        private static bool TryEvaluateScalarFormula(string formula, out object value, out object? error)
        {
            try
            {
                var evaluated = XlCall.Excel(XlCall.xlfEvaluate, formula);
                if (!TryGetScalarValue(evaluated, out value))
                {
                    error = ExcelErrors.Value;
                    return false;
                }

                error = null;
                return true;
            }
            catch
            {
                value = string.Empty;
                error = ExcelErrors.Value;
                return false;
            }
        }

        private static bool TryGetScalarValue(object? value, out object scalar)
        {
            switch (value)
            {
                case object[,] matrix when matrix.GetLength(0) == 1 && matrix.GetLength(1) == 1:
                    scalar = matrix[0, 0];
                    return true;
                case object[,] or object[]:
                    scalar = string.Empty;
                    return false;
                default:
                    scalar = value ?? string.Empty;
                    return true;
            }
        }

        private static bool LooksLikeFormula(string text)
            => text.TrimStart().StartsWith('=');
    }
}
