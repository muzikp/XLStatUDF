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
            Description = "[Rozdělení] Zkopíruje hodnotu nebo vzorec do spill sloupce",
            Category = "XLStatUDF",
            IsVolatile = true)]
        public static object FillDown(
            [ExcelArgument(Name = "co", Description = "Hodnota nebo textovy vzorec zacinatejici =")] object what,
            [ExcelArgument(Name = "pocet", Description = "Pocet opakovani; cele cislo >= 1")] object count)
        {
            if (!DataHelper.TryGetDouble(count, out var countDouble))
            {
                return ExcelErrors.Value;
            }

            if (Math.Abs(countDouble - Math.Round(countDouble)) > 1e-9)
            {
                return ExcelErrors.Num;
            }

            var repeatCount = (int)Math.Round(countDouble);
            if (repeatCount < 1)
            {
                return ExcelErrors.Num;
            }

            if (!TryCreateValueFactory(what, out var valueFactory, out var error))
            {
                return error!;
            }

            var result = new object[repeatCount, 1];
            for (var i = 0; i < repeatCount; i++)
            {
                if (!valueFactory(out var value, out error))
                {
                    return error!;
                }

                result[i, 0] = value;
            }

            return result;
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
