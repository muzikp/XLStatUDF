namespace XLStatUDF.Helpers.Formatting
{
    using System.Globalization;
    using ExcelDna.Integration;

    internal enum CorrelMatrixOutputLayout
    {
        Coefficients = 0,
        PValues = 1,
        CoefficientsAndPValues = 2,
        CoefficientsPValuesAndSignificance = 3,
        CoefficientsWithSignificance = 4
    }

    internal static class CorrelMatrixFormatter
    {
        private const string DecimalFormat = "0.00000";
        private const int WhiteColor = 16777215;
        private const int MediumDarkGrayColor = 8421504;
        private const int NoLineStyle = -4142;
        private const int LowestValueCondition = 1;
        private const int HighestValueCondition = 2;
        private const int NumberCondition = 0;
        private const int NegativeFillColor = 13421823;
        private const int NeutralFillColor = 16777215;
        private const int PositiveFillColor = 13434828;

        public static OutputFormatter.FormattedRangeState? Apply(OutputFormatter.FormattingRequest request)
        {
            var callerReference = new ExcelReference(request.RowFirst, request.RowFirst, request.ColumnFirst, request.ColumnFirst, request.SheetId);
            var callerAddress = Convert.ToString(XlCall.Excel(XlCall.xlfReftext, callerReference, true), CultureInfo.InvariantCulture);
            if (string.IsNullOrWhiteSpace(callerAddress))
            {
                return null;
            }

            var callerRange = GetRange(callerAddress);
            if (callerRange is null)
            {
                return null;
            }

            var spillRange = TryGetSpillRange(callerRange) ?? callerRange;
            var formattedAddress = GetAddress(spillRange);
            if (string.IsNullOrWhiteSpace(formattedAddress))
            {
                return null;
            }

            ApplyBaseStyle(spillRange);
            ApplyLayoutFormats(spillRange, request.Layout);
            return OutputFormatter.FormattedRangeState.FromRequest(request, callerAddress, formattedAddress);
        }

        public static bool IsFormulaStillCompatible(OutputFormatter.FormattedRangeState state)
        {
            var callerRange = GetRange(state.CallerAddress);
            if (callerRange is null)
            {
                return false;
            }

            var formula = Convert.ToString(ExcelReflection.TryGetProperty(callerRange, "Formula"), CultureInfo.InvariantCulture);
            if (!LooksLikeCorrelMatrixFormula(formula))
            {
                return false;
            }

            var currentSpillRange = TryGetSpillRange(callerRange) ?? callerRange;
            var currentAddress = GetAddress(currentSpillRange);
            if (!string.IsNullOrWhiteSpace(currentAddress) && !state.HasSameFormattedAddress(currentAddress))
            {
                Clear(state.FormattedAddress);
            }

            return true;
        }

        public static void Clear(string formattedAddress)
        {
            var range = GetRange(formattedAddress);
            if (range is null)
            {
                return;
            }

            var formatConditions = ExcelReflection.TryGetProperty(range, "FormatConditions");
            ExcelReflection.TryInvokeMember(formatConditions, "Delete");

            var borders = ExcelReflection.TryGetProperty(range, "Borders");
            ExcelReflection.TrySetProperty(borders, "LineStyle", NoLineStyle);

            ExcelReflection.TrySetProperty(ExcelReflection.TryGetProperty(range, "Interior"), "Color", WhiteColor);
            ExcelReflection.TryInvokeMember(range, "ClearFormats");
        }

        internal static int GetBlockHeight(CorrelMatrixOutputLayout layout)
            => layout switch
            {
                CorrelMatrixOutputLayout.CoefficientsAndPValues => 2,
                CorrelMatrixOutputLayout.CoefficientsPValuesAndSignificance => 3,
                _ => 1
            };

        private static object? TryGetSpillRange(object callerRange)
            => ExcelReflection.TryGetProperty(callerRange, "SpillingToRange")
                ?? ExcelReflection.TryGetProperty(callerRange, "CurrentArray");

        private static void ApplyBaseStyle(object spillRange)
        {
            ExcelReflection.TrySetProperty(spillRange, "HorizontalAlignment", -4108);
            ExcelReflection.TrySetProperty(spillRange, "VerticalAlignment", -4108);
            ExcelReflection.TrySetProperty(ExcelReflection.TryGetProperty(spillRange, "Interior"), "Color", WhiteColor);

            var topRow = ExcelReflection.TryInvokeMember(ExcelReflection.TryGetProperty(spillRange, "Rows"), "Item", 1);
            var firstColumn = ExcelReflection.TryInvokeMember(ExcelReflection.TryGetProperty(spillRange, "Columns"), "Item", 1);
            if (topRow is not null)
            {
                ExcelReflection.TrySetProperty(ExcelReflection.TryGetProperty(topRow, "Font"), "Bold", true);
                ExcelReflection.TrySetProperty(topRow, "HorizontalAlignment", -4108);
            }

            if (firstColumn is not null)
            {
                ExcelReflection.TrySetProperty(ExcelReflection.TryGetProperty(firstColumn, "Font"), "Bold", true);
            }

            var borders = ExcelReflection.TryGetProperty(spillRange, "Borders");
            if (borders is not null)
            {
                ExcelReflection.TrySetProperty(borders, "LineStyle", 1);
                ExcelReflection.TrySetProperty(borders, "Weight", 2);
                ExcelReflection.TrySetProperty(borders, "Color", MediumDarkGrayColor);
            }
        }

        private static void ApplyLayoutFormats(object spillRange, CorrelMatrixOutputLayout layout)
        {
            var rowCount = ExcelReflection.TryGetIntProperty(spillRange, "Rows", "Count");
            var columnCount = ExcelReflection.TryGetIntProperty(spillRange, "Columns", "Count");
            if (rowCount < 2 || columnCount < 2)
            {
                return;
            }

            switch (layout)
            {
                case CorrelMatrixOutputLayout.Coefficients:
                case CorrelMatrixOutputLayout.PValues:
                    ApplyNumberFormat(spillRange, 2, rowCount, 2, columnCount, DecimalFormat);
                    break;
                case CorrelMatrixOutputLayout.CoefficientsAndPValues:
                case CorrelMatrixOutputLayout.CoefficientsPValuesAndSignificance:
                    ApplyStackedLayout(spillRange, rowCount, columnCount, layout);
                    break;
                case CorrelMatrixOutputLayout.CoefficientsWithSignificance:
                    ExcelReflection.TrySetProperty(spillRange, "NumberFormat", "@");
                    break;
            }
        }

        private static void ApplyStackedLayout(object spillRange, int rowCount, int columnCount, CorrelMatrixOutputLayout layout)
        {
            var blockHeight = GetBlockHeight(layout);
            BoldColumn(spillRange, 2);
            ClearConditionalFormatting(spillRange);

            for (var row = 2; row <= rowCount; row += blockHeight)
            {
                ApplyNumberFormat(spillRange, row, row, 3, columnCount, DecimalFormat);
                ApplyCoefficientColorScale(spillRange, row, columnCount);

                if (blockHeight >= 2 && row + 1 <= rowCount)
                {
                    ApplyNumberFormat(spillRange, row + 1, row + 1, 3, columnCount, DecimalFormat);
                }

                if (blockHeight >= 3 && row + 2 <= rowCount)
                {
                    ApplyNumberFormat(spillRange, row + 2, row + 2, 3, columnCount, "@");
                }
            }
        }

        private static void BoldColumn(object spillRange, int columnIndex)
        {
            var column = ExcelReflection.TryGetSubRange(
                spillRange,
                1,
                ExcelReflection.TryGetIntProperty(spillRange, "Rows", "Count"),
                columnIndex,
                columnIndex);

            if (column is null)
            {
                return;
            }

            ExcelReflection.TrySetProperty(ExcelReflection.TryGetProperty(column, "Font"), "Bold", true);
        }

        private static void ApplyNumberFormat(object spillRange, int rowStart, int rowEnd, int columnStart, int columnEnd, string numberFormat)
        {
            var target = ExcelReflection.TryGetSubRange(spillRange, rowStart, rowEnd, columnStart, columnEnd);
            if (target is null)
            {
                return;
            }

            ExcelReflection.TrySetProperty(target, "NumberFormat", numberFormat);
        }

        private static void ClearConditionalFormatting(object spillRange)
        {
            var formatConditions = ExcelReflection.TryGetProperty(spillRange, "FormatConditions");
            ExcelReflection.TryInvokeMember(formatConditions, "Delete");
        }

        private static void ApplyCoefficientColorScale(object spillRange, int rowIndex, int columnCount)
        {
            var target = ExcelReflection.TryGetSubRange(spillRange, rowIndex, rowIndex, 3, columnCount);
            if (target is null)
            {
                return;
            }

            var formatConditions = ExcelReflection.TryGetProperty(target, "FormatConditions");
            ExcelReflection.TryInvokeMember(formatConditions, "Delete");

            var colorScale = ExcelReflection.TryInvokeMember(formatConditions, "AddColorScale", 3);
            var criteria = ExcelReflection.TryGetProperty(colorScale, "ColorScaleCriteria");
            if (criteria is null)
            {
                return;
            }

            ConfigureColorScaleCriterion(criteria, 1, LowestValueCondition, null, NegativeFillColor);
            ConfigureColorScaleCriterion(criteria, 2, NumberCondition, 0, NeutralFillColor);
            ConfigureColorScaleCriterion(criteria, 3, HighestValueCondition, null, PositiveFillColor);
        }

        private static void ConfigureColorScaleCriterion(object criteria, int index, int type, object? value, int color)
        {
            var criterion = ExcelReflection.TryInvokeMember(criteria, "Item", index);
            if (criterion is null)
            {
                return;
            }

            ExcelReflection.TrySetProperty(criterion, "Type", type);
            if (value is not null)
            {
                ExcelReflection.TrySetProperty(criterion, "Value", value);
            }

            var formatColor = ExcelReflection.TryGetProperty(criterion, "FormatColor");
            ExcelReflection.TrySetProperty(formatColor, "Color", color);
        }

        private static object? GetRange(string address)
        {
            var application = ExcelDnaUtil.Application;
            return application is null || string.IsNullOrWhiteSpace(address)
                ? null
                : ExcelReflection.TryInvokeMember(application, "Range", address);
        }

        private static string? GetAddress(object range)
            => Convert.ToString(ExcelReflection.TryGetProperty(range, "Address"), CultureInfo.InvariantCulture);

        private static bool LooksLikeCorrelMatrixFormula(string? formula)
            => !string.IsNullOrWhiteSpace(formula)
                && formula.Contains("CORREL.MATRIX", StringComparison.OrdinalIgnoreCase);
    }
}
