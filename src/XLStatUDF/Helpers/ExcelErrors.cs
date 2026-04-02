/// <summary>
/// Centrální definice excelových chyb a textových fallbacků pro chyby bez přímého Excel ekvivalentu.
/// </summary>
namespace XLStatUDF.Helpers
{
    using ExcelDna.Integration;

    public static class ExcelErrors
    {
        public static object Value => ExcelError.ExcelErrorValue;

        public static object Num => ExcelError.ExcelErrorNum;

        public static object NA => ExcelError.ExcelErrorNA;

        public static object DivZero => ExcelError.ExcelErrorDiv0;

        public static object Count => "#POČET!";

        public static object Length => "#DÉLKA!";
    }
}
