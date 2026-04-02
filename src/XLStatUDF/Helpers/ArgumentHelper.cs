/// <summary>
/// Parsuje společné pomocné argumenty funkcí, například příznak přítomnosti záhlaví.
/// </summary>
namespace XLStatUDF.Helpers
{
    public enum HeaderMode
    {
        AutoDetect = 0,
        HasHeader = 1,
        NoHeader = 2
    }

    public static class ArgumentHelper
    {
        public static bool TryParseHeaderMode(object? input, out HeaderMode headerMode)
        {
            switch (input)
            {
                case null:
                case ExcelDna.Integration.ExcelMissing:
                case ExcelDna.Integration.ExcelEmpty:
                    headerMode = HeaderMode.AutoDetect;
                    return true;
                case double number:
                    if (number is 0 or 1 or 2)
                    {
                        headerMode = (HeaderMode)(int)number;
                        return true;
                    }

                    break;
                case int number:
                    if (number is 0 or 1 or 2)
                    {
                        headerMode = (HeaderMode)number;
                        return true;
                    }

                    break;
            }

            headerMode = HeaderMode.AutoDetect;
            return false;
        }
    }
}
