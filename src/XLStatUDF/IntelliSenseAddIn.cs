/// <summary>
/// Zapíná Excel-DNA IntelliSense pro popisy funkcí a argumentů přímo při psaní vzorců v Excelu.
/// </summary>
namespace XLStatUDF
{
    using ExcelDna.Integration;
    using ExcelDna.IntelliSense;

    public sealed class IntelliSenseAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            IntelliSenseServer.Uninstall();
        }
    }
}
