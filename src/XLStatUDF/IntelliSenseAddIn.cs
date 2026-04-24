namespace XLStatUDF
{
    using System.Linq;
    using ExcelDna.Integration;
    using ExcelDna.IntelliSense;
    using ExcelDna.Registration;
    using XLStatUDF.Helpers;
    using XLStatUDF.Helpers.Formatting;

    public sealed class IntelliSenseAddIn : IExcelAddIn
    {
        public void AutoOpen()
        {
            var registrations = ExcelRegistration.GetExcelFunctions().ToList();
            var localized = FunctionMetadataLocalizer.Apply(registrations);
            var processed = ExcelRegistration.ProcessFunctions(localized);
            ExcelRegistration.RegisterFunctions(processed);
            OutputFormatter.Initialize();
            IntelliSenseServer.Install();
        }

        public void AutoClose()
        {
            OutputFormatter.Shutdown();
            IntelliSenseServer.Uninstall();
        }
    }
}
