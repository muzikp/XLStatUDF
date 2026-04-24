namespace XLStatUDF.Helpers.Formatting
{
    using System.Collections.Concurrent;
    using ExcelDna.Integration;

    internal static class OutputFormatter
    {
        private static readonly ConcurrentDictionary<string, byte> PendingRequests = new(StringComparer.Ordinal);
        private static readonly ConcurrentDictionary<string, FormattedRangeState> FormattedRanges = new(StringComparer.Ordinal);
        private static int sweepPending;

        public static void Initialize()
        {
            ExcelAsyncUtil.CalculationEnded += OnCalculationEnded;
        }

        public static void Shutdown()
        {
            ExcelAsyncUtil.CalculationEnded -= OnCalculationEnded;
        }

        public static void ScheduleCorrelMatrix(CorrelMatrixOutputLayout layout)
        {
            try
            {
                if (XlCall.Excel(XlCall.xlfCaller) is not ExcelReference caller)
                {
                    return;
                }

                var request = FormattingRequest.FromCaller(caller, layout);
                if (!PendingRequests.TryAdd(request.Key, 0))
                {
                    return;
                }

                ExcelAsyncUtil.QueueAsMacro(() =>
                {
                    try
                    {
                        var state = CorrelMatrixFormatter.Apply(request);
                        if (state is not null)
                        {
                            FormattedRanges.AddOrUpdate(request.CallerKey, state.Value, (_, _) => state.Value);
                        }
                    }
                    finally
                    {
                        PendingRequests.TryRemove(request.Key, out _);
                    }
                });
            }
            catch
            {
                // Formatting is a best-effort enhancement and must never break UDF evaluation.
            }
        }

        private static void OnCalculationEnded()
        {
            if (FormattedRanges.IsEmpty || Interlocked.Exchange(ref sweepPending, 1) == 1)
            {
                return;
            }

            ExcelAsyncUtil.QueueAsMacro(() =>
            {
                try
                {
                    foreach (var pair in FormattedRanges)
                    {
                        if (!CorrelMatrixFormatter.IsFormulaStillCompatible(pair.Value))
                        {
                            CorrelMatrixFormatter.Clear(pair.Value.FormattedAddress);
                            FormattedRanges.TryRemove(pair.Key, out _);
                        }
                    }
                }
                finally
                {
                    Interlocked.Exchange(ref sweepPending, 0);
                }
            });
        }

        internal readonly record struct FormattingRequest(
            nint SheetId,
            int RowFirst,
            int ColumnFirst,
            CorrelMatrixOutputLayout Layout)
        {
            public string CallerKey => $"{SheetId}:{RowFirst}:{ColumnFirst}";

            public string Key => $"{CallerKey}:{(int)Layout}";

            public static FormattingRequest FromCaller(ExcelReference caller, CorrelMatrixOutputLayout layout)
                => new(caller.SheetId, caller.RowFirst, caller.ColumnFirst, layout);
        }

        internal readonly record struct FormattedRangeState(
            nint SheetId,
            int RowFirst,
            int ColumnFirst,
            CorrelMatrixOutputLayout Layout,
            string CallerAddress,
            string FormattedAddress)
        {
            public bool HasSameFormattedAddress(string address)
                => string.Equals(FormattedAddress, address, StringComparison.OrdinalIgnoreCase);

            public static FormattedRangeState FromRequest(FormattingRequest request, string callerAddress, string formattedAddress)
                => new(request.SheetId, request.RowFirst, request.ColumnFirst, request.Layout, callerAddress, formattedAddress);
        }
    }
}
