namespace XLStatUDF.Helpers.Formatting
{
    using System.Reflection;

    internal static class ExcelReflection
    {
        public static object? TryGetProperty(object? target, string propertyName)
        {
            if (target is null)
            {
                return null;
            }

            try
            {
                return target.GetType().InvokeMember(propertyName, BindingFlags.GetProperty, null, target, null);
            }
            catch
            {
                return null;
            }
        }

        public static int TryGetIntProperty(object? target, string collectionPropertyName, string countPropertyName)
        {
            var collection = TryGetProperty(target, collectionPropertyName);
            var count = TryGetProperty(collection, countPropertyName);
            return count switch
            {
                int value => value,
                short value => value,
                long value => (int)value,
                double value => (int)value,
                _ => 0
            };
        }

        public static bool TrySetProperty(object? target, string propertyName, object value)
        {
            if (target is null)
            {
                return false;
            }

            try
            {
                target.GetType().InvokeMember(propertyName, BindingFlags.SetProperty, null, target, [value]);
                return true;
            }
            catch
            {
                return false;
            }
        }

        public static object? TryInvokeMember(object? target, string memberName, params object[] args)
        {
            if (target is null)
            {
                return null;
            }

            try
            {
                return target.GetType().InvokeMember(memberName, BindingFlags.InvokeMethod | BindingFlags.GetProperty, null, target, args);
            }
            catch
            {
                return null;
            }
        }

        public static object? TryGetSubRange(object spillRange, int rowStart, int rowEnd, int columnStart, int columnEnd)
        {
            var topLeft = TryInvokeMember(spillRange, "Cells", rowStart, columnStart);
            var bottomRight = TryInvokeMember(spillRange, "Cells", rowEnd, columnEnd);
            if (topLeft is null || bottomRight is null)
            {
                return null;
            }

            return TryInvokeMember(spillRange, "Range", topLeft, bottomRight);
        }
    }
}
