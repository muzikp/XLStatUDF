/// <summary>
/// Sestavuje dvou a vícesloupcové spill výstupy ve formátu object[,].
/// </summary>
namespace XLStatUDF.Helpers
{
    using System.Linq;

    public static class SpillBuilder
    {
        public static void AddRow(List<object[]> rows, string label, object value)
            => rows.Add([label, value]);

        public static void AddSeparator(List<object[]> rows, int cols)
            => rows.Add(Enumerable.Repeat<object>(string.Empty, cols).ToArray());

        public static void AddHeader(List<object[]> rows, string title, int cols)
        {
            var row = Enumerable.Repeat<object>(string.Empty, cols).ToArray();
            row[0] = title;
            rows.Add(row);
        }

        public static object[,] Build(List<object[]> rows)
        {
            var rowCount = rows.Count;
            var colCount = rows.Count == 0 ? 0 : rows.Max(r => r.Length);
            var result = new object[rowCount, colCount];

            for (var i = 0; i < rowCount; i++)
            {
                for (var j = 0; j < colCount; j++)
                {
                    result[i, j] = j < rows[i].Length ? rows[i][j] : string.Empty;
                }
            }

            return result;
        }
    }
}
