/// <summary>
/// Pomocné funkce pro řazení a výpočet midrank pořadí při vázaných hodnotách.
/// </summary>
namespace XLStatUDF.Helpers
{
    public static class RankHelper
    {
        public static double[] MidRank(double[] values)
        {
            var indexed = values
                .Select((value, index) => (value, index))
                .OrderBy(item => item.value)
                .ToArray();

            var ranks = new double[values.Length];
            var i = 0;

            while (i < indexed.Length)
            {
                var j = i + 1;
                while (j < indexed.Length && indexed[j].value.Equals(indexed[i].value))
                {
                    j++;
                }

                var averageRank = ((i + 1) + j) / 2.0;
                for (var k = i; k < j; k++)
                {
                    ranks[indexed[k].index] = averageRank;
                }

                i = j;
            }

            return ranks;
        }
    }
}
