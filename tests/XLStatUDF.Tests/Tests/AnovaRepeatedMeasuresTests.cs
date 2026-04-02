/// <summary>
/// Testy pro ANOVA.RM.
/// </summary>
namespace XLStatUDF.Tests.Tests
{
    using Xunit;
    using XLStatUDF.Functions.Tests;

    public sealed class AnovaRepeatedMeasuresTests
    {
        [Fact]
        public void ReturnsStructuredRepeatedMeasuresOutput()
        {
            var result = AnovaRepeatedMeasures.Run(
                new object[,]
                {
                    { "T1", "T2", "T3" },
                    { 10.0, 12.0, 14.0 },
                    { 9.0, 11.0, 13.0 },
                    { 11.0, 13.0, 15.0 },
                    { 12.0, 14.0, 16.0 }
                },
                1.0,
                0.05,
                2.0);

            static int FindRow(object[,] matrix, string label)
            {
                for (var row = 0; row < matrix.GetLength(0); row++)
                {
                    if (Equals(matrix[row, 0], label))
                    {
                        return row;
                    }
                }

                return -1;
            }

            Assert.Equal("POPISNE STATISTIKY", result[0, 0]);
            Assert.Equal("sₓ", result[1, 4]);
            Assert.Equal("ANOVA S OPAKOVANYM MERENIM", result[6, 0]);
            Assert.Equal("Podminky", result[8, 0]);
            Assert.Equal("POZNAMKA", result[15, 0]);
            Assert.True(FindRow(result, "POST-HOC: BONFERRONI") >= 0);
        }

        [Fact]
        public void SupportsDefaultNamesWithoutHeader()
        {
            var result = AnovaRepeatedMeasures.Run(
                new object[,]
                {
                    { 10.0, 12.0 },
                    { 9.0, 11.0 },
                    { 11.0, 13.0 }
                });

            Assert.Equal("Podminka 1", result[2, 0]);
            Assert.Equal("Podminka 2", result[3, 0]);
        }

        [Fact]
        public void RejectsSingleCondition()
        {
            var result = AnovaRepeatedMeasures.Run(
                new object[,]
                {
                    { 10.0 },
                    { 11.0 },
                    { 12.0 }
                });

            Assert.Equal("#POČET!", result[0, 0]);
        }
    }
}
