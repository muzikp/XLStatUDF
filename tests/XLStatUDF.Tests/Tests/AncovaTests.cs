/// <summary>
/// Testy pro ANCOVA.G.
/// </summary>
namespace XLStatUDF.Tests.Tests
{
    using Xunit;
    using XLStatUDF.Functions.Tests;

    public sealed class AncovaTests
    {
        [Fact]
        public void ReturnsStructuredAncovaOutput()
        {
            var result = Ancova.Run(
                new object[] { "A", "A", "A", "B", "B", "B", "C", "C", "C" },
                new object[] { 10.0, 12.0, 13.0, 16.0, 18.0, 19.0, 20.0, 22.0, 24.0 },
                new object[,]
                {
                    { 1.0, 10.0 },
                    { 2.0, 11.0 },
                    { 3.0, 12.0 },
                    { 1.0, 12.0 },
                    { 2.0, 13.0 },
                    { 3.0, 14.0 },
                    { 1.0, 14.0 },
                    { 2.0, 15.0 },
                    { 3.0, 16.0 }
                },
                2.0,
                0.05);

            Assert.Equal("POPISNE STATISTIKY", result[0, 0]);
            Assert.Equal("ANCOVA", result[6, 0]);
            Assert.Equal("η²", result[7, 6]);
            Assert.Equal("ω²p", result[7, 9]);
            Assert.Equal("Skupina × Kovariata 1", result[11, 0]);
            Assert.True(FindRow(result, "ADJUSTED MEANS") >= 0);
            Assert.True(FindRow(result, "POST-HOC: BONFERRONI") >= 0);
        }

        [Fact]
        public void SupportsHeadersAndDefaultPostHoc()
        {
            var result = Ancova.Run(
                new object[] { "Skupina", "A", "A", "A", "B", "B", "B" },
                new object[] { "Y", 10.0, 12.0, 13.0, 16.0, 18.0, 19.0 },
                new object[,]
                {
                    { "vek", "bmi" },
                    { 1.0, 10.0 },
                    { 2.0, 11.0 },
                    { 3.0, 12.0 },
                    { 1.0, 12.0 },
                    { 2.0, 13.0 },
                    { 3.0, 14.0 }
                },
                null,
                0.05,
                1.0);

            Assert.Equal("POPISNE STATISTIKY", result[0, 0]);
            Assert.Equal("vek", result[1, 3]);
            Assert.Equal("bmi", result[1, 4]);
            Assert.Equal(-1, FindRow(result, "POST-HOC: BONFERRONI"));
        }

        [Fact]
        public void RejectsWhenCovariateRangeLengthDoesNotMatch()
        {
            var result = Ancova.Run(
                new object[] { "A", "A", "B", "B" },
                new object[] { 1.0, 2.0, 3.0, 4.0 },
                new object[,]
                {
                    { 1.0 },
                    { 2.0 },
                    { 3.0 }
                });

            Assert.Equal("#DÉLKA!", result[0, 0]);
        }

        private static int FindRow(object[,] matrix, string label)
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
    }
}
