/// <summary>
/// Testy pro WILCOXON.PAIRED.
/// </summary>
namespace XLStatUDF.Tests.Tests
{
    using Xunit;
    using XLStatUDF.Functions.Tests;

    public sealed class WilcoxonPairedTests
    {
        [Fact]
        public void ReturnsStructuredWilcoxonOutput()
        {
            var result = WilcoxonPaired.Run(
                new object[] { 10.0, 12.0, 15.0, 18.0, 20.0 },
                new object[] { 9.0, 11.0, 14.0, 17.0, 18.0 },
                null,
                0.05,
                0.0);

            Assert.Equal("n", result[0, 0]);
            Assert.Equal(5, result[0, 1]);
            Assert.Equal("med(d)", result[1, 0]);
            Assert.Equal("W+", result[3, 0]);
            Assert.Equal("W-", result[4, 0]);
            Assert.Equal("W", result[5, 0]);
            Assert.Equal("z", result[6, 0]);
            Assert.Equal("p", result[8, 0]);
            Assert.Equal("r", result[9, 0]);
        }

        [Fact]
        public void SupportsHeaders()
        {
            var result = WilcoxonPaired.Run(
                new object[] { "Pred", 10.0, 12.0, 15.0, 18.0, 20.0 },
                new object[] { "Po", 9.0, 11.0, 14.0, 17.0, 18.0 },
                1.0,
                0.05,
                0.0);

            Assert.Equal("n", result[0, 0]);
            Assert.Equal(5, result[0, 1]);
        }

        [Fact]
        public void RejectsWhenAllDifferencesAreZero()
        {
            var result = WilcoxonPaired.Run(
                new object[] { 1.0, 2.0, 3.0 },
                new object[] { 1.0, 2.0, 3.0 });

            Assert.Equal("#POČET!", result[0, 0]);
        }
    }
}
