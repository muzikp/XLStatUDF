/// <summary>
/// Testy pro Spearmanův korelační koeficient.
/// </summary>
namespace XLStatUDF.Tests.Correlation
{
    using Xunit;
    using XLStatUDF.Functions.Correlation;

    public sealed class SpearmanCorrelTests
    {
        [Fact]
        public void ReturnsPerfectCorrelationForMonotonicData()
        {
            var result = SpearmanCorrel.Run(
                new object[] { 10.0, 20.0, 30.0, 40.0, 50.0 },
                new object[] { 1.0, 2.0, 3.0, 4.0, 5.0 },
                0.0);

            Assert.Equal("ρ", result[0, 0]);
            Assert.Equal(1.0, (double)result[0, 1], 10);
            Assert.Equal("n", result[1, 0]);
            Assert.Equal(5, result[1, 1]);
            Assert.Equal("α", result[2, 0]);
            Assert.Equal(0.05, (double)result[2, 1], 10);
        }

        [Fact]
        public void AcceptsNumericDirectionCodes()
        {
            var result = SpearmanCorrel.Run(
                new object[] { 1.0, 2.0, 3.0, 4.0, 5.0 },
                new object[] { 5.0, 4.0, 3.0, 2.0, 1.0 },
                1.0);

            Assert.Equal("p", result[6, 0]);
        }
    }
}
