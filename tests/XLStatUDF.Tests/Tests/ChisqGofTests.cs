/// <summary>
/// Testy pro chí-kvadrát test dobré shody.
/// </summary>
namespace XLStatUDF.Tests.Tests
{
    using Xunit;
    using XLStatUDF.Functions.Tests;

    public sealed class ChisqGofTests
    {
        [Fact]
        public void ComputesExpectedStatisticFromObservedAndExpectedCounts()
        {
            var result = ChisqGof.Run(
                new object[] { 42.0, 63.0, 48.0, 47.0 },
                new object[] { 50.0, 50.0, 50.0, 50.0 },
                new object[] { "Červená", "Modrá", "Zelená", "Žlutá" },
                0.05);

            Assert.Equal("χ² GOF", result[0, 0]);
            Assert.Equal("χ²", result[1, 0]);
            Assert.Equal(4.92, (double)result[1, 1], 2);
            Assert.Equal("α", result[3, 0]);
            Assert.Equal(0.05, (double)result[3, 1], 10);
            Assert.Equal("KATEGORIE", result[7, 0]);
        }

        [Fact]
        public void AutoDetectSkipsCategoryHeader()
        {
            var result = ChisqGof.Run(
                new object[] { "Pozorovane", 42.0, 63.0, 48.0, 47.0 },
                new object[] { "Ocekavane", 50.0, 50.0, 50.0, 50.0 },
                new object[] { "Kategorie", "Cervena", "Modra", "Zelena", "Zluta" },
                0.05,
                0.0);

            Assert.Equal("KATEGORIE", result[7, 0]);
            Assert.Equal("Cervena", result[9, 0]);
        }
    }
}
