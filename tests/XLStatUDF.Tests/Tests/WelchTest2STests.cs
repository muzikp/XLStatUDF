/// <summary>
/// Testy pro Welchův dvouvýběrový t-test.
/// </summary>
namespace XLStatUDF.Tests.Tests
{
    using Xunit;
    using XLStatUDF.Functions.Tests;

    public sealed class WelchTest2STests
    {
        [Fact]
        public void ReturnsExpectedStructureAndTestStatistic()
        {
            var result = WelchTest2S.Run(
                new object[] { "A", "A", "A", "A", "B", "B", "B", "B" },
                new object[] { 10.0, 12.0, 11.0, 13.0, 16.0, 18.0, 17.0, 19.0 },
                null,
                0.05,
                0.0);

            Assert.Equal("Skupina", result[0, 0]);
            Assert.Equal("sₓ", result[0, 4]);
            Assert.Equal("A", result[1, 0]);
            Assert.Equal("B", result[2, 0]);
            Assert.Equal("Výsledky testu", result[4, 0]);
            Assert.Equal("α", result[5, 0]);
            Assert.Equal(0.05, (double)result[5, 1], 10);
            Assert.Equal("t", result[6, 0]);
            Assert.True((double)result[6, 1] < 0);
            Assert.Equal("t₁₋α⁄₂", result[8, 0]);
            Assert.IsType<double>(result[8, 1]);
            Assert.Equal("p", result[9, 0]);
        }

        [Fact]
        public void RejectsWrongNumberOfGroups()
        {
            var result = WelchTest2S.Run(
                new object[] { "A", "A", "B", "C" },
                new object[] { 1.0, 2.0, 3.0, 4.0 });

            Assert.Equal("#POČET!", result[0, 0]);
        }

        [Fact]
        public void IgnoresLeadingHeadersInGroupedInputs()
        {
            var result = WelchTest2S.Run(
                new object[] { "Skupina", "A", "A", "A", "A", "B", "B", "B", "B" },
                new object[] { "Hodnota", 10.0, 12.0, 11.0, 13.0, 16.0, 18.0, 17.0, 19.0 },
                1.0,
                0.05,
                0.0);

            Assert.Equal("A", result[1, 0]);
            Assert.Equal("Výsledky testu", result[4, 0]);
        }

        [Fact]
        public void AutoDetectModeSkipsGroupedHeaders()
        {
            var result = WelchTest2S.Run(
                new object[] { "Skupina", "A", "A", "A", "A", "B", "B", "B", "B" },
                new object[] { "Hodnota", 10.0, 12.0, 11.0, 13.0, 16.0, 18.0, 17.0, 19.0 },
                0.0,
                0.05,
                0.0);

            Assert.Equal("A", result[1, 0]);
            Assert.Equal("Výsledky testu", result[4, 0]);
        }
    }
}
