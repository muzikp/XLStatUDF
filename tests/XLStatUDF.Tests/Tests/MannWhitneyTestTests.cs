/// <summary>
/// Testy pro MANN.WHITNEY.G.
/// </summary>
namespace XLStatUDF.Tests.Tests
{
    using Xunit;
    using XLStatUDF.Functions.Tests;

    public sealed class MannWhitneyTestTests
    {
        [Fact]
        public void ReturnsStructuredMannWhitneyOutput()
        {
            var result = MannWhitneyTest.Run(
                new object[] { "A", "A", "A", "B", "B", "B" },
                new object[] { 1.0, 2.0, 3.0, 7.0, 8.0, 9.0 },
                null,
                0.05,
                0.0);

            Assert.Equal("Skupina", result[0, 0]);
            Assert.Equal("A", result[1, 0]);
            Assert.Equal("B", result[2, 0]);
            Assert.Equal("Výsledky testu", result[4, 0]);
            Assert.Equal("U", result[6, 0]);
            Assert.Equal("z", result[9, 0]);
            Assert.Equal("p", result[11, 0]);
            Assert.Equal("r", result[12, 0]);
        }

        [Fact]
        public void SupportsHeaders()
        {
            var result = MannWhitneyTest.Run(
                new object[] { "Skupina", "A", "A", "A", "B", "B", "B" },
                new object[] { "Hodnota", 1.0, 2.0, 3.0, 7.0, 8.0, 9.0 },
                1.0,
                0.05,
                0.0);

            Assert.Equal("A", result[1, 0]);
            Assert.Equal("B", result[2, 0]);
        }

        [Fact]
        public void RejectsWhenThereAreNotExactlyTwoGroups()
        {
            var result = MannWhitneyTest.Run(
                new object[] { "A", "A", "B", "B", "C", "C" },
                new object[] { 1.0, 2.0, 3.0, 4.0, 5.0, 6.0 });

            Assert.Equal("#POČET!", result[0, 0]);
        }
    }
}
