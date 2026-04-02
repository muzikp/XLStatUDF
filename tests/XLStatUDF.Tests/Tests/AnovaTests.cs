/// <summary>
/// Testy pro jednofaktorovou ANOVA.
/// </summary>
namespace XLStatUDF.Tests.Tests
{
    using Xunit;
    using XLStatUDF.Functions.Tests;

    public sealed class AnovaTests
    {
        [Fact]
        public void ReturnsStructuredAnovaOutput()
        {
            var result = Anova.Run(
                new object[] { "A", "A", "A", "B", "B", "B", "C", "C", "C" },
                new object[] { 10.0, 11.0, 12.0, 15.0, 16.0, 17.0, 21.0, 22.0, 23.0 },
                null,
                0.05,
                2.0);

            Assert.Equal("POPISNÉ STATISTIKY", result[0, 0]);
            Assert.Equal("sₓ", result[1, 4]);
            Assert.Equal("ANOVA", result[6, 0]);
            Assert.Equal("OVĚŘENÍ PŘEDPOKLADŮ", result[14, 0]);
            Assert.Equal("VELIKOST ÚČINKU", result[18, 0]);
            Assert.Equal("POST-HOC: BONFERRONI", result[23, 0]);
            Assert.Equal("α", result[11, 0]);
            Assert.Equal(0.05, (double)result[11, 1], 10);
        }

        [Fact]
        public void RejectsLessThanThreeGroups()
        {
            var result = Anova.Run(
                new object[] { "A", "A", "B", "B" },
                new object[] { 1.0, 2.0, 3.0, 4.0 });

            Assert.Equal("#POČET!", result[0, 0]);
        }

        [Fact]
        public void GroupedAliasSupportsHeaders()
        {
            var result = Anova.Run(
                new object[] { "Skupina", "A", "A", "A", "B", "B", "B", "C", "C", "C" },
                new object[] { "Hodnota", 10.0, 11.0, 12.0, 15.0, 16.0, 17.0, 21.0, 22.0, 23.0 },
                1.0,
                0.05,
                2.0);

            Assert.Equal("POPISNÉ STATISTIKY", result[0, 0]);
        }

        [Fact]
        public void DisplaysGroupNamesUppercase()
        {
            var result = Anova.Run(
                new object[] { "muzi", "muzi", "muzi", "zeny", "zeny", "zeny", "deti", "deti", "deti" },
                new object[] { 10.0, 11.0, 12.0, 15.0, 16.0, 17.0, 21.0, 22.0, 23.0 });

            Assert.Equal("DETI", result[2, 0]);
            Assert.Equal("MUZI", result[3, 0]);
            Assert.Equal("ZENY", result[4, 0]);
        }

        [Fact]
        public void DefaultPostHocIsNone()
        {
            var result = Anova.Run(
                new object[] { "A", "A", "A", "B", "B", "B", "C", "C", "C" },
                new object[] { 10.0, 11.0, 12.0, 15.0, 16.0, 17.0, 21.0, 22.0, 23.0 });

            Assert.Equal("VELIKOST ÚČINKU", result[18, 0]);
            Assert.NotEqual("POST-HOC: BONFERRONI", result[result.GetLength(0) - 1, 0]);
        }

    }
}
