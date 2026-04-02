/// <summary>
/// Testy pro analyzu kontingencnich tabulek.
/// </summary>
namespace XLStatUDF.Tests.Tests
{
    using Xunit;
    using XLStatUDF.Functions.Tests;

    public sealed class ContingencyTests
    {
        [Fact]
        public void TableVariantComputesExpectedSections()
        {
            var result = Contingency.RunTable(
                new object[,]
                {
                    { "", "Muz", "Zena" },
                    { "Ano", 30.0, 20.0 },
                    { "Ne", 10.0, 40.0 }
                },
                1.0,
                0.05);

            Assert.Equal("POZOROVANÁ KONTINGENČNÍ TABULKA", result[0, 0]);
            Assert.Equal("Muz", result[1, 1]);
            Assert.Equal("Ano", result[2, 0]);
            Assert.Equal(30.0, (double)result[2, 1], 10);
            Assert.Equal("OČEKÁVANÉ ČETNOSTI", result[6, 0]);
            Assert.Equal("SOUHRN TESTU", result[12, 0]);
            Assert.Equal("χ²", result[16, 0]);
            Assert.Equal("MÍRY ASOCIACE", result[20, 0]);
            Assert.Equal("phi", result[23, 0]);
            Assert.IsType<double>(result[23, 1]);
        }

        [Fact]
        public void GroupedVariantBuildsTableFromImplicitCounts()
        {
            var result = Contingency.RunGrouped(
                new object[] { "Muz", "Muz", "Muz", "Zena", "Zena", "Zena" },
                new object[] { "Ano", "Ano", "Ne", "Ano", "Ne", "Ne" },
                null,
                0.05,
                null);

            Assert.Equal("POZOROVANÁ KONTINGENČNÍ TABULKA", result[0, 0]);
            Assert.Equal("Ano", result[2, 0]);
            Assert.Equal("Muz", result[1, 1]);
            Assert.Equal(2.0, (double)result[2, 1], 10);
            Assert.Equal(1.0, (double)result[2, 2], 10);
        }

        [Fact]
        public void GroupedVariantSupportsExplicitCountsAndHeaders()
        {
            var result = Contingency.RunGrouped(
                new object[] { "Sloupec", "Muz", "Muz", "Zena", "Zena" },
                new object[] { "Radek", "Ano", "Ne", "Ano", "Ne" },
                new object[] { "Pocet", 5.0, 3.0, 4.0, 8.0 },
                0.05,
                1.0);

            Assert.Equal(5.0, (double)result[2, 1], 10);
            Assert.Equal(4.0, (double)result[2, 2], 10);
            Assert.Equal(3.0, (double)result[3, 1], 10);
            Assert.Equal(8.0, (double)result[3, 2], 10);
        }

        [Fact]
        public void TableVariantWithoutHeadersAssignsNumericLabels()
        {
            var result = Contingency.RunTable(
                new object[,]
                {
                    { 1.0, 2.0, 3.0 },
                    { 3.0, 4.0, 5.0 }
                },
                2.0,
                0.05);

            Assert.Equal("1", result[1, 1]);
            Assert.Equal("1", result[2, 0]);
            Assert.Equal(string.Empty, result[23, 1]);
        }
    }
}
