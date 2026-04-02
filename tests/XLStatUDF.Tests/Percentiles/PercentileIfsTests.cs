/// <summary>
/// Testy pro percentily s filtrováním.
/// </summary>
namespace XLStatUDF.Tests.Percentiles
{
    using Xunit;
    using XLStatUDF.Functions.Percentiles;

    public sealed class PercentileIfsTests
    {
        [Fact]
        public void PercentileIncIfs_FiltersByTextAndNumberCriteria()
        {
            var result = PercentileIfs.PercentileIncIfs(
                new object[] { 5.0, 10.0, 15.0, 20.0, 25.0 },
                0.5,
                new object[] { "A", "A", "B", "A", "B" }, "A",
                new object[] { 5.0, 10.0, 15.0, 20.0, 25.0 }, ">5");

            Assert.Equal(15.0, (double)result, 10);
        }

        [Fact]
        public void PercentileExcIfs_SupportsWildcardCriteria()
        {
            var result = PercentileIfs.PercentileExcIfs(
                new object[] { 10.0, 20.0, 30.0, 40.0, 50.0 },
                0.5,
                new object[] { "Praha", "Brno", "Plzeň", "Prachatice", "Ostrava" }, "Pra*");

            Assert.Equal(25.0, (double)result, 10);
        }
    }
}
