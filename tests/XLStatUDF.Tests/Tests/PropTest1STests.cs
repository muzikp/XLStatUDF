/// <summary>
/// Testy pro jednovýběrový test podílu.
/// </summary>
namespace XLStatUDF.Tests.Tests
{
    using Xunit;
    using XLStatUDF.Functions.Tests;

    public sealed class PropTest1STests
    {
        [Fact]
        public void ReturnsExpectedZStatistic()
        {
            var result = PropTest1S.Run(new object[] { 1.0, 1.0, 1.0, 0.0, 0.0 }, 0.5, 0.0);

            Assert.Equal("p̂", result[0, 0]);
            Assert.Equal(0.6, (double)result[0, 1], 10);
            Assert.Equal("α", result[4, 0]);
            Assert.Equal(0.05, (double)result[4, 1], 10);
            Assert.Equal("z", result[5, 0]);
            Assert.Equal(0.4472135955, (double)result[5, 1], 8);
        }
    }
}
