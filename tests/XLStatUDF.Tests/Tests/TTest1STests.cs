/// <summary>
/// Testy pro jednovýběrový t-test.
/// </summary>
namespace XLStatUDF.Tests.Tests
{
    using ExcelDna.Integration;
    using Xunit;
    using XLStatUDF.Functions.Tests;

    public sealed class TTest1STests
    {
        [Fact]
        public void ReturnsExpectedTestStatisticAndMetadata()
        {
            var result = TTest1S.Run(new object[] { 2.0, 4.0, 6.0, 8.0, 10.0 }, 5.0, 2.0);

            Assert.Equal("x̄", result[0, 0]);
            Assert.Equal(6.0, (double)result[0, 1], 10);
            Assert.Equal("α", result[4, 0]);
            Assert.Equal(0.05, (double)result[4, 1], 10);
            Assert.Equal("t", result[5, 0]);
            Assert.Equal(0.7071067812, (double)result[5, 1], 8);
            Assert.Equal("df", result[6, 0]);
            Assert.Equal(4, result[6, 1]);
            Assert.Equal("sₓ", result[2, 0]);
            Assert.IsType<double>(result[7, 1]);
        }

        [Fact]
        public void IgnoresLeadingHeaderInNumericRange()
        {
            var result = TTest1S.Run(new object[] { "Měření", 2.0, 4.0, 6.0, 8.0, 10.0 }, 5.0, 0.0, 0.05, 1.0);

            Assert.Equal("x̄", result[0, 0]);
            Assert.Equal(6.0, (double)result[0, 1], 10);
        }

        [Fact]
        public void AutoDetectHeaderModeSkipsLeadingHeader()
        {
            var result = TTest1S.Run(new object[] { "Měření", 2.0, 4.0, 6.0, 8.0, 10.0 }, 5.0, 0.0, 0.05, 0.0);

            Assert.Equal("x̄", result[0, 0]);
            Assert.Equal(6.0, (double)result[0, 1], 10);
        }

        [Fact]
        public void NoHeaderModeDoesNotSkipLeadingText()
        {
            var result = TTest1S.Run(new object[] { "Měření", 2.0, 4.0, 6.0, 8.0, 10.0 }, 5.0, 0.0, 0.05, 2.0);

            Assert.Equal(ExcelError.ExcelErrorValue, result[0, 0]);
        }
    }
}
