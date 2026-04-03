/// <summary>
/// Testy pro generování náhodných celých čísel.
/// </summary>
namespace XLStatUDF.Tests.Distributions
{
    using Xunit;
    using XLStatUDF.Functions.Distributions;

    public sealed class GenerateIntTests
    {
        [Fact]
        public void ReturnsSingleScalarValueWithinRequestedInterval()
        {
            var result = GenerateInt.GenerateIntegerSample(10.0, 20.0);
            var value = Assert.IsType<long>(result);
            Assert.InRange(value, 10L, 20L);
        }

        [Fact]
        public void UsesDefaultArgumentsWhenOmitted()
        {
            var result = GenerateInt.GenerateIntegerSample();
            Assert.IsType<long>(result);
        }

        [Fact]
        public void RejectsInvertedInterval()
        {
            var result = GenerateInt.GenerateIntegerSample(10.0, 5.0);
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorNum, result);
        }

        [Theory]
        [InlineData(-0.1)]
        [InlineData(1.1)]
        public void RejectsInvalidOutlierRate(double outlierRate)
        {
            var result = GenerateInt.GenerateIntegerSample(10.0, 20.0, outlierRate);
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorValue, result);
        }
    }
}
