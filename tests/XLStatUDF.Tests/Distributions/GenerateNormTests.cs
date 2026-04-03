/// <summary>
/// Testy pro generování náhodného normálního výběru.
/// </summary>
namespace XLStatUDF.Tests.Distributions
{
    using Xunit;
    using XLStatUDF.Functions.Distributions;

    public sealed class GenerateNormTests
    {
        [Fact]
        public void ReturnsSingleScalarValue()
        {
            var result = GenerateNorm.GenerateNormalSample(10.0, 2.0);
            Assert.IsType<double>(result);
        }

        [Fact]
        public void RejectsNonPositiveStandardDeviation()
        {
            var result = GenerateNorm.GenerateNormalSample(10.0, 0.0);
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorNum, result);
        }

        [Theory]
        [InlineData(-0.1)]
        [InlineData(1.1)]
        public void RejectsInvalidOutlierRate(double outlierRate)
        {
            var result = GenerateNorm.GenerateNormalSample(10.0, 2.0, outlierRate);
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorValue, result);
        }
    }
}
