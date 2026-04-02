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
        public void ReturnsSingleColumnSpillWithRequestedCount()
        {
            var result = GenerateNorm.GenerateNormalSample(10.0, 2.0, 5.0);

            var matrix = Assert.IsType<object[,]>(result);
            Assert.Equal(5, matrix.GetLength(0));
            Assert.Equal(1, matrix.GetLength(1));

            for (var i = 0; i < 5; i++)
            {
                Assert.IsType<double>(matrix[i, 0]);
            }
        }

        [Fact]
        public void RejectsInvalidCount()
        {
            var result = GenerateNorm.GenerateNormalSample(10.0, 2.0, 0.0);
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorNum, result);
        }
    }
}
