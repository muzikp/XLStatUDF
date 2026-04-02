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
        public void ReturnsSingleColumnSpillWithRequestedCount()
        {
            var result = GenerateInt.GenerateIntegerSample(5.0, 10.0, 20.0);

            var matrix = Assert.IsType<object[,]>(result);
            Assert.Equal(5, matrix.GetLength(0));
            Assert.Equal(1, matrix.GetLength(1));

            for (var i = 0; i < 5; i++)
            {
                var value = Assert.IsType<long>(matrix[i, 0]);
                Assert.InRange(value, 10L, 20L);
            }
        }

        [Fact]
        public void UsesDefaultArgumentsWhenOmitted()
        {
            var result = GenerateInt.GenerateIntegerSample();

            var matrix = Assert.IsType<object[,]>(result);
            Assert.Equal(1, matrix.GetLength(0));
            Assert.Equal(1, matrix.GetLength(1));
            Assert.IsType<long>(matrix[0, 0]);
        }

        [Fact]
        public void RejectsInvalidCount()
        {
            var result = GenerateInt.GenerateIntegerSample(0.0, 0.0, 10.0);
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorNum, result);
        }

        [Fact]
        public void RejectsInvertedInterval()
        {
            var result = GenerateInt.GenerateIntegerSample(3.0, 10.0, 5.0);
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorNum, result);
        }
    }
}
