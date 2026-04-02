/// <summary>
/// Testy pro jednovýběrový Kolmogorov-Smirnovův test.
/// </summary>
namespace XLStatUDF.Tests.Normality
{
    using Xunit;
    using XLStatUDF.Functions.Normality;

    public sealed class KolmogorovSmirnovTests
    {
        [Fact]
        public void ReturnsStatisticAndPValueForNormalDistribution()
        {
            var result = KolmogorovSmirnov.Run(new object[] { -1.2, -0.7, -0.1, 0.0, 0.2, 0.5, 0.9, 1.1 }, 0.0);

            Assert.Equal("D", result[0, 0]);
            Assert.InRange((double)result[0, 1], 0.0, 1.0);
            Assert.Equal("p", result[1, 0]);
            Assert.InRange((double)result[1, 1], 0.0, 1.0);
        }

        [Fact]
        public void RejectsInvalidLognormalInput()
        {
            var result = KolmogorovSmirnov.Run(new object[] { -1.0, 1.0, 2.0, 3.0, 4.0 }, 1.0);
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorNum, result[0, 0]);
        }
    }
}
