/// <summary>
/// Testy pro Shapiro-Wilkův test normality.
/// </summary>
namespace XLStatUDF.Tests.Normality
{
    using Xunit;
    using XLStatUDF.Functions.Normality;

    public sealed class ShapiroWilkTests
    {
        [Fact]
        public void ReturnsHigherStatisticForMoreNormalSample()
        {
            var moreNormal = ShapiroWilk.Run(new object[] { -1.2, -0.7, -0.1, 0.0, 0.2, 0.5, 0.9, 1.1 });
            var lessNormal = ShapiroWilk.Run(new object[] { -3.0, -2.0, -1.0, 0.0, 0.1, 0.2, 4.0, 8.0 });

            Assert.Equal("W", moreNormal[0, 0]);
            Assert.Equal("W", lessNormal[0, 0]);
            Assert.True((double)moreNormal[0, 1] > (double)lessNormal[0, 1]);
            Assert.Equal("p", moreNormal[1, 0]);
            Assert.InRange((double)moreNormal[1, 1], 0.0, 1.0);
        }

        [Fact]
        public void RejectsTooShortSamples()
        {
            var result = ShapiroWilk.Run(new object[] { 1.0, 2.0 });
            Assert.Equal("#POČET!", result[0, 0]);
        }

        [Fact]
        public void IgnoresLeadingHeader()
        {
            var result = ShapiroWilk.Run(new object[] { "Hodnoty", -1.2, -0.7, -0.1, 0.0, 0.2, 0.5, 0.9, 1.1 }, 1.0);

            Assert.Equal("W", result[0, 0]);
            Assert.InRange((double)result[0, 1], 0.0, 1.0);
        }
    }
}
