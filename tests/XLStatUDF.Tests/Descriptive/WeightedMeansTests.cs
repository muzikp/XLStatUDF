/// <summary>
/// Testy pro vážené průměry.
/// </summary>
namespace XLStatUDF.Tests.Descriptive
{
    using Xunit;
    using XLStatUDF.Functions.Descriptive;

    public sealed class WeightedMeansTests
    {
        [Fact]
        public void AverageW_ComputesWeightedArithmeticMean()
        {
            var result = WeightedMeans.AverageW(
                new object[] { 1.0, 2.0, 5.0 },
                new object[] { 1.0, 1.0, 2.0 });

            Assert.IsType<double>(result);
            Assert.Equal(3.25, (double)result, 10);
        }

        [Fact]
        public void HarmonicAndGeometricMeans_RespectWeights()
        {
            var harmonic = WeightedMeans.HarMeanW(
                new object[] { 1.0, 2.0, 4.0 },
                new object[] { 1.0, 1.0, 2.0 });
            var geometric = WeightedMeans.GeoMeanW(
                new object[] { 1.0, 2.0, 4.0 },
                new object[] { 1.0, 1.0, 2.0 });

            Assert.Equal(2.0, (double)harmonic, 10);
            Assert.Equal(2.3784142300, (double)geometric, 8);
        }
    }
}
