/// <summary>
/// Testy pro vážené rozptyly a směrodatné odchylky.
/// </summary>
namespace XLStatUDF.Tests.Descriptive
{
    using Xunit;
    using XLStatUDF.Functions.Descriptive;

    public sealed class WeightedVarianceTests
    {
        [Fact]
        public void PopulationAndSampleWeightedVariance_MatchReferenceValues()
        {
            var values = new object[] { 1.0, 2.0, 5.0 };
            var weights = new object[] { 1.0, 1.0, 2.0 };

            var populationVariance = WeightedVariance.VarPW(values, weights);
            var sampleVariance = WeightedVariance.VarSW(values, weights);
            var populationStdev = WeightedVariance.StdevPW(values, weights);
            var sampleStdev = WeightedVariance.StdevSW(values, weights);

            Assert.Equal(3.1875, (double)populationVariance, 10);
            Assert.Equal(4.25, (double)sampleVariance, 10);
            Assert.Equal(Math.Sqrt(3.1875), (double)populationStdev, 10);
            Assert.Equal(Math.Sqrt(4.25), (double)sampleStdev, 10);
        }
    }
}
