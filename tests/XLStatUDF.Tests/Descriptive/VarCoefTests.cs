/// <summary>
/// Testy pro variační koeficienty.
/// </summary>
namespace XLStatUDF.Tests.Descriptive
{
    using Xunit;
    using XLStatUDF.Functions.Descriptive;

    public sealed class VarCoefTests
    {
        [Fact]
        public void ReturnsPopulationAndSampleVariationCoefficients()
        {
            var values = new object[] { 1.0, 2.0, 3.0, 4.0 };
            var weights = new object[] { 1.0, 1.0, 1.0, 2.0 };

            var population = VarCoef.VarCoefP(values);
            var sample = VarCoef.VarCoefS(values);
            var weightedPopulation = VarCoef.VarCoefW(values, weights);
            var weightedSample = VarCoef.VarCoefSW(values, weights);

            Assert.Equal(0.4472135955, (double)population, 8);
            Assert.Equal(0.5163977795, (double)sample, 8);
            Assert.Equal(0.4164965639, (double)weightedPopulation, 8);
            Assert.Equal(0.4656573147, (double)weightedSample, 8);
        }
    }
}
