/// <summary>
/// Testy pro intervalovou pravděpodobnost normálního rozdělení.
/// </summary>
namespace XLStatUDF.Tests.Distributions
{
    using Xunit;
    using XLStatUDF.Functions.Distributions;

    public sealed class NormDistRangeTests
    {
        [Fact]
        public void ReturnsProbabilityWithinOneSigma()
        {
            var result = NormDistRange.NormalDistributionRange(0.0, 1.0, -1.0, 1.0);

            Assert.IsType<double>(result);
            Assert.Equal(0.6826894921, (double)result, 8);
        }
    }
}
