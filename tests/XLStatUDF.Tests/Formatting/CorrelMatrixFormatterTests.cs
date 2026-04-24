namespace XLStatUDF.Tests.Formatting
{
    using Xunit;
    using XLStatUDF.Helpers.Formatting;

    public sealed class CorrelMatrixFormatterTests
    {
        [Theory]
        [InlineData(0, 1)]
        [InlineData(1, 1)]
        [InlineData(2, 2)]
        [InlineData(3, 3)]
        [InlineData(4, 1)]
        public void ReturnsExpectedBlockHeight(int layoutCode, int expected)
            => Assert.Equal(expected, CorrelMatrixFormatter.GetBlockHeight((CorrelMatrixOutputLayout)layoutCode));
    }
}
