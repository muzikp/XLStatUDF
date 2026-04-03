/// <summary>
/// Testy pro CORREL.MATRIX.
/// </summary>
namespace XLStatUDF.Tests.Correlation
{
    using Xunit;
    using XLStatUDF.Functions.Correlation;

    public sealed class CorrelMatrixTests
    {
        [Fact]
        public void ReturnsPearsonCoefficientMatrixByDefault()
        {
            var result = CorrelMatrix.Run(
                null,
                new object[,]
                {
                    { "A", "B", "C" },
                    { 1.0, 2.0, 3.0 },
                    { 2.0, 4.0, 6.0 },
                    { 3.0, 6.0, 9.0 },
                    { 4.0, 8.0, 12.0 }
                },
                null,
                null,
                1.0);

            Assert.Equal("A", result[0, 1]);
            Assert.Equal("B", result[0, 2]);
            Assert.Equal("C", result[0, 3]);
            Assert.Equal(1.0, (double)result[1, 2], 10);
            Assert.Equal(1.0, (double)result[2, 3], 10);
        }

        [Fact]
        public void ReturnsPValuesWhenRequested()
        {
            var result = CorrelMatrix.Run(
                null,
                new object[,]
                {
                    { 1.0, 1.0 },
                    { 2.0, 2.0 },
                    { 3.0, 3.0 },
                    { 4.0, 4.0 }
                },
                0.0,
                1.0,
                2.0);

            Assert.Equal("Proměnná 1", result[0, 1]);
            Assert.Equal(0.0, (double)result[1, 2], 10);
        }

        [Fact]
        public void ReturnsStackedSpearmanOutputWithSignificance()
        {
            var result = CorrelMatrix.Run(
                null,
                new object[,]
                {
                    { "X", "Y" },
                    { 1.0, 10.0 },
                    { 2.0, 20.0 },
                    { 3.0, 30.0 },
                    { 4.0, 40.0 }
                },
                1.0,
                3.0,
                1.0);

            Assert.Equal("X", result[1, 0]);
            Assert.Equal("p (X)", result[2, 0]);
            Assert.Equal("sig. (X)", result[3, 0]);
            Assert.Equal("***", result[3, 2]);
        }

        [Fact]
        public void ReturnsCoefficientStringsWithSignificanceWhenRequested()
        {
            var result = CorrelMatrix.Run(
                null,
                new object[,]
                {
                    { "X", "Y" },
                    { 1.0, 10.0 },
                    { 2.0, 20.0 },
                    { 3.0, 30.0 },
                    { 4.0, 40.0 }
                },
                0.0,
                4.0,
                1.0);

            Assert.Equal("1***", result[1, 2]);
            Assert.Equal("1***", result[2, 1]);
        }

        [Fact]
        public void FiltersOnlySignificantLinksWhenPMinimumIsSpecified()
        {
            var result = CorrelMatrix.Run(
                0.05,
                new object[,]
                {
                    { "A", "B", "C" },
                    { 1.0, 1.0, 1.0 },
                    { 2.0, 2.0, 1.1 },
                    { 3.0, 3.0, 0.9 },
                    { 4.0, 4.0, 1.2 },
                    { 5.0, 5.0, 0.8 }
                },
                0.0,
                4.0,
                1.0);

            Assert.Equal(string.Empty, result[1, 1]);
            Assert.Equal("1***", result[1, 2]);
            Assert.Equal(string.Empty, result[1, 3]);
        }
    }
}
