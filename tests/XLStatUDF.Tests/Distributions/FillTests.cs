/// <summary>
/// Testy pro spill funkci FILL.
/// </summary>
namespace XLStatUDF.Tests.Distributions
{
    using Xunit;
    using XLStatUDF.Functions.Distributions;

    public sealed class FillTests
    {
        [Fact]
        public void RepeatsNumericValueIntoSingleColumnSpill()
        {
            var result = Fill.FillDown(42.0, 3.0);

            var matrix = Assert.IsType<object[,]>(result);
            Assert.Equal(3, matrix.GetLength(0));
            Assert.Equal(1, matrix.GetLength(1));
            Assert.Equal(42.0, matrix[0, 0]);
            Assert.Equal(42.0, matrix[1, 0]);
            Assert.Equal(42.0, matrix[2, 0]);
        }

        [Fact]
        public void RepeatsTextValueIntoSingleColumnSpill()
        {
            var result = Fill.FillDown("ahoj", 2.0);

            var matrix = Assert.IsType<object[,]>(result);
            Assert.Equal("ahoj", matrix[0, 0]);
            Assert.Equal("ahoj", matrix[1, 0]);
        }

        [Fact]
        public void RejectsInvalidCount()
        {
            var result = Fill.FillDown("x", 0.0);
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorNum, result);
        }

        [Fact]
        public void RejectsNonScalarInput()
        {
            var result = Fill.FillDown(new object[,] { { 1.0 }, { 2.0 } }, 2.0);
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorValue, result);
        }
    }
}
