/// <summary>
/// Testy pro spill funkci FILL.
/// </summary>
namespace XLStatUDF.Tests.Distributions
{
    using System.Linq;
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

        [Fact]
        public void SupportsMultipleWhatCountPairs()
        {
            var result = Fill.FillDown("muz", 2.0, "zena", 2.0, "dite", 1.0);

            var matrix = Assert.IsType<object[,]>(result);
            Assert.Equal(5, matrix.GetLength(0));
            Assert.Equal("muz", matrix[0, 0]);
            Assert.Equal("muz", matrix[1, 0]);
            Assert.Equal("zena", matrix[2, 0]);
            Assert.Equal("zena", matrix[3, 0]);
            Assert.Equal("dite", matrix[4, 0]);
        }

        [Fact]
        public void RejectsOddNumberOfAdditionalArguments()
        {
            var result = Fill.FillDown("x", 1.0, "y");
            Assert.Equal(ExcelDna.Integration.ExcelError.ExcelErrorValue, result);
        }

        [Fact]
        public void FillRandomPreservesRequestedCounts()
        {
            var result = Fill.FillDownRandom("muz", 3.0, "zena", 2.0, "dite", 1.0);

            var matrix = Assert.IsType<object[,]>(result);
            Assert.Equal(6, matrix.GetLength(0));

            var values = Enumerable.Range(0, matrix.GetLength(0))
                .Select(index => Assert.IsType<string>(matrix[index, 0]))
                .ToArray();

            Assert.Equal(3, values.Count(value => value == "muz"));
            Assert.Equal(2, values.Count(value => value == "zena"));
            Assert.Equal(1, values.Count(value => value == "dite"));
        }
    }
}
