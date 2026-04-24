/// <summary>
/// Testy pro pivot funkce.
/// </summary>
namespace XLStatUDF.Tests.Descriptive
{
    using System.Globalization;
    using ExcelDna.Integration;
    using Xunit;
    using XLStatUDF.Functions.Descriptive;
    using XLStatUDF.Helpers;

    public sealed class PivotTests
    {
        private static string TotalLabel => AddInLanguage.IsCzech ? "Celkem" : "Total";

        public PivotTests()
        {
            var culture = CultureInfo.GetCultureInfo("en-US");
            CultureInfo.CurrentCulture = culture;
            CultureInfo.CurrentUICulture = culture;
        }

        [Fact]
        public void BuildsAveragePivot()
        {
            var result = Pivot.PivotAverage(
                new object[,] { { "Group" }, { "A" }, { "A" }, { "B" }, { "B" } },
                new object[,] { { "Sex" }, { "M" }, { "F" }, { "M" }, { "F" } },
                new object[,] { { "Score" }, { 10.0 }, { 20.0 }, { 30.0 }, { 40.0 } });

            Assert.Equal("F", result[0, 1]);
            Assert.Equal("M", result[0, 2]);
            Assert.Equal(TotalLabel, result[0, 3]);
            Assert.Equal("Group", result[0, 0]);
            Assert.Equal("A", result[1, 0]);
            Assert.Equal(20.0, (double)result[1, 1], 10);
            Assert.Equal(10.0, (double)result[1, 2], 10);
            Assert.Equal(15.0, (double)result[1, 3], 10);
            Assert.Equal(TotalLabel, result[3, 0]);
            Assert.Equal(30.0, (double)result[3, 1], 10);
            Assert.Equal(20.0, (double)result[3, 2], 10);
            Assert.Equal(25.0, (double)result[3, 3], 10);
        }

        [Fact]
        public void SupportsConfidenceDirection()
        {
            var result = Pivot.PivotConfidenceNorm(
                new object[,] { { "Group" }, { "A" }, { "A" }, { "A" } },
                ExcelEmpty.Value,
                new object[,] { { "Score" }, { 10.0 }, { 20.0 }, { 30.0 } },
                0.05,
                1);

            Assert.Equal("Group", result[0, 0]);
            Assert.True((double)result[1, 1] > 0.0);
            Assert.Equal(TotalLabel, result[2, 0]);
        }

        [Fact]
        public void SupportsMadAndIqr()
        {
            var values = new object[,] { { "Value" }, { 1.0 }, { 2.0 }, { 3.0 }, { 100.0 } };
            var mad = Pivot.PivotMad(
                new object[,] { { "Group" }, { "A" }, { "A" }, { "A" }, { "A" } },
                ExcelEmpty.Value,
                values);
            var iqr = Pivot.PivotIqr(
                new object[,] { { "Group" }, { "A" }, { "A" }, { "A" }, { "A" } },
                ExcelEmpty.Value,
                values);

            Assert.Equal(1.0, (double)mad[1, 1], 10);
            Assert.Equal(1.0, (double)mad[2, 1], 10);
            Assert.Equal(25.5, (double)iqr[1, 1], 10);
            Assert.Equal(25.5, (double)iqr[2, 1], 10);
        }

        [Fact]
        public void RejectsInvalidDirection()
        {
            var result = Pivot.PivotConfidenceT(
                new object[,] { { "Group" }, { "A" }, { "A" } },
                ExcelEmpty.Value,
                new object[,] { { "Score" }, { 10.0 }, { 20.0 } },
                0.05,
                2);

            Assert.Equal(ExcelErrors.Value, result[0, 0]);
        }

        [Fact]
        public void SortsRowsAndColumnsAlphabetically()
        {
            var result = Pivot.PivotSum(
                new object[,] { { "Group" }, { "B" }, { "A" } },
                new object[,] { { "Type" }, { "Z" }, { "M" } },
                new object[,] { { "Score" }, { 20.0 }, { 10.0 } });

            Assert.Equal("M", result[0, 1]);
            Assert.Equal("Z", result[0, 2]);
            Assert.Equal(TotalLabel, result[0, 3]);
            Assert.Equal("A", result[1, 0]);
            Assert.Equal("B", result[2, 0]);
            Assert.Equal(10.0, (double)result[1, 1], 10);
            Assert.Equal(20.0, (double)result[2, 2], 10);
            Assert.Equal(TotalLabel, result[3, 0]);
        }

        [Fact]
        public void OmitsRowsWithOnlyBlankAggregates()
        {
            var result = Pivot.PivotSum(
                new object[,] { { "Genre" }, { "A" }, { "B" } },
                ExcelEmpty.Value,
                new object[,] { { "Sales" }, { 10.0 }, { ExcelEmpty.Value } });

            Assert.Equal(3, result.GetLength(0));
            Assert.Equal("A", result[1, 0]);
            Assert.Equal(TotalLabel, result[2, 0]);
        }

        [Fact]
        public void OmitsColumnsWithOnlyBlankAggregates()
        {
            var result = Pivot.PivotSum(
                new object[,] { { "Publisher" }, { "Activision" }, { "Electronic Arts" }, { "Ubisoft" } },
                new object[,] { { "Genre" }, { ExcelEmpty.Value }, { "Action" }, { "RPG" } },
                new object[,] { { "Sales" }, { ExcelEmpty.Value }, { 10.0 }, { 20.0 } });

            Assert.Equal(4, result.GetLength(1));
            Assert.Equal("Action", result[0, 1]);
            Assert.Equal("RPG", result[0, 2]);
            Assert.Equal(TotalLabel, result[0, 3]);
        }

        [Fact]
        public void ConfidencePivotLeavesSparseCellsBlankInsteadOfFailing()
        {
            var result = Pivot.PivotConfidenceT(
                new object[,] { { "Publisher" }, { "A" }, { "A" }, { "B" } },
                new object[,] { { "Genre" }, { "Action" }, { "Action" }, { "RPG" } },
                new object[,] { { "Sales" }, { 10.0 }, { 20.0 }, { 30.0 } },
                0.05,
                0);

            Assert.Equal("Action", result[0, 1]);
            Assert.Equal(3, result.GetLength(1));
            Assert.Equal("A", result[1, 0]);
            Assert.True((double)result[1, 1] > 0.0);
            Assert.Equal(TotalLabel, result[0, 2]);
        }

        [Fact]
        public void AddsTotalRowAndTotalColumn()
        {
            var result = Pivot.PivotSum(
                new object[,] { { "Publisher" }, { "Activision" }, { "Activision" }, { "Ubisoft" } },
                new object[,] { { "Genre" }, { "Action" }, { "RPG" }, { "Action" } },
                new object[,] { { "Sales" }, { 1.0 }, { 2.0 }, { 3.0 } });

            Assert.Equal(TotalLabel, result[0, 3]);
            Assert.Equal(TotalLabel, result[3, 0]);
            Assert.Equal(3.0, (double)result[1, 3], 10);
            Assert.Equal(4.0, (double)result[3, 1], 10);
            Assert.Equal(6.0, (double)result[3, 3], 10);
        }
    }
}
