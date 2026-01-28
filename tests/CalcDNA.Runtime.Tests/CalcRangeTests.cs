namespace CalcDNA.Runtime.Tests;

public class CalcRangeTests
{
    #region Constructor Tests

    [Fact]
    public void Constructor_WithMultiDimensionalArray_SetsCorrectDimensions()
    {
        var data = new object[,] { { 1, 2, 3 }, { 4, 5, 6 } };
        var range = new CalcRange(data);

        Assert.Equal(2, range.Rows);
        Assert.Equal(3, range.Columns);
    }

    [Fact]
    public void Constructor_WithJaggedArray_SetsCorrectDimensions()
    {
        var data = new object[][] { [1, 2, 3], [4, 5, 6] };
        var range = new CalcRange(data);

        Assert.Equal(2, range.Rows);
        Assert.Equal(3, range.Columns);
    }

    [Fact]
    public void Constructor_WithRaggedJaggedArray_PadsWithNull()
    {
        var data = new object[][] { [1, 2, 3], [4] }; // Second row shorter
        var range = new CalcRange(data);

        Assert.Equal(2, range.Rows);
        Assert.Equal(3, range.Columns);
        Assert.Equal(4, range[1, 0]);
        Assert.Null(range[1, 1]);
        Assert.Null(range[1, 2]);
    }

    [Fact]
    public void Constructor_WithEmptyJaggedArray_CreatesEmptyRange()
    {
        var data = Array.Empty<object[]>();
        var range = new CalcRange(data);

        Assert.Equal(0, range.Rows);
        Assert.Equal(0, range.Columns);
    }

    #endregion

    #region Indexer Tests

    [Fact]
    public void Indexer_ReturnsCorrectValue()
    {
        var data = new object[,] { { "a", "b" }, { "c", "d" } };
        var range = new CalcRange(data);

        Assert.Equal("a", range[0, 0]);
        Assert.Equal("b", range[0, 1]);
        Assert.Equal("c", range[1, 0]);
        Assert.Equal("d", range[1, 1]);
    }

    [Fact]
    public void Indexer_OutOfRange_ThrowsException()
    {
        var data = new object[,] { { 1, 2 } };
        var range = new CalcRange(data);

        Assert.Throws<IndexOutOfRangeException>(() => range[5, 0]);
    }

    #endregion

    #region GetAs Tests

    [Fact]
    public void GetAs_WithCompatibleType_ConvertsCorrectly()
    {
        var data = new object[,] { { 1.5, 2.7 } };
        var range = new CalcRange(data);

        Assert.Equal(1.5, range.GetAs<double>(0, 0));
        Assert.Equal(2, range.GetAs<int>(0, 0)); // Rounded (Convert.ToInt32 uses banker's rounding)
    }

    [Fact]
    public void GetAs_WithStringToInt_ConvertsCorrectly()
    {
        var data = new object[,] { { "42" } };
        var range = new CalcRange(data);

        Assert.Equal(42, range.GetAs<int>(0, 0));
    }

    #endregion

    #region Values Tests

    [Fact]
    public void Values_ReturnsAllValuesInOrder()
    {
        var data = new object[,] { { 1, 2 }, { 3, 4 } };
        var range = new CalcRange(data);

        var values = range.Values().ToList();

        Assert.Equal(4, values.Count);
        Assert.Equal([1, 2, 3, 4], values);
    }

    [Fact]
    public void Values_WithEmptyRange_ReturnsEmptyEnumerable()
    {
        var data = new object[0, 0];
        var range = new CalcRange(data);

        Assert.Empty(range.Values());
    }

    #endregion

    #region ToJaggedArray Tests

    [Fact]
    public void ToJaggedArray_ReturnsCorrectStructure()
    {
        var data = new object[,] { { 1, 2 }, { 3, 4 } };
        var range = new CalcRange(data);

        var jagged = range.ToJaggedArray();

        Assert.Equal(2, jagged.Length);
        Assert.Equal(2, jagged[0].Length);
        Assert.Equal(1, jagged[0][0]);
        Assert.Equal(4, jagged[1][1]);
    }

    [Fact]
    public void ToJaggedArray_RoundTrips_WithJaggedConstructor()
    {
        var original = new object[][] { [1, 2, 3], [4, 5, 6] };
        var range = new CalcRange(original);
        var result = range.ToJaggedArray();

        Assert.Equal(original.Length, result.Length);
        for (int i = 0; i < original.Length; i++)
        {
            Assert.Equal(original[i], result[i]);
        }
    }

    #endregion
}
