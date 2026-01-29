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

        Assert.Equal(1, range[0, 0]);
        Assert.Equal(2, range[0, 1]);
        Assert.Equal(3, range[0, 2]);
        Assert.Equal(4, range[1, 0]);
        Assert.Equal(5, range[1, 1]);
        Assert.Equal(6, range[1, 2]);
    }

    [Fact]
    public void Constructor_WithJaggedArray_SetsCorrectDimensions()
    {
        var data = new object[][] { [1, 2, 3], [4, 5, 6] };
        var range = new CalcRange(data);

        Assert.Equal(2, range.Rows);
        Assert.Equal(3, range.Columns);

        Assert.Equal(1, range[0, 0]);
        Assert.Equal(2, range[0, 1]);
        Assert.Equal(3, range[0, 2]);
        Assert.Equal(4, range[1, 0]);
        Assert.Equal(5, range[1, 1]);
        Assert.Equal(6, range[1, 2]);

    }

    [Fact]
    public void Constructor_WithRaggedJaggedArray_PadsWithNull()
    {
        var data = new object[][] { [1, 2, 3], [4] }; // Second row shorter
        var range = new CalcRange(data);

        Assert.Equal(2, range.Rows);
        Assert.Equal(3, range.Columns);

        Assert.Equal(1, range[0, 0]);
        Assert.Equal(2, range[0, 1]);
        Assert.Equal(3, range[0, 2]);
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

    [Fact]
    public void GetAs_WithIncompatibleType_ReturnsDefault()
    {
        var data = new object[,] { { "not a number" } };
        var range = new CalcRange(data);
        Assert.Equal(0, range.GetAs<int>(0, 0));
    }

    [Fact]
    public void GetAs_WithNullValue_ReturnsDefault()
    {
        var data = new object[,] { { null } };
        var range = new CalcRange(data);
        Assert.Null(range.GetAs<string>(0, 0));
        Assert.Equal(0.0, range.GetAs<double>(0, 0));
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

    #region New Feature Tests

    [Fact]
    public void Indexer_Setter_UpdatesValue()
    {
        var range = new CalcRange(new object[2, 2]);
        range[0, 0] = "new value";
        Assert.Equal("new value", range[0, 0]);
    }

    [Fact]
    public void Count_ReturnsTotalElements()
    {
        var range = new CalcRange(new object[3, 4]);
        Assert.Equal(12, range.Count);
    }

    [Fact]
    public void GetRow_ReturnsCorrectData()
    {
        var data = new object[,] { { 1, 2 }, { 3, 4 } };
        var range = new CalcRange(data);
        var row = range.GetRow(1);
        Assert.Equal([3, 4], row);
    }

    [Fact]
    public void GetColumn_ReturnsCorrectData()
    {
        var data = new object[,] { { 1, 2 }, { 3, 4 } };
        var range = new CalcRange(data);
        var col = range.GetColumn(1);
        Assert.Equal([2, 4], col);
    }

    [Fact]
    public void Flatten_ReturnsCorrectArray()
    {
        var data = new object[,] { { 1, 2 }, { 3, 4 } };
        var range = new CalcRange(data);
        var flat = range.Flatten();
        Assert.Equal([1, 2, 3, 4], flat);
    }

    [Fact]
    public void IEnumerable_Implementation_Works()
    {
        var data = new object[,] { { 1, 2 }, { 3, 4 } };
        var range = new CalcRange(data);
        var list = range.Cast<int>().ToList();
        Assert.Equal([1, 2, 3, 4], list);
    }

    [Fact]
    public void ToString_ReturnsDimensions()
    {
        var range = new CalcRange(new object[5, 10]);
        Assert.Equal("CalcRange [5x10]", range.ToString());
    }

    [Fact]
    public void GetRow_OutOfBounds_ThrowsException()
    {
        var range = new CalcRange(new object[1, 1]);
        Assert.Throws<IndexOutOfRangeException>(() => range.GetRow(5));
    }

    [Fact]
    public void GetColumn_OutOfBounds_ThrowsException()
    {
        var range = new CalcRange(new object[1, 1]);
        Assert.Throws<IndexOutOfRangeException>(() => range.GetColumn(5));
    }

    #endregion
}
