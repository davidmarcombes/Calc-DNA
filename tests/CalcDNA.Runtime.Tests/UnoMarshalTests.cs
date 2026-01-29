namespace CalcDNA.Runtime.Tests;

public class UnoMarshalTests
{
    #region To1DArray Tests

    [Fact]
    public void To1DArray_WithValidDoubles_ReturnsTypedArray()
    {
        object[] source = [1.0, 2.0, 3.0];
        var result = UnoMarshal.To1DArray<double>(source);
        Assert.Equal([1.0, 2.0, 3.0], result);
    }

    [Fact]
    public void To1DArray_WithNullSource_ReturnsEmptyArray()
    {
        var result = UnoMarshal.To1DArray<int>(null);
        Assert.Empty(result);
    }

    [Fact]
    public void To1DArray_WithMixedNumericTypes_ConvertsCorrectly()
    {
        // UNO often sends doubles for all numeric types
        // Convert.ToInt32 rounds to nearest even (banker's rounding)
        object[] source = [1.5, 2.7, 3.9];
        var result = UnoMarshal.To1DArray<int>(source);
        Assert.Equal([2, 3, 4], result); // Rounded
    }

    [Fact]
    public void To1DArray_WithInvalidType_ThrowsException()
    {
        object[] source = ["not", "numbers"];
        Assert.Throws<InvalidCastException>(() => UnoMarshal.To1DArray<int>(source));
    }

    #endregion

    #region To2DArray Tests

    [Fact]
    public void To2DArray_WithValidData_ReturnsTypedArray()
    {
        object[][] source = [[1.0, 2.0], [3.0, 4.0]];
        var result = UnoMarshal.To2DArray<double>(source);

        Assert.Equal(2, result.GetLength(0));
        Assert.Equal(2, result.GetLength(1));
        Assert.Equal(1.0, result[0, 0]);
        Assert.Equal(4.0, result[1, 1]);
    }

    [Fact]
    public void To2DArray_WithRaggedArray_PadsWithDefaults()
    {
        object[][] source = [[1.0, 2.0, 3.0], [4.0]]; // Second row shorter
        var result = UnoMarshal.To2DArray<double>(source);

        Assert.Equal(2, result.GetLength(0));
        Assert.Equal(3, result.GetLength(1));
        Assert.Equal(4.0, result[1, 0]);
        Assert.Equal(0.0, result[1, 1]); // Default
        Assert.Equal(0.0, result[1, 2]); // Default
    }

    [Fact]
    public void To2DArray_WithNullSource_ReturnsEmptyArray()
    {
        var result = UnoMarshal.To2DArray<int>(null);
        Assert.Equal(0, result.GetLength(0));
        Assert.Equal(0, result.GetLength(1));
    }

    #endregion

    #region ToCalcRange Tests

    [Fact]
    public void ToCalcRange_WithValidData_ReturnsCalcRange()
    {
        object[][] source = [[1, 2], [3, 4]];
        var result = UnoMarshal.ToCalcRange(source);

        Assert.Equal(2, result.Rows);
        Assert.Equal(2, result.Columns);
        Assert.Equal(1, result[0, 0]);
    }

    [Fact]
    public void ToCalcRange_WithNull_ReturnsEmptyRange()
    {
        var result = UnoMarshal.ToCalcRange(null);
        Assert.Equal(0, result.Rows);
    }

    #endregion

    #region ToList Tests

    [Fact]
    public void ToList_WithValidData_ReturnsTypedList()
    {
        object[] source = [1.0, 2.0, 3.0];
        var result = UnoMarshal.ToList<double>(source);

        Assert.Equal(3, result.Count);
        Assert.Equal([1.0, 2.0, 3.0], result);
    }

    [Fact]
    public void ToList_WithNullSource_ReturnsEmptyList()
    {
        var result = UnoMarshal.ToList<string>(null);
        Assert.Empty(result);
    }

    [Fact]
    public void ToList_WithMixedNumericTypes_ConvertsCorrectly()
    {
        object[] source = [1.0, 2.0];
        var result = UnoMarshal.ToList<int>(source);
        Assert.Equal([1, 2], result);
    }

    #endregion

    #region ToUnoArray Tests

    [Fact]
    public void ToUno1DArray_WithValidData_ReturnsObjectArray()
    {
        int[] source = [1, 2, 3];
        var result = UnoMarshal.ToUno1DArray(source);
        Assert.Equal(3, result.Length);
        Assert.Equal(1, result[0]);
    }

    [Fact]
    public void ToUno1DArray_WithNull_ReturnsEmptyArray()
    {
        var result = UnoMarshal.ToUno1DArray<int>(null);
        Assert.Empty(result);
    }

    [Fact]
    public void ToUno2DArray_WithValidData_ReturnsJaggedArray()
    {
        int[,] source = { { 1, 2 }, { 3, 4 } };
        var result = UnoMarshal.ToUno2DArray(source);
        Assert.Equal(2, result.Length);
        Assert.Equal(2, result[0].Length);
        Assert.Equal(1, result[0][0]);
        Assert.Equal(4, result[1][1]);
    }

    [Fact]
    public void ToUno2DArray_WithNull_ReturnsEmptyJaggedArray()
    {
        var result = UnoMarshal.ToUno2DArray<int>(null);
        Assert.Empty(result);
    }

    [Fact]
    public void ToUno2DArray_WithCalcRange_ReturnsJaggedArray()
    {
        var range = new CalcRange(new object[,] { { 1, 2 } });
        var result = UnoMarshal.ToUno2DArray(range);
        Assert.Single(result);
        Assert.Equal(2, result[0].Length);
    }

    #endregion


    #region ToDateTime Tests

    [Fact]
    public void ToDateTime_WithOleDate_ConvertsCorrectly()
    {
        // OLE date 44197 = 2021-01-01
        var result = UnoMarshal.ToDateTime(44197.0);
        Assert.Equal(new DateTime(2021, 1, 1), result);
    }

    [Fact]
    public void ToDateTime_WithDateTime_ReturnsUnchanged()
    {
        var expected = new DateTime(2023, 6, 15);
        var result = UnoMarshal.ToDateTime(expected);
        Assert.Equal(expected, result);
    }

    [Fact]
    public void ToDateTime_WithNull_ThrowsException()
    {
        Assert.Throws<InvalidCastException>(() => UnoMarshal.ToDateTime(null));
    }

    #endregion

    #region UnwrapOptional Primitive Tests

    [Theory]
    [InlineData(true, false, true)]
    [InlineData(null, true, true)]
    [InlineData(false, true, false)]
    public void UnwrapOptionalBool_WithVariousInputs_ReturnsExpected(object? value, bool defaultValue, bool expected)
    {
        var result = UnoMarshal.UnwrapOptionalBool(value, defaultValue);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData(42, 0, 42)]
    [InlineData(null, 99, 99)]
    [InlineData(42.0, 0, 42)] // Double to int conversion
    public void UnwrapOptionalInt_WithVariousInputs_ReturnsExpected(object? value, int defaultValue, int expected)
    {
        var result = UnoMarshal.UnwrapOptionalInt(value, defaultValue);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData(3.14, 0.0, 3.14)]
    [InlineData(null, 1.5, 1.5)]
    [InlineData("", 2.5, 2.5)] // Empty string = empty
    public void UnwrapOptionalDouble_WithVariousInputs_ReturnsExpected(object? value, double defaultValue, double expected)
    {
        var result = UnoMarshal.UnwrapOptionalDouble(value, defaultValue);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("hello", "default", "hello")]
    [InlineData(null, "default", "default")]
    [InlineData("", "default", "default")] // Empty string = empty
    public void UnwrapOptionalString_WithVariousInputs_ReturnsExpected(object? value, string defaultValue, string expected)
    {
        var result = UnoMarshal.UnwrapOptionalString(value, defaultValue);
        Assert.Equal(expected, result);
    }

    [Fact]
    public void UnwrapOptionalDateTime_WithOleDate_ConvertsCorrectly()
    {
        var defaultDate = DateTime.MinValue;
        var result = UnoMarshal.UnwrapOptionalDateTime(44197.0, defaultDate);
        Assert.Equal(new DateTime(2021, 1, 1), result);
    }

    [Fact]
    public void UnwrapOptionalDateTime_WithNull_ReturnsDefault()
    {
        var defaultDate = new DateTime(2000, 1, 1);
        var result = UnoMarshal.UnwrapOptionalDateTime(null, defaultDate);
        Assert.Equal(defaultDate, result);
    }

    [Theory]
    [InlineData(10, 0, 10)]
    [InlineData(null, 5, 5)]
    public void UnwrapOptionalByte_WithVariousInputs_ReturnsExpected(object? value, byte defaultValue, byte expected)
    {
        var result = UnoMarshal.UnwrapOptionalByte(value, defaultValue);
        Assert.Equal(expected, result);
    }

    [Fact]
    public void UnwrapOptional1DArray_WithValidData_ReturnsArray()
    {
        object[] source = [1.0, 2.0];
        var result = UnoMarshal.UnwrapOptional1DArray(source, [0.0]);
        Assert.Equal([1.0, 2.0], result);
    }

    [Fact]
    public void UnwrapOptional1DArray_WithNull_ReturnsDefault()
    {
        var defaultValue = new double[] { 99.0 };
        var result = UnoMarshal.UnwrapOptional1DArray(null, defaultValue);
        Assert.Same(defaultValue, result);
    }

    [Fact]
    public void UnwrapOptional2DArray_WithValidData_ReturnsArray()
    {
        object[][] source = [[1.0]];
        var result = UnoMarshal.UnwrapOptional2DArray(source, new double[0, 0]);
        Assert.Equal(1.0, result[0, 0]);
    }

    [Fact]
    public void UnwrapOptionalCalcRange_WithValidData_ReturnsRange()
    {
        object[][] source = [[1, 2]];
        var result = UnoMarshal.UnwrapOptionalCalcRange(source, new CalcRange(new object[0, 0]));
        Assert.Equal(1, result[0, 0]);
    }

    [Fact]
    public void UnwrapOptionalList_WithValidData_ReturnsList()
    {
        object[] source = [1.0];
        var result = UnoMarshal.UnwrapOptionalList(source, new List<double>());
        Assert.Single(result);
    }

    #endregion


    #region UnwrapNullable Tests

    [Fact]
    public void UnwrapNullableInt_WithValue_ReturnsValue()
    {
        var result = UnoMarshal.UnwrapNullableInt(42);
        Assert.Equal(42, result);
    }

    [Fact]
    public void UnwrapNullableInt_WithNull_ReturnsNull()
    {
        var result = UnoMarshal.UnwrapNullableInt(null);
        Assert.Null(result);
    }

    [Fact]
    public void UnwrapNullableDouble_WithDoubleValue_ReturnsValue()
    {
        var result = UnoMarshal.UnwrapNullableDouble(3.14);
        Assert.Equal(3.14, result);
    }

    [Fact]
    public void UnwrapNullableString_WithNull_ReturnsNull()
    {
        var result = UnoMarshal.UnwrapNullableString(null);
        Assert.Null(result);
    }

    [Fact]
    public void UnwrapNullableDateTime_WithOleDate_ConvertsCorrectly()
    {
        var result = UnoMarshal.UnwrapNullableDateTime(44197.0);
        Assert.Equal(new DateTime(2021, 1, 1), result);
    }

    [Fact]
    public void UnwrapNullableBool_WithBoolValue_ReturnsValue()
    {
        var result = UnoMarshal.UnwrapNullableBool(true);
        Assert.True(result);
    }

    [Fact]
    public void UnwrapNullable1DArray_WithValidData_ReturnsArray()
    {
        object[] source = [1.0];
        var result = UnoMarshal.UnwrapNullable1DArray<double>(source);
        Assert.Single(result!);
    }

    [Fact]
    public void UnwrapNullable2DArray_WithNull_ReturnsNull()
    {
        var result = UnoMarshal.UnwrapNullable2DArray<double>(null);
        Assert.Null(result);
    }

    [Fact]
    public void UnwrapNullableCalcRange_WithValidData_ReturnsRange()
    {
        object[][] source = [[1, 2]];
        var result = UnoMarshal.UnwrapNullableCalcRange(source);
        Assert.Equal(1, result![0, 0]);
    }

    [Fact]
    public void UnwrapNullableList_WithValidData_ReturnsList()
    {
        object[] source = [1.0];
        var result = UnoMarshal.UnwrapNullableList<double>(source);
        Assert.Single(result!);
    }

    #endregion


    #region Edge Cases

    [Fact]
    public void UnwrapOptional_WithDBNull_TreatsAsEmpty()
    {
        var result = UnoMarshal.UnwrapOptionalInt(DBNull.Value, 42);
        Assert.Equal(42, result);
    }

    [Fact]
    public void UnwrapOptional_WithNestedEmptyArray_TreatsAsEmpty()
    {
        // 2D range with a single empty cell
        object[][] emptyJagged = [[null]];
        var result = UnoMarshal.UnwrapOptionalInt(emptyJagged, 123);
        Assert.Equal(123, result);
    }

    [Fact]
    public void ConvertValue_CharFromString_ReturnsFirstChar()
    {
        var result = UnoMarshal.ConvertValue<char>("ABC");
        Assert.Equal('A', result);
    }

    [Fact]
    public void ConvertValue_StringFromInt_ReturnsStringRep()
    {
        var result = UnoMarshal.ConvertValue<string>(123);
        Assert.Equal("123", result);
    }

    [Fact]
    public void UnwrapNullable_WithValidValue_ReturnsValue()
    {
        var result = UnoMarshal.UnwrapNullable<double>(1.23);
        Assert.Equal(1.23, result);
    }
    [Fact]
    public void IsEmpty_WithMissing_ReturnsTrue()
    {
        // System.Reflection.Missing is used in some COM/UNO scenarios
        var result = UnoMarshal.UnwrapOptionalInt(System.Reflection.Missing.Value, 42);
        Assert.Equal(42, result);
    }

    [Fact]
    public void IsEmpty_WithNestedEmptyJaggedArray_ReturnsTrue()
    {
        object[][] source = [[null]];
        var result = UnoMarshal.UnwrapOptionalInt(source, 100);
        Assert.Equal(100, result);
    }

    [Fact]
    public void ConvertValue_DateTimeFromDouble_ConvertsCorrectly()
    {
        // 44197.0 is 2021-01-01
        var result = UnoMarshal.ConvertValue<DateTime>(44197.0);
        Assert.Equal(new DateTime(2021, 1, 1), result);
    }

    [Fact]
    public void ConvertValue_NullableInt_ConvertsCorrectly()
    {
        var result = UnoMarshal.ConvertValue<int?>(42.0);
        Assert.Equal(42, result);
    }

    [Fact]
    public void ConvertValue_NullToNullable_ReturnsNull()
    {
        var result = UnoMarshal.ConvertValue<int?>(null);
        Assert.Null(result);
    }
    #endregion

}
