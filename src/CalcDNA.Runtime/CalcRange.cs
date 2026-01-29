using System.Collections;

namespace CalcDNA.Runtime;

/// <summary>
/// Represents a 2D grid of values, facilitating interoperability between 
/// multi-dimensional arrays and jagged arrays (UNO format).
/// </summary>
/// <remarks>
/// This class encapsulates an internal 2D object array to ensure a rectangular 
/// structure even when initialized from irregular (ragged) data.
/// </remarks>
public class CalcRange : IEnumerable<object?>
{
    private readonly object?[,] _data;

    /// <summary>
    /// Gets the total number of rows in the range.
    /// </summary>
    public int Rows => _data.GetLength(0);

    /// <summary>
    /// Gets the total number of columns in the range.
    /// </summary>
    public int Columns => _data.GetLength(1);

    /// <summary>
    /// Gets the total number of elements in the range.
    /// </summary>
    public int Count => _data.Length;

    /// <summary>
    /// Initializes a new instance of the <see cref="CalcRange"/> class using a 2D array.
    /// </summary>
    /// <param name="data">The 2D object array to wrap.</param>
    public CalcRange(object?[,] data) => _data = data;

    /// <summary>
    /// Initializes a new instance of the <see cref="CalcRange"/> class from a jagged array.
    /// </summary>
    /// <param name="jaggedData">An array of arrays representing rows and columns.</param>
    /// <remarks>
    /// If the input rows are of unequal length, the range is normalized to the width 
    /// of the widest row. Missing values are filled with <see langword="null"/>.
    /// </remarks>
    public CalcRange(object?[][] jaggedData)
    {
        int rows = jaggedData.Length;
        int cols = 0;
        for (int r = 0; r < rows; r++)
        {
            if (jaggedData[r]?.Length > cols)
                cols = jaggedData[r].Length;
        }

        _data = new object?[rows, cols];

        for (int r = 0; r < rows; r++)
        {
            var row = jaggedData[r];
            if (row == null) continue;
            
            for (int c = 0; c < row.Length; c++)
            {
                _data[r, c] = row[c];
            }
        }
    }

    /// <summary>
    /// Accesses or modifies the value at the specified row and column indices.
    /// </summary>
    /// <param name="row">The zero-based row index.</param>
    /// <param name="col">The zero-based column index.</param>
    /// <returns>The <see cref="object"/> stored at the specified location.</returns>
    /// <exception cref="IndexOutOfRangeException">Thrown when indices are outside the range bounds.</exception>
    public object? this[int row, int col]
    {
        get => _data[row, col];
        set => _data[row, col] = value;
    }

    /// <summary>
    /// Retrieves a value and attempts to convert it to the specified type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The target type for the value conversion.</typeparam>
    /// <param name="row">The zero-based row index.</param>
    /// <param name="col">The zero-based column index.</param>
    /// <returns>The converted value of type <typeparamref name="T"/>, or default if conversion fails or value is null.</returns>
    public T? GetAs<T>(int row, int col)
    {
        var val = _data[row, col];
        if (val == null || val == DBNull.Value)
            return default;

        var targetType = typeof(T);
        
        // Direct type match
        if (val is T typed)
            return typed;

        // Handle nullable types
        var underlyingType = Nullable.GetUnderlyingType(targetType);
        if (underlyingType != null)
        {
            targetType = underlyingType;
        }

        // Use Convert for IConvertible types
        if (val is IConvertible)
        {
            try
            {
                return (T)Convert.ChangeType(val, targetType);
            }
            catch
            {
                // Fallback or just return default
                return default;
            }
        }

        return default;
    }

    /// <summary>
    /// Gets a specific row from the range.
    /// </summary>
    public object?[] GetRow(int rowIndex)
    {
        var result = new object?[Columns];
        for (int c = 0; c < Columns; c++)
            result[c] = _data[rowIndex, c];
        return result;
    }

    /// <summary>
    /// Gets a specific column from the range.
    /// </summary>
    public object?[] GetColumn(int colIndex)
    {
        var result = new object?[Rows];
        for (int r = 0; r < Rows; r++)
            result[r] = _data[r, colIndex];
        return result;
    }

    /// <summary>
    /// Returns an enumerator that iterates through all values in the range in row-major order.
    /// </summary>
    /// <returns>An <see cref="IEnumerator{Object}"/> containing all values in the range.</returns>
    public IEnumerator<object?> GetEnumerator()
    {
        for (int r = 0; r < Rows; r++)
            for (int c = 0; c < Columns; c++)
                yield return _data[r, c];
    }

    IEnumerator IEnumerable.GetEnumerator() => GetEnumerator();

    /// <summary>
    /// Backward compatible method for retrieving all values in row-major order.
    /// </summary>
    public IEnumerable<object?> Values() => this;

    /// <summary>
    /// Converts the 2D range back into a jagged array (Array of Arrays).
    /// </summary>
    /// <returns>An array of object arrays compatible with UNO-style interfaces.</returns>
    public object?[][] ToJaggedArray()
    {
        var result = new object?[Rows][];
        for (int r = 0; r < Rows; r++)
        {
            result[r] = new object?[Columns];
            for (int c = 0; c < Columns; c++)
                result[r][c] = _data[r, c];
        }
        return result;
    }

    /// <summary>
    /// Creates a 1D array with all values in row-major order.
    /// </summary>
    public object?[] Flatten()
    {
        var result = new object?[Count];
        int i = 0;
        foreach (var val in this)
        {
            result[i++] = val;
        }
        return result;
    }

    /// <summary>
    /// Returns a string representation of the range dimensions.
    /// </summary>
    public override string ToString() => $"CalcRange [{Rows}x{Columns}]";
}