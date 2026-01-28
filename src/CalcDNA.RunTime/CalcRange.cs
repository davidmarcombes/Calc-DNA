namespace CalcDNA.Runtime;
/// <summary>
/// Represents a 2D grid of values, facilitating interoperability between 
/// multi-dimensional arrays and jagged arrays (UNO format).
/// </summary>
/// <remarks>
/// This class encapsulates an internal 2D object array to ensure a rectangular 
/// structure even when initialized from irregular (ragged) data.
/// </remarks>
public class CalcRange
{
    private readonly object[,] _data;

    /// <summary>
    /// Gets the total number of rows in the range.
    /// </summary>
    public int Rows => _data.GetLength(0);

    /// <summary>
    /// Gets the total number of columns in the range.
    /// </summary>
    public int Columns => _data.GetLength(1);

    /// <summary>
    /// Initializes a new instance of the <see cref="CalcRange"/> class using a 2D array.
    /// </summary>
    /// <param name="data">The 2D object array to wrap.</param>
    public CalcRange(object[,] data) => _data = data;

    /// <summary>
    /// Initializes a new instance of the <see cref="CalcRange"/> class from a jagged array.
    /// </summary>
    /// <param name="jaggedData">An array of arrays representing rows and columns.</param>
    /// <remarks>
    /// If the input rows are of unequal length, the range is normalized to the width 
    /// of the first row. Missing values are filled with <see langword="null"/>.
    /// </remarks>
    public CalcRange(object[][] jaggedData)
    {
        int rows = jaggedData.Length;
        int cols = rows > 0 ? jaggedData[0].Length : 0;
        _data = new object[rows, cols];

        for (int r = 0; r < rows; r++)
        {
            int rowLength = jaggedData[r].Length;
            for (int c = 0; c < cols; c++)
            {
                _data[r, c] = c < rowLength ? jaggedData[r][c] : null!;
            }
        }
    }

    /// <summary>
    /// Accesses the value at the specified row and column indices.
    /// </summary>
    /// <param name="row">The zero-based row index.</param>
    /// <param name="col">The zero-based column index.</param>
    /// <returns>The <see cref="object"/> stored at the specified location.</returns>
    /// <exception cref="IndexOutOfRangeException">Thrown when indices are outside the range bounds.</exception>
    public object this[int row, int col] => _data[row, col];

    /// <summary>
    /// Retrieves a value and attempts to convert it to the specified type <typeparamref name="T"/>.
    /// </summary>
    /// <typeparam name="T">The target type for the value conversion.</typeparam>
    /// <param name="row">The zero-based row index.</param>
    /// <param name="col">The zero-based column index.</param>
    /// <returns>The converted value of type <typeparamref name="T"/>.</returns>
    /// <exception cref="InvalidCastException">Thrown if the value cannot be converted to <typeparamref name="T"/>.</exception>
    public T GetAs<T>(int row, int col) => (T)Convert.ChangeType(_data[row, col], typeof(T));

    /// <summary>
    /// Returns an enumerator that iterates through all values in the range in row-major order.
    /// </summary>
    /// <returns>An <see cref="IEnumerable{Object}"/> containing all values in the range.</returns>
    public IEnumerable<object> Values()
    {
        for (int r = 0; r < Rows; r++)
            for (int c = 0; c < Columns; c++)
                yield return _data[r, c];
    }

    /// <summary>
    /// Converts the 2D range back into a jagged array (Array of Arrays).
    /// </summary>
    /// <returns>An array of object arrays compatible with UNO-style interfaces.</returns>
    public object[][] ToJaggedArray()
    {
        var result = new object[Rows][];
        for (int r = 0; r < Rows; r++)
        {
            result[r] = new object[Columns];
            for (int c = 0; c < Columns; c++)
                result[r][c] = _data[r, c];
        }
        return result;
    }
}