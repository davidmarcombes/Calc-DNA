namespace CalcDNA.Runtime;

public class CalcRange
{
    private readonly object[,] _data;

    public int Rows => _data.GetLength(0);
    public int Columns => _data.GetLength(1);

    public CalcRange(object[,] data) => _data = data;

    // Indexer for easier access: range[row, col]
    public object this[int row, int col] => _data[row, col];

    // Helper: Convert to a specific type safely
    public T GetAs<T>(int row, int col) => (T)Convert.ChangeType(_data[row, col], typeof(T));

    // Helper: Iterate through all values (useful for LINQ)
    public IEnumerable<object> Values()
    {
        for (int r = 0; r < Rows; r++)
            for (int c = 0; c < Columns; c++)
                yield return _data[r, c];
    }
}