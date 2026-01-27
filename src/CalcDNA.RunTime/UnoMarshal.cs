namespace CalcDNA.Runtime;

public static class UnoMarshal
{
    public static T[,] To2DArray<T>(object[][] source)
    {
        int rows = source.Length;
        int cols = rows > 0 ? source[0].Length : 0;
        T[,] target = new T[rows, cols];

        for (int i = 0; i < rows; i++)
            for (int j = 0; j < cols; j++)
                target[i, j] = (T)Convert.ChangeType(source[i][j], typeof(T));

        return target;
    }

    public static double UnwrapOptionalDouble(object value, double defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is double)
            return (double)value;
        throw new InvalidCastException("Cannot convert value to double.");
    }
}