namespace CalcDNA.Runtime;

public static class UnoMarshal
{

    public static T[] To1DArray<T>(object[] source)
    {
        int rows = source.Length;
        T[] target = new T[rows];

        for (int i = 0; i < rows; i++)
            target[i] = (T)Convert.ChangeType(source[i], typeof(T));

        return target;
    }

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

    public static CalcRange ToCalcRange(object[][] source)
    {
        return new CalcRange(source);
    }

    public static List<T> ToList<T>(object[] source)
    {
        int rows = source.Length;
        List<T> target = new List<T>(rows);

        for (int i = 0; i < rows; i++)
            target.Add((T)Convert.ChangeType(source[i], typeof(T)));

        return target;
    }   

    public static bool UnwrapOptionalBool(object value, bool defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is bool)
            return (bool)value;
        throw new InvalidCastException("Cannot convert value to bool.");
    }

    public static sbyte UnwrapOptionalSByte(object value, sbyte defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is sbyte)
            return (sbyte)value;
        throw new InvalidCastException("Cannot convert value to sbyte.");
    }

    public static short UnwrapOptionalShort(object value, short defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is short)
            return (short)value;
        throw new InvalidCastException("Cannot convert value to short.");
    }

    public static int UnwrapOptionalInt(object value, int defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is int)
            return (int)value;
        throw new InvalidCastException("Cannot convert value to int.");
    }

    public static long UnwrapOptionalLong(object value, long defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is long)
            return (long)value;
        throw new InvalidCastException("Cannot convert value to long.");
    }

    public static double UnwrapOptionalDouble(object value, double defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is double)
            return (double)value;
        throw new InvalidCastException("Cannot convert value to double.");
    }

    public static float UnwrapOptionalFloat(object value, float defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is float)
            return (float)value;
        throw new InvalidCastException("Cannot convert value to float.");
    }

    public static char UnwrapOptionalChar(object value, char defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is char)
            return (char)value;
        throw new InvalidCastException("Cannot convert value to char.");
    }

    public static string UnwrapOptionalString(object value, string defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is string)
            return (string)value;
        throw new InvalidCastException("Cannot convert value to string.");
    }

    public static object UnwrapOptionalObject(object value, object defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        return value;
    }

    public static T[] UnwrapOptional1DArray<T>(object value, T[] defaultValue)
    {
        if (value == null)
            return defaultValue; // or any other default value for optional parameters
        if (value is object[])
            return To1DArray<T>((object[])value);
        throw new InvalidCastException("Cannot convert value to T[]");
    }

    public static T[,] UnwrapOptional2DArray<T>(object value, T[,] defaultValue)
    {
        if (value == null)
            return defaultValue;
        if (value is object[,])
            return To2DArray<T>((object[][])value);
        throw new InvalidCastException("Cannot convert value to T[,]" + value.GetType().Name);
    }

    public static CalcRange UnwrapOptionalCalcRange(object value, CalcRange defaultValue)
    {
        if (value == null)
            return defaultValue;
        if (value is object[][])
            return ToCalcRange((object[][])value);
        throw new InvalidCastException("Cannot convert value to CalcRange" + value.GetType().Name);
    }

    public static List<T> UnwrapOptionalList<T>(object value, List<T> defaultValue)
    {
        if (value == null)
            return defaultValue;
        if (value is object[])
            return ToList<T>((object[])value);
        throw new InvalidCastException("Cannot convert value to T" + value.GetType().Name);
    }
}