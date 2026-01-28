namespace CalcDNA.Runtime;

/// <summary>
/// Provides marshaling utilities for converting between UNO types and .NET types.
/// UNO (LibreOffice's component model) uses untyped object arrays and specific type conventions
/// that need to be converted to strongly-typed .NET equivalents.
/// </summary>
public static class UnoMarshal
{
    /// <summary>
    /// Converts a UNO object array to a typed 1D array.
    /// </summary>
    public static T[] To1DArray<T>(object? source)
    {
        if (source == null)
            return Array.Empty<T>();
        if (source is not object[] arr)
            throw MarshalException("array", source);

        var target = new T[arr.Length];
        for (int i = 0; i < arr.Length; i++)
        {
            target[i] = ConvertValue<T>(arr[i]);
        }
        return target;
    }

    /// <summary>
    /// Converts a UNO jagged array to a typed 2D array.
    /// Handles ragged arrays by using default values for missing elements.
    /// </summary>
    public static T[,] To2DArray<T>(object? source)
    {
        if (source == null)
            return new T[0, 0];
        if (source is not object[][] jagged)
            throw MarshalException("2D array", source);

        int rows = jagged.Length;
        int cols = GetMaxColumns(jagged);
        var target = new T[rows, cols];

        for (int i = 0; i < rows; i++)
        {
            var row = jagged[i];
            int rowLen = row?.Length ?? 0;
            for (int j = 0; j < cols; j++)
            {
                target[i, j] = j < rowLen ? ConvertValue<T>(row![j]) : default!;
            }
        }
        return target;
    }

    /// <summary>
    /// Converts a UNO jagged array to a CalcRange.
    /// </summary>
    public static CalcRange ToCalcRange(object? source)
    {
        if (source == null)
            return new CalcRange(Array.Empty<object[]>());
        if (source is not object[][] jagged)
            throw MarshalException("CalcRange (object[][])", source);
        return new CalcRange(jagged);
    }

    /// <summary>
    /// Converts a UNO object array to a typed List.
    /// </summary>
    public static List<T> ToList<T>(object? source)
    {
        if (source == null)
            return new List<T>();
        if (source is not object[] arr)
            throw MarshalException("array", source);

        var target = new List<T>(arr.Length);
        for (int i = 0; i < arr.Length; i++)
        {
            target.Add(ConvertValue<T>(arr[i]));
        }
        return target;
    }

    /// <summary>
    /// Converts a UNO value to DateTime.
    /// LibreOffice/Calc passes dates as OLE Automation dates (doubles).
    /// </summary>
    public static DateTime ToDateTime(object? value)
    {
        if (value == null)
            throw MarshalException("DateTime", value);
        if (value is DateTime dt)
            return dt;
        if (value is double d)
            return DateTime.FromOADate(d);
        if (value is IConvertible)
            return Convert.ToDateTime(value);
        throw MarshalException("DateTime", value);
    }

    #region Optional Parameter Unwrappers

    public static bool UnwrapOptionalBool(object? value, bool defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is bool b)
            return b;
        if (value is IConvertible)
            return Convert.ToBoolean(value);
        throw MarshalException("bool", value);
    }

    public static byte UnwrapOptionalByte(object? value, byte defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is byte b)
            return b;
        if (value is IConvertible)
            return Convert.ToByte(value);
        throw MarshalException("byte", value);
    }

    public static sbyte UnwrapOptionalSByte(object? value, sbyte defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is sbyte sb)
            return sb;
        if (value is IConvertible)
            return Convert.ToSByte(value);
        throw MarshalException("sbyte", value);
    }

    public static short UnwrapOptionalShort(object? value, short defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is short s)
            return s;
        if (value is IConvertible)
            return Convert.ToInt16(value);
        throw MarshalException("short", value);
    }

    public static int UnwrapOptionalInt(object? value, int defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is int i)
            return i;
        if (value is IConvertible)
            return Convert.ToInt32(value);
        throw MarshalException("int", value);
    }

    public static long UnwrapOptionalLong(object? value, long defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is long l)
            return l;
        if (value is IConvertible)
            return Convert.ToInt64(value);
        throw MarshalException("long", value);
    }

    public static float UnwrapOptionalFloat(object? value, float defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is float f)
            return f;
        if (value is IConvertible)
            return Convert.ToSingle(value);
        throw MarshalException("float", value);
    }

    public static double UnwrapOptionalDouble(object? value, double defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is double d)
            return d;
        if (value is IConvertible)
            return Convert.ToDouble(value);
        throw MarshalException("double", value);
    }

    public static char UnwrapOptionalChar(object? value, char defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is char c)
            return c;
        if (value is string s && s.Length > 0)
            return s[0];
        if (value is IConvertible)
            return Convert.ToChar(value);
        throw MarshalException("char", value);
    }

    public static string UnwrapOptionalString(object? value, string defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue ?? string.Empty;
        if (value is string s)
            return s;
        // Convert any value to string representation
        return value?.ToString() ?? defaultValue ?? string.Empty;
    }

    public static object UnwrapOptionalObject(object? value, object defaultValue)
    {
        return value is null || IsEmpty(value) ? defaultValue : value;
    }

    public static DateTime UnwrapOptionalDateTime(object? value, DateTime defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is DateTime dt)
            return dt;
        if (value is double  d)
            return DateTime.FromOADate(d);
        // TODO: Do we support string parsing
        throw MarshalException("DateTime", value);
    }

    public static T[] UnwrapOptional1DArray<T>(object? value, T[] defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is object[] arr)
            return To1DArray<T>(arr);
        throw MarshalException("array", value);
    }

    public static T[,] UnwrapOptional2DArray<T>(object? value, T[,] defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue ?? new T[0, 0];
        if (value is object[][] jagged)
            return To2DArray<T>(jagged);
        throw MarshalException("2D array", value);
    }

    public static CalcRange UnwrapOptionalCalcRange(object? value, CalcRange defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is object[][] jagged)
            return ToCalcRange(jagged);
        throw MarshalException("CalcRange", value);
    }

    public static List<T> UnwrapOptionalList<T>(object? value, List<T> defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        if (value is object[] arr)
            return ToList<T>(arr);
        throw MarshalException("List", value);
    }

    public static bool? UnwrapNullableBool(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is bool b)
            return b;
        if (value is IConvertible)
            return Convert.ToBoolean(value);
        throw MarshalException("bool?", value);
    }

    public static byte? UnwrapNullableByte(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is byte b)
            return b;
        if (value is IConvertible)
            return Convert.ToByte(value);
        throw MarshalException("byte?", value);
    }

    public static sbyte? UnwrapNullableSByte(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is sbyte sb)
            return sb;
        if (value is IConvertible)
            return Convert.ToSByte(value);
        throw MarshalException("sbyte?", value);
    }

    public static short? UnwrapNullableShort(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is short s)
            return s;
        if (value is IConvertible)
            return Convert.ToInt16(value);
        throw MarshalException("short?", value);
    }

    public static int? UnwrapNullableInt(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is int i)
            return i;
        if (value is IConvertible)
            return Convert.ToInt32(value);
        throw MarshalException("int?", value);
    }

    public static long? UnwrapNullableLong(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is long l)
            return l;
        if (value is IConvertible)
            return Convert.ToInt64(value);
        throw MarshalException("long?", value);
    }

    public static float? UnwrapNullableFloat(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is float f)
            return f;
        if (value is IConvertible)
            return Convert.ToSingle(value);
        throw MarshalException("float?", value);
    }

    public static double? UnwrapNullableDouble(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is double d)
            return d;
        if (value is IConvertible)
            return Convert.ToDouble(value);
        throw MarshalException("double?", value);
    }

    public static char? UnwrapNullableChar(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is char c)
            return c;
        if (value is string s && s.Length > 0)
            return s[0];
        if (value is IConvertible)
            return Convert.ToChar(value);
        throw MarshalException("char?", value);
    }

    public static string? UnwrapNullableString(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is string s)
            return s;
        // Convert any value to string representation
        return value?.ToString();
    }

    public static object? UnwrapNullableObject(object? value)
    {
        return IsEmpty(value) ? null : value;
    }

    public static DateTime? UnwrapNullableDateTime(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is DateTime dt)
            return dt;
        if (value is double d)
            return DateTime.FromOADate(d);
        throw MarshalException("DateTime?", value);
    }

    public static T[]? UnwrapNullable1DArray<T>(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is object[] arr)
            return To1DArray<T>(arr);
        throw MarshalException("array", value);

    }

    public static T[,]? UnwrapNullable2DArray<T>(object? value)
    {
         if (IsEmpty(value))
            return null;
        if (value is object[][] jagged)
            return To2DArray<T>(jagged);
        throw MarshalException("2D array", value);

    }

    public static CalcRange? UnwrapNullableCalcRange(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is object[][] jagged)
            return ToCalcRange(jagged);
        throw MarshalException("CalcRange", value);
    }

    public static List<T>? UnwrapNullableList<T>(object? value)
    {
        if (IsEmpty(value))
            return null;
        if (value is object[] arr)
            return ToList<T>(arr);
        throw MarshalException("List", value);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// Checks if a value represents an empty/missing UNO value.
    /// In UNO, empty cells can be null, DBNull, or a special Empty type.
    /// </summary>
    private static bool IsEmpty(object? value)
    {
        if (value == null)
            return true;
        if (value == DBNull.Value)
            return true;
        if (value is System.Reflection.Missing)
            return true;
        // Check for UNO's "void" or empty marker (type name varies by binding)
        var typeName = value.GetType().Name;
        if (typeName is "Empty" or "Void" or "Missing")
            return true;
        // Empty string is considered empty for optional parameters
        if (value is string s && s.Length == 0)
            return true;
        // Array of one cell that is empty (e.g. A1:A1 with empty cell)
        if (value is object[] arr && ( arr.Length == 0 || (arr.Length == 1 && IsEmpty(arr[0])) ) )
            return true;
        return false;
    }

    /// <summary>
    /// Converts a UNO value to the specified type.
    /// Handles common UNO conventions like doubles for all numeric types.
    /// </summary>
    private static T ConvertValue<T>(object? value)
    {
        if (value == null || value == DBNull.Value)
            return default!;

        var targetType = typeof(T);

        // Direct type match
        if (value is T typed)
            return typed;

        // Handle nullable types
        var underlyingType = Nullable.GetUnderlyingType(targetType);
        if (underlyingType != null)
        {
            targetType = underlyingType;
        }

        // Use Convert for IConvertible types
        if (value is IConvertible)
        {
            try
            {
                return (T)Convert.ChangeType(value, targetType);
            }
            catch (Exception ex)
            {
                throw new InvalidCastException(
                    $"Cannot convert value '{value}' of type '{value.GetType().Name}' to '{typeof(T).Name}'.", ex);
            }
        }

        throw new InvalidCastException(
            $"Cannot convert value of type '{value.GetType().Name}' to '{typeof(T).Name}'.");
    }

    /// <summary>
    /// Gets the maximum column count in a jagged array (handles ragged arrays).
    /// </summary>
    private static int GetMaxColumns(object[][] source)
    {
        int max = 0;
        for (int i = 0; i < source.Length; i++)
        {
            if (source[i] != null && source[i].Length > max)
                max = source[i].Length;
        }
        return max;
    }

    /// <summary>
    /// Creates a consistent marshal exception with type information.
    /// </summary>
    private static InvalidCastException MarshalException(string targetType, object? value)
    {
        var sourceType = value?.GetType().Name ?? "null";
        return new InvalidCastException(
            $"Cannot convert value of type '{sourceType}' to {targetType}.");
    }

    #endregion
}
