using System;
using System.Collections.Generic;
using System.Linq;

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

    /// <summary>
    /// Converts a .NET enumerable to a UNO object array.
    /// </summary>
    public static object[] ToUno1DArray<T>(IEnumerable<T>? source)
    {
        if (source == null)
            return Array.Empty<object>();

        if (source is object[] objArr)
            return objArr;

        return source.Select(v => (object?)v ?? DBNull.Value).ToArray()!;
    }

    /// <summary>
    /// Converts a .NET 2D array to a UNO jagged array (object[][]).
    /// </summary>
    public static object[][] ToUno2DArray<T>(T[,]? source)
    {
        if (source == null)
            return Array.Empty<object[]>();

        int rows = source.GetLength(0);
        int cols = source.GetLength(1);
        var target = new object[rows][];

        for (int i = 0; i < rows; i++)
        {
            target[i] = new object[cols];
            for (int j = 0; j < cols; j++)
            {
                var val = source[i, j];
                target[i][j] = (object?)val ?? DBNull.Value;
            }
        }
        return target;
    }

    /// <summary>
    /// Converts a CalcRange back to a UNO jagged array.
    /// </summary>
    public static object[][] ToUno2DArray(CalcRange? source)
    {
        if (source == null)
            return Array.Empty<object[]>();
        
        return source.ToJaggedArray()!;
    }

    #region Optional Parameter Unwrappers

    public static bool UnwrapOptionalBool(object? value, bool defaultValue) => UnwrapOptional(value, defaultValue);
    public static byte UnwrapOptionalByte(object? value, byte defaultValue) => UnwrapOptional(value, defaultValue);
    public static sbyte UnwrapOptionalSByte(object? value, sbyte defaultValue) => UnwrapOptional(value, defaultValue);
    public static short UnwrapOptionalShort(object? value, short defaultValue) => UnwrapOptional(value, defaultValue);
    public static int UnwrapOptionalInt(object? value, int defaultValue) => UnwrapOptional(value, defaultValue);
    public static long UnwrapOptionalLong(object? value, long defaultValue) => UnwrapOptional(value, defaultValue);
    public static float UnwrapOptionalFloat(object? value, float defaultValue) => UnwrapOptional(value, defaultValue);
    public static double UnwrapOptionalDouble(object? value, double defaultValue) => UnwrapOptional(value, defaultValue);
    public static char UnwrapOptionalChar(object? value, char defaultValue) => UnwrapOptional(value, defaultValue);
    public static string UnwrapOptionalString(object? value, string defaultValue) => UnwrapOptional(value, defaultValue ?? string.Empty);
    public static object UnwrapOptionalObject(object? value, object defaultValue) => UnwrapOptional(value, defaultValue);
    public static DateTime UnwrapOptionalDateTime(object? value, DateTime defaultValue) => UnwrapOptional(value, defaultValue);

    public static T[] UnwrapOptional1DArray<T>(object? value, T[] defaultValue)
    {
        if (IsEmpty(value)) return defaultValue;
        if (value is object[] arr) return To1DArray<T>(arr);
        throw MarshalException("array", value);
    }

    public static T[,] UnwrapOptional2DArray<T>(object? value, T[,] defaultValue)
    {
        if (IsEmpty(value)) return defaultValue ?? new T[0, 0];
        if (value is object[][] jagged) return To2DArray<T>(jagged);
        throw MarshalException("2D array", value);
    }

    public static CalcRange UnwrapOptionalCalcRange(object? value, CalcRange defaultValue)
    {
        if (IsEmpty(value)) return defaultValue;
        if (value is object[][] jagged) return ToCalcRange(jagged);
        throw MarshalException("CalcRange", value);
    }

    public static List<T> UnwrapOptionalList<T>(object? value, List<T> defaultValue)
    {
        if (IsEmpty(value)) return defaultValue;
        if (value is object[] arr) return ToList<T>(arr);
        throw MarshalException("List", value);
    }

    public static bool? UnwrapNullableBool(object? value) => UnwrapNullable<bool>(value);
    public static byte? UnwrapNullableByte(object? value) => UnwrapNullable<byte>(value);
    public static sbyte? UnwrapNullableSByte(object? value) => UnwrapNullable<sbyte>(value);
    public static short? UnwrapNullableShort(object? value) => UnwrapNullable<short>(value);
    public static int? UnwrapNullableInt(object? value) => UnwrapNullable<int>(value);
    public static long? UnwrapNullableLong(object? value) => UnwrapNullable<long>(value);
    public static float? UnwrapNullableFloat(object? value) => UnwrapNullable<float>(value);
    public static double? UnwrapNullableDouble(object? value) => UnwrapNullable<double>(value);
    public static char? UnwrapNullableChar(object? value) => UnwrapNullable<char>(value);
    public static string? UnwrapNullableString(object? value) => IsEmpty(value) ? null : ConvertValue<string>(value);
    public static object? UnwrapNullableObject(object? value) => IsEmpty(value) ? null : value;
    public static DateTime? UnwrapNullableDateTime(object? value) => UnwrapNullable<DateTime>(value);

    public static T[]? UnwrapNullable1DArray<T>(object? value)
    {
        if (IsEmpty(value)) return null;
        if (value is object[] arr) return To1DArray<T>(arr);
        throw MarshalException("array", value);
    }

    public static T[,]? UnwrapNullable2DArray<T>(object? value)
    {
        if (IsEmpty(value)) return null;
        if (value is object[][] jagged) return To2DArray<T>(jagged);
        throw MarshalException("2D array", value);
    }

    public static CalcRange? UnwrapNullableCalcRange(object? value)
    {
        if (IsEmpty(value)) return null;
        if (value is object[][] jagged) return ToCalcRange(jagged);
        throw MarshalException("CalcRange", value);
    }

    public static List<T>? UnwrapNullableList<T>(object? value)
    {
        if (IsEmpty(value)) return null;
        if (value is object[] arr) return ToList<T>(arr);
        throw MarshalException("List", value);
    }

    /// <summary>
    /// Generic unwrap for optional parameters.
    /// </summary>
    public static T UnwrapOptional<T>(object? value, T defaultValue)
    {
        if (IsEmpty(value))
            return defaultValue;
        return ConvertValue<T>(value);
    }

    /// <summary>
    /// Generic unwrap for nullable parameters.
    /// </summary>
    public static T? UnwrapNullable<T>(object? value) where T : struct
    {
        if (IsEmpty(value))
            return null;
        return ConvertValue<T>(value);
    }

    #endregion

    #region Helper Methods

    /// <summary>
    /// Checks if a value represents an empty/missing UNO value.
    /// In UNO, empty cells can be null, DBNull, or a special Empty type.
    /// </summary>
    private static bool IsEmpty(object? value)
    {
        if (value is null or DBNull or System.Reflection.Missing)
            return true;

        // Common types are not empty
        if (value is double or int or bool or DateTime or float or long)
            return false;

        // Check for UNO's "void" or empty marker (type name varies by binding)
        var typeName = value.GetType().Name;
        if (typeName is "Empty" or "Void" or "Missing")
            return true;

        // Empty string is considered empty for optional parameters
        if (value is string s)
             return s.Length == 0;

        // Array handling: check if the range is empty (e.g. A1:A1 with empty cell)
        if (value is object[] arr)
             return arr.Length == 0 || (arr.Length == 1 && IsEmpty(arr[0]));

        if (value is object[][] jagged)
             return jagged.Length == 0 || (jagged.Length == 1 && IsEmpty(jagged[0]));

        return false;
    }

    /// <summary>
    /// Converts a UNO value to the specified type.
    /// Handles common UNO conventions like doubles for all numeric types.
    /// </summary>
    public static T ConvertValue<T>(object? value)
    {
        if (value == null || value == DBNull.Value)
            return default!;

        var targetType = typeof(T);

        // Direct type match
        if (value is T typed)
            return typed;

        // Handle string special case
        if (targetType == typeof(string))
            return (T)(object)value.ToString()!;

        // Handle char from string
        if (targetType == typeof(char) && value is string s && s.Length > 0)
            return (T)(object)s[0];

        // Handle DateTime from OLE date (double)
        if (targetType == typeof(DateTime) && value is double d)
            return (T)(object)DateTime.FromOADate(d);

        // Handle nullable types
        var underlyingType = Nullable.GetUnderlyingType(targetType);
        if (underlyingType != null)
        {
            // If we are here, T is a Nullable<U>, and value is not null.
            // We need to convert value to U.
            // We can call ConvertValue<underlyingType> recursively if we want,
            // but for performance we just switch targetType.
            targetType = underlyingType;
            if (value is T typedNullable) return typedNullable;
        }

        // Use IConvertible for common types
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
