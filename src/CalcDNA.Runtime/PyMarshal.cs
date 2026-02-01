using System;
using System.Collections.Generic;
using System.Linq;

namespace CalcDNA.Runtime;

/// <summary>
/// Provides marshaling utilities for converting between Python (via pythonnet) and .NET types.
/// Unlike UnoMarshal which assumes object[]/object[][] inputs from the .NET UNO bridge,
/// PyMarshal accepts object and iterates IEnumerable to handle Python tuples/lists
/// transparently. pythonnet wraps Python sequences as IEnumerable in .NET.
/// </summary>
public static class PyMarshal
{
    /// <summary>
    /// Converts any iterable (Python list/tuple, .NET array) to a typed 1D array.
    /// </summary>
    public static T[] To1DArray<T>(object? source)
    {
        if (source == null)
            return Array.Empty<T>();
        if (source is object[] arr)
            return UnoMarshal.To1DArray<T>(arr);

        return IterateFlat(source).Select(UnoMarshal.ConvertValue<T>).ToArray();
    }

    /// <summary>
    /// Converts nested iterables (Python tuple of tuples, list of lists, etc.)
    /// to a typed rectangular 2D array.
    /// </summary>
    public static T[,] To2DArray<T>(object? source)
    {
        if (source == null)
            return new T[0, 0];
        if (source is object[][] jagged)
            return UnoMarshal.To2DArray<T>(jagged);

        return UnoMarshal.To2DArray<T>(ToJaggedArray(source));
    }

    /// <summary>
    /// Converts nested iterables to a CalcRange.
    /// </summary>
    public static CalcRange ToCalcRange(object? source)
    {
        if (source == null)
            return new CalcRange(Array.Empty<object[]>());
        if (source is object[][] jagged)
            return UnoMarshal.ToCalcRange(jagged);

        return new CalcRange(ToJaggedArray(source));
    }

    /// <summary>
    /// Converts any iterable to a typed List.
    /// </summary>
    public static List<T> ToList<T>(object? source)
    {
        if (source == null)
            return new List<T>();
        if (source is object[] arr)
            return UnoMarshal.ToList<T>(arr);

        return IterateFlat(source).Select(UnoMarshal.ConvertValue<T>).ToList();
    }

    /// <summary>
    /// Converts a .NET array or collection to a Python-friendly object array.
    /// Nulls are preserved (not replaced with DBNull) for Python compatibility.
    /// </summary>
    public static object[] ToPy1DArray<T>(IEnumerable<T>? source)
    {
        if (source == null)
            return Array.Empty<object>();

        return source.Select(v => (object?)v).ToArray()!;
    }

    /// <summary>
    /// Converts a .NET rectangular 2D array to nested object arrays for Python.
    /// </summary>
    public static object[][] ToPy2DArray<T>(T[,]? source)
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
                target[i][j] = (object?)source[i, j]!;
        }
        return target;
    }

    /// <summary>
    /// Converts a jagged 2D array to nested object arrays for Python.
    /// </summary>
    public static object[][] ToPy2DArray<T>(T[][]? source)
    {
        if (source == null)
            return Array.Empty<object[]>();

        return source.Select(row => row.Select(v => (object)v!).ToArray()).ToArray();
    }

    /// <summary>
    /// Converts a CalcRange back to nested object arrays for Python.
    /// </summary>
    public static object[][] ToPy2DArray(CalcRange? source)
    {
        if (source == null)
            return Array.Empty<object[]>();

        return source.ToJaggedArray()!;
    }

    #region Optional/Nullable Unwrappers for Complex Types

    public static T[] UnwrapOptional1DArray<T>(object? value, T[] defaultValue)
    {
        if (UnoMarshal.IsEmpty(value)) return defaultValue;
        return To1DArray<T>(value);
    }

    public static T[,] UnwrapOptional2DArray<T>(object? value, T[,] defaultValue)
    {
        if (UnoMarshal.IsEmpty(value)) return defaultValue ?? new T[0, 0];
        return To2DArray<T>(value);
    }

    public static CalcRange UnwrapOptionalCalcRange(object? value, CalcRange defaultValue)
    {
        if (UnoMarshal.IsEmpty(value)) return defaultValue;
        return ToCalcRange(value);
    }

    public static List<T> UnwrapOptionalList<T>(object? value, List<T> defaultValue)
    {
        if (UnoMarshal.IsEmpty(value)) return defaultValue;
        return ToList<T>(value);
    }

    public static T[]? UnwrapNullable1DArray<T>(object? value)
    {
        if (UnoMarshal.IsEmpty(value)) return null;
        return To1DArray<T>(value);
    }

    public static T[,]? UnwrapNullable2DArray<T>(object? value)
    {
        if (UnoMarshal.IsEmpty(value)) return null;
        return To2DArray<T>(value);
    }

    public static CalcRange? UnwrapNullableCalcRange(object? value)
    {
        if (UnoMarshal.IsEmpty(value)) return null;
        return ToCalcRange(value);
    }

    public static List<T>? UnwrapNullableList<T>(object? value)
    {
        if (UnoMarshal.IsEmpty(value)) return null;
        return ToList<T>(value);
    }

    #endregion

    #region Helpers

    /// <summary>
    /// Converts any nested iterable to object[][] (jagged array).
    /// Handles Python tuples, lists, .NET arrays, and any IEnumerable.
    /// </summary>
    private static object[][] ToJaggedArray(object source)
    {
        var rows = new List<object[]>();

        foreach (var row in IterateFlat(source))
        {
            if (row is object[] objRow)
                rows.Add(objRow);
            else if (row is System.Collections.IEnumerable rowEnum && row is not string)
                rows.Add(rowEnum.Cast<object>().ToArray());
            else
                rows.Add(new object[] { row! });
        }

        return rows.ToArray();
    }

    /// <summary>
    /// Iterates any IEnumerable (generic or non-generic) as object sequence.
    /// </summary>
    private static IEnumerable<object> IterateFlat(object source)
    {
        if (source is IEnumerable<object> generic)
            return generic;
        if (source is System.Collections.IEnumerable nonGeneric)
            return nonGeneric.Cast<object>();

        throw new InvalidCastException(
            $"Cannot iterate value of type '{source.GetType().Name}'.");
    }

    #endregion
}
