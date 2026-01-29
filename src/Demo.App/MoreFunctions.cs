using CalcDNA.Attributes;
using CalcDNA.Runtime;
using System;
using System.Collections.Generic;
using System.Linq;

namespace Demo.App;

[CalcAddIn]
public static partial class MoreFunctions
{
    [CalcFunction(Description = "Sum a range with a list of numbers")]
    public static double SumRangeAndList(
        [CalcParameter(Description = "A range of numbers")] CalcRange range,
        [CalcParameter(Description = "A list of numbers")] List<double> numbers)
    {
        double sum = 0;
        foreach (var val in range.Values())
        {
            if (val is IConvertible)
                sum += Convert.ToDouble(val);
        }
        foreach (var num in numbers)
        {
            sum += num;
        }
        return sum;
    }

    [CalcFunction(Description = "Filters a list of numbers to keep only those greater than a threshold")]
    public static List<double> FilterGreater(List<double> values, double threshold)
    {
        return values.Where(v => v > threshold).ToList();
    }

    [CalcFunction(Description = "Advanced string joining from a list")]
    public static string JoinAll(List<string> items, string separator = ", ")
    {
        return string.Join(separator, items);
    }

    [CalcFunction(Description = "Transposes a 2D matrix of doubles")]
    public static double[,] Transpose(double[,] matrix)
    {
        int rows = matrix.GetLength(0);
        int cols = matrix.GetLength(1);
        var result = new double[cols, rows];
        for (int i = 0; i < rows; i++)
            for (int j = 0; j < cols; j++)
                result[j, i] = matrix[i, j];
        return result;
    }

    [CalcFunction(Description = "Generates a multiplication table")]
    public static double[,] MultiTable(int size)
    {
        var result = new double[size, size];
        for (int i = 0; i < size; i++)
            for (int j = 0; j < size; j++)
                result[i, j] = (i + 1) * (j + 1);
        return result;
    }
}