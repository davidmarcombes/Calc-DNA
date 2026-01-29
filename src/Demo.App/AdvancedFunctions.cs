using CalcDNA.Attributes;
using CalcDNA.Runtime;
using System.Linq;

namespace Demo.App;

[CalcAddIn]
public static partial class AdvancedFunctions
{
    [CalcFunction(Description = "Calculates the average of non-empty cells in a range")]
    public static double AverageNonEmpty(CalcRange range)
    {
        var values = range.Values().Cast<object>().Where(v => v is double).Cast<double>().ToList();
        return values.Count > 0 ? values.Average() : 0;
    }

    [CalcFunction(Description = "Returns the maximum value in a range, or a default if empty")]
    public static double MaxWithDefault(CalcRange range, double defaultValue = -1.0)
    {
        var values = range.Values().Cast<object>().Where(v => v is double).Cast<double>().ToList();
        return values.Count > 0 ? values.Max() : defaultValue;
    }

    [CalcFunction(Description = "Extracts unique strings from a range")]
    public static List<string> UniqueStrings(CalcRange range)
    {
        return range.Values().Cast<object>()
            .Where(v => v is string)
            .Cast<string>()
            .Where(s => !string.IsNullOrEmpty(s))
            .Distinct()
            .ToList();
    }

    [CalcFunction(Description = "Sums only non-null values from a list of nullable doubles")]
    public static double SumNonNull(List<double?> values)
    {
        return values.Where(v => v.HasValue).Sum(v => v!.Value);
    }

    [CalcFunction(Description = "Returns the range bounds as a 1D array [rows, cols]")]
    public static double[] GetRangeSize(CalcRange range)
    {
        return new double[] { range.Rows, range.Columns };
    }

    [CalcFunction(Description = "Returns the range as a string matrix with headers")]
    public static string[,] ToLabeledMatrix(CalcRange range)
    {
        int rows = range.Rows;
        int cols = range.Columns;
        var result = new string[rows + 1, cols + 1];
        
        for (int i = 0; i <= rows; i++) result[i, 0] = i == 0 ? "IDX" : $"R{i}";
        for (int j = 1; j <= cols; j++) result[0, j] = $"C{j}";
        
        var data = range.ToJaggedArray();
        for (int i = 0; i < rows; i++)
        {
            for (int j = 0; j < cols; j++)
            {
                result[i + 1, j + 1] = data[i][j]?.ToString() ?? "";
            }
        }
        return result;
    }
}
