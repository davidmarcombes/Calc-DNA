using CalcDNA.Attributes;
using CalcDNA.Runtime;

namespace Demo.App;

[CalcAddIn(Name = "Demo.App.More", Description = "More Demo Functions")]
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
}