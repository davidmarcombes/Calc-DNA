using CalcDNA.Attributes;


namespace Demo.App;


[CalcAddIn]
static public partial class Functions
{

    [CalcFunction(Name="Add", Description="Add two numbers")]
    public static double Add(
        [CalcParameter(Name="a", Description="First number")] double a,
        [CalcParameter(Name="b", Description="Second number")] double b = 0.0)
    {
        return a + b;
    }

    [CalcFunction(Description = "Concatenates two strings with a separator")]
    public static string Concat(
        string first, 
        string second, 
        string separator = " ") => $"{first}{separator}{second}";

    [CalcFunction(Description = "Returns true if the number is even")]
    public static bool IsEven(int value) => value % 2 == 0;

    [CalcFunction(Description = "Adds a number of days to a date")]
    public static DateTime AddDays(
        [CalcParameter(Description = "Starting date")] DateTime start, 
        [CalcParameter(Description = "Number of days to add")] int days) 
        => start.AddDays(days);

    [CalcFunction(Description = "Returns the current system time")]
    public static DateTime Now() => DateTime.Now;
}
