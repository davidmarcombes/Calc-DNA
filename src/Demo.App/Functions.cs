using CalcDNA.Attributes;


namespace Demo.App;


[CalcAddIn(Name = "Demo.App", Description = "Demo Add-In")]
static public partial class Functions
{

    [CalcFunction(Name="Add", Description="Add two numbers")]
    public static double Add(
        [CalcParameter(Name="a", Description="First number")] double a,
        [CalcParameter(Name="b", Description="Second number")] double b = 0.0)
    {
        return a + b;
    }
}
