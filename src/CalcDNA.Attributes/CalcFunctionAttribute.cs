using System;

namespace CalcDNA.Attributes
{
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    public sealed class CalcFunctionAttribute : Attribute
    {
        // Support named parameters like: [CalcFunction(Name="Add", Description="...")]
        public string Name { get; set; }
        public string Description { get; set; }
        public string Category { get; set; } = "Add-In";

        public CalcFunctionAttribute() { }
        public CalcFunctionAttribute(string name)
        {
            Name = name;
        }
    }
}