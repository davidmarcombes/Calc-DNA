using System;

namespace CalcDNA.Attributes
{
    [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    public sealed class CalcAddInAttribute : Attribute
    {
        // Make these settable so attributes can be used with named parameters:
        // [CalcAddIn(Name="...", Description="...")]
        public string Name { get; set; }
        public string Description { get; set; }

        public CalcAddInAttribute() { }
        public CalcAddInAttribute(string name)
        {
            Name = name;
        }
    }
}