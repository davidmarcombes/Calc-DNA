using System;

namespace CalcDNA.Attributes
{
    [AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
    public sealed class CalcParameterAttribute : Attribute
    {
        // Support named parameters like: [CalcParameter(Name="a", Description="...")]
        public string Name { get; set; }
        public string Description { get; set; }

        // Support optional parameters like: [CalcParameter(Optional=true)]
        public bool Optional { get; set; }

        public CalcParameterAttribute() { }
        public CalcParameterAttribute(string name)
        {
            Name = name;
        }
    }
}