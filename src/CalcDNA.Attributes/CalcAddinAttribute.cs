using System;

namespace CalcDNA.Attributes
{
    /// <summary>
    /// Marks a static class as containing LibreOffice Calc functions.
    /// Classes with this attribute will be scanned for methods marked with [CalcFunction].
    /// </summary>
    /// <remarks>
    /// This is a marker attribute - extension-level metadata (version, publisher, etc.)
    /// should be specified using [assembly: ExtensionMetadata(...)] instead.
    /// </remarks>
    /// <example>
    /// [CalcAddIn]
    /// public static partial class MyFunctions
    /// {
    ///     [CalcFunction(Description = "Adds two numbers")]
    ///     public static double Add(double a, double b) => a + b;
    /// }
    /// </example>
    [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    public sealed class CalcAddInAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="CalcAddInAttribute"/> class.
        /// </summary>
        public CalcAddInAttribute()
        {
        }
    }
}
