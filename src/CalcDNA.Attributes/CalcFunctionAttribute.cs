using System;

namespace CalcDNA.Attributes
{
    /// <summary>
    /// Identifies a method as a custom calculation function and provides metadata for function browsers.
    /// </summary>
    [AttributeUsage(AttributeTargets.Method, Inherited = false, AllowMultiple = false)]
    public sealed class CalcFunctionAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets the name of the function as it will appear in the formula engine.
        /// </summary>
        /// <value>The function's identifier (e.g., "SUM", "CUSTOM_LOG").</value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets a description explaining what the function does and what its arguments are.
        /// </summary>
        /// <value>A user-friendly description of the method's purpose.</value>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets the group or category this function belongs to for organization in the UI.
        /// </summary>
        /// <value>The category name. The default is "Add-In".</value>
        public string Category { get; set; } = "Add-In";

        /// <summary>
        /// Initializes a new instance of the <see cref="CalcFunctionAttribute"/> class.
        /// </summary>
        public CalcFunctionAttribute() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="CalcFunctionAttribute"/> class with a specified name.
        /// </summary>
        /// <param name="name">The name of the function to be used in calculations.</param>
        public CalcFunctionAttribute(string name)
        {
            Name = name;
        }
    }
}