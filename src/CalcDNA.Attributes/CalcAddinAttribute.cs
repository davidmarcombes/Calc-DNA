using System;

namespace CalcDNA.Attributes
{
    /// <summary>
    /// Specifies metadata for a calculator add-in, allowing the system to identify 
    /// and describe the component at runtime.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, Inherited = false, AllowMultiple = false)]
    public sealed class CalcAddInAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets the display name of the add-in.
        /// </summary>
        /// <value>The friendly name of the add-in as shown in the UI.</value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets a brief description of what the add-in does.
        /// </summary>
        /// <value>A string describing the mathematical functions or purpose of the add-in.</value>
        public string Description { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CalcAddInAttribute"/> class.
        /// </summary>
        public CalcAddInAttribute() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="CalcAddInAttribute"/> class with a specific name.
        /// </summary>
        /// <param name="name">The name to assign to the add-in.</param>
        public CalcAddInAttribute(string name)
        {
            Name = name;
        }
    }
}