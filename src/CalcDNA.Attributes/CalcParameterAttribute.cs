using System;


namespace CalcDNA.Attributes
{
    /// <summary>
    /// Provides metadata for a specific parameter within a calculator function, 
    /// such as its display name and whether it is optional.
    /// </summary>
    [AttributeUsage(AttributeTargets.Parameter, Inherited = false, AllowMultiple = false)]
    public sealed class CalcParameterAttribute : Attribute
    {
        /// <summary>
        /// Gets or sets the name of the parameter as it should appear in the function's signature.
        /// </summary>
        /// <value>The parameter name (e.g., "Radius", "Value", or "X").</value>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets a description of what this specific parameter represents.
        /// </summary>
        /// <value>A helpful tooltip describing the expected input or units.</value>
        public string Description { get; set; }

        /// <summary>
        /// Gets or sets a value indicating whether this parameter is optional for the function.
        /// </summary>
        /// <value><c>true</c> if the parameter can be omitted; otherwise, <c>false</c>.</value>
        public bool Optional { get; set; }

        /// <summary>
        /// Initializes a new instance of the <see cref="CalcParameterAttribute"/> class.
        /// </summary>
        public CalcParameterAttribute() { }

        /// <summary>
        /// Initializes a new instance of the <see cref="CalcParameterAttribute"/> class with a specified name.
        /// </summary>
        /// <param name="name">The display name of the parameter.</param>
        public CalcParameterAttribute(string name)
        {
            Name = name;
        }
    }
}