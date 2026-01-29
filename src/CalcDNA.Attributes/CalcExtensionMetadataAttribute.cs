using System;

namespace CalcDNA.Attributes
{
    /// <summary>
    /// Specifies metadata for the LibreOffice Calc extension package.
    /// Apply this attribute at the assembly level to configure version, publisher, and other package information.
    /// </summary>
    /// <example>
    /// [assembly: CalcExtensionMetadata(
    ///     Version = "1.0.0",
    ///     Publisher = "Your Name",
    ///     Description = "My awesome Calc add-in"
    /// )]
    /// </example>
    [AttributeUsage(AttributeTargets.Assembly, AllowMultiple = false)]
    public class CalcExtensionMetadataAttribute : Attribute
    {
        /// <summary>
        /// Extension version (e.g., "1.0.0", "2.1.5-beta").
        /// This version is used for upgrade detection in LibreOffice.
        /// </summary>
        public string Version { get; set; }

        /// <summary>
        /// Display name for the extension shown in LibreOffice Extension Manager.
        /// If not specified, the assembly name is used.
        /// </summary>
        public string DisplayName { get; set; }

        /// <summary>
        /// Publisher or author name shown in Extension Manager.
        /// </summary>
        public string Publisher { get; set; }

        /// <summary>
        /// Short description of the extension's purpose and functionality.
        /// </summary>
        public string Description { get; set; }

        /// <summary>
        /// Unique identifier for the extension (e.g., "com.example.myaddin").
        /// If not specified, one is generated from the assembly name.
        /// Two extensions with the same identifier cannot be installed simultaneously.
        /// </summary>
        public string Identifier { get; set; }

        /// <summary>
        /// Minimum LibreOffice version required (e.g., "7.0", "24.2").
        /// Default is "7.0".
        /// </summary>
        public string MinLibreOfficeVersion { get; set; }

        /// <summary>
        /// Maximum LibreOffice version supported (optional).
        /// Leave empty for no maximum version restriction.
        /// </summary>
        public string MaxLibreOfficeVersion { get; set; }

        /// <summary>
        /// URL to the update.xml file for automatic update checking.
        /// When provided, users can check for updates via Extension Manager.
        /// Example: "https://example.com/updates/MyAddIn.update.xml"
        /// </summary>
        public string UpdateUrl { get; set; }

        /// <summary>
        /// URL to release notes for this version.
        /// Example: "https://example.com/releases/v1.0.0.html"
        /// </summary>
        public string ReleaseNotesUrl { get; set; }

        /// <summary>
        /// URL to the extension icon (optional).
        /// </summary>
        public string IconUrl { get; set; }
    }
}
