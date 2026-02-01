using System.Xml.Linq;

namespace CalcDNA.CLI;

/// <summary>
/// Generates the description.xml file for LibreOffice extension packages.
/// This file contains extension metadata like version, display name, and requirements.
/// </summary>
internal static class DescriptionGenerator
{
    private static readonly XNamespace DescNs = "http://openoffice.org/extensions/description/2006";
    private static readonly XNamespace Xlink = "http://www.w3.org/1999/xlink";

    /// <summary>
    /// Extension metadata configuration.
    /// </summary>
    public class ExtensionInfo
    {
        public string Identifier { get; set; } = "";
        public string Version { get; set; } = "1.0.0";
        public string DisplayName { get; set; } = "";
        public string Publisher { get; set; } = "";
        public string Description { get; set; } = "";
        public string MinLibreOfficeVersion { get; set; } = "4.0";
        public string MaxLibreOfficeVersion { get; set; } = "";

        /// <summary>
        /// URL to the update information XML file for automatic update checking.
        /// When provided, LibreOffice can check for updates automatically.
        /// Example: "https://example.com/updates/extension.update.xml"
        /// </summary>
        public string? UpdateUrl { get; set; } = null;

        /// <summary>
        /// Additional mirror URLs for update information (for redundancy).
        /// </summary>
        public List<string> UpdateMirrorUrls { get; set; } = new();

        /// <summary>
        /// URL to the release notes for this version.
        /// </summary>
        public string? ReleaseNotesUrl { get; set; } = null;

        /// <summary>
        /// URL to the extension icon (optional).
        /// </summary>
        public string? IconUrl { get; set; } = null;
    }

    /// <summary>
    /// Generates the description.xml content for an OXT package.
    /// </summary>
    /// <param name="info">Extension metadata</param>
    /// <param name="logger">Logger instance for output</param>
    /// <returns>XML content for description.xml</returns>
    public static string BuildDescription(ExtensionInfo info, Logger logger)
    {
        logger.Debug("Building description.xml...", true);

        // Validate required fields
        if (string.IsNullOrEmpty(info.Identifier))
            throw new ArgumentException("Extension identifier is required");
        if (string.IsNullOrEmpty(info.DisplayName))
            throw new ArgumentException("Extension display name is required");

        var elements = new List<XElement>();

        // Add version element
        elements.Add(new XElement(DescNs + "version",
            new XAttribute("value", info.Version)
        ));

        // Add identifier element
        elements.Add(new XElement(DescNs + "identifier",
            new XAttribute("value", info.Identifier)
        ));

        // Add display name
        elements.Add(new XElement(DescNs + "display-name",
            new XElement(DescNs + "name",
                new XAttribute("lang", "en"),
                info.DisplayName
            )
        ));

        // Add description if provided
        if (!string.IsNullOrEmpty(info.Description))
        {
            elements.Add(new XElement(DescNs + "extension-description",
                new XElement(DescNs + "src",
                    new XAttribute(Xlink + "href", "description_en.txt"),
                    new XAttribute("lang", "en")
                )
            ));
        }

        // Add publisher if provided
        if (!string.IsNullOrEmpty(info.Publisher))
        {
            elements.Add(new XElement(DescNs + "publisher",
                new XElement(DescNs + "name",
                    new XAttribute(Xlink + "href", ""),
                    new XAttribute("lang", "en"),
                    info.Publisher
                )
            ));
        }

        // Add release notes URL if provided
        if (!string.IsNullOrEmpty(info.ReleaseNotesUrl))
        {
            elements.Add(new XElement(DescNs + "release-notes",
                new XElement(DescNs + "src",
                    new XAttribute(Xlink + "href", info.ReleaseNotesUrl),
                    new XAttribute("lang", "en")
                )
            ));
        }

        // Add update information URL if provided
        if (!string.IsNullOrEmpty(info.UpdateUrl))
        {
            var updateSources = new List<XElement>
            {
                new XElement(DescNs + "src",
                    new XAttribute(Xlink + "href", info.UpdateUrl)
                )
            };

            // Add mirror URLs for redundancy
            foreach (var mirrorUrl in info.UpdateMirrorUrls)
            {
                if (!string.IsNullOrEmpty(mirrorUrl))
                {
                    updateSources.Add(new XElement(DescNs + "src",
                        new XAttribute(Xlink + "href", mirrorUrl)
                    ));
                }
            }

            elements.Add(new XElement(DescNs + "update-information",
                updateSources
            ));

            logger.Debug($"Added update information URL: {info.UpdateUrl}", true);
        }

        // Add icon if provided
        if (!string.IsNullOrEmpty(info.IconUrl))
        {
            elements.Add(new XElement(DescNs + "icon",
                new XElement(DescNs + "default",
                    new XAttribute(Xlink + "href", info.IconUrl)
                )
            ));
        }

        // Add LibreOffice version dependencies
        if (!string.IsNullOrEmpty(info.MinLibreOfficeVersion))
        {
            var depAttrs = new List<XAttribute>
            {
                new XAttribute("value", info.MinLibreOfficeVersion),
                new XAttribute(DescNs + "name", $"LibreOffice {info.MinLibreOfficeVersion}")
            };

            if (!string.IsNullOrEmpty(info.MaxLibreOfficeVersion))
            {
                depAttrs.Add(new XAttribute(DescNs + "maximum-version", info.MaxLibreOfficeVersion));
            }

            elements.Add(new XElement(DescNs + "dependencies",
                new XElement(DescNs + "OpenOffice.org-minimal-version",
                    depAttrs
                )
            ));

            logger.Debug($"Added dependency: LibreOffice >= {info.MinLibreOfficeVersion}", true);
        }

        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(DescNs + "description",
                new XAttribute(XNamespace.Xmlns + "d", DescNs),
                new XAttribute(XNamespace.Xmlns + "xlink", Xlink),
                elements
            )
        );

        logger.Debug("Description.xml generated successfully", true);
        return doc.ToString();
    }
}
