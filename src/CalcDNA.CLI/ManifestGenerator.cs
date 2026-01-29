using System.Xml.Linq;

namespace CalcDNA.CLI;

/// <summary>
/// Generates the META-INF/manifest.xml file for LibreOffice extension packages.
/// This file lists all files in the extension and their media types.
/// </summary>
internal static class ManifestGenerator
{
    private static readonly XNamespace ManifestNs = "http://openoffice.org/2001/manifest";

    /// <summary>
    /// Generates the manifest.xml content for an OXT package.
    /// </summary>
    /// <param name="addInName">Name of the add-in (used for file naming)</param>
    /// <param name="assemblyFiles">List of assembly DLL files to include</param>
    /// <param name="logger">Logger instance for output</param>
    /// <returns>XML content for manifest.xml</returns>
    public static string BuildManifest(string addInName, IEnumerable<string> assemblyFiles, Logger logger)
    {
        logger.Debug("Building manifest.xml...", true);

        var fileEntries = new List<XElement>();

        // Add the XCU configuration file
        fileEntries.Add(CreateFileEntry(
            $"{addInName}.xcu",
            "application/vnd.sun.star.configuration-data"
        ));

        // Add the RDB type library
        fileEntries.Add(CreateFileEntry(
            $"{addInName}.rdb",
            "application/vnd.sun.star.uno-typelibrary;type=RDB"
        ));

        // Add all assembly files
        foreach (var assemblyFile in assemblyFiles)
        {
            string fileName = Path.GetFileName(assemblyFile);

            // Determine media type based on file extension
            string mediaType = Path.GetExtension(fileName).ToLowerInvariant() switch
            {
                ".dll" => "application/vnd.sun.star.uno-component;type=.NET",
                ".json" => "application/json",
                _ => "application/octet-stream"
            };

            fileEntries.Add(CreateFileEntry(fileName, mediaType));
            logger.Debug($"  Added: {fileName} ({mediaType})", true);
        }

        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(ManifestNs + "manifest",
                fileEntries
            )
        );

        logger.Debug("Manifest.xml generated successfully", true);
        return doc.ToString();
    }

    /// <summary>
    /// Creates a file-entry element for the manifest.
    /// </summary>
    private static XElement CreateFileEntry(string path, string mediaType)
    {
        return new XElement(ManifestNs + "file-entry",
            new XAttribute(ManifestNs + "full-path", path),
            new XAttribute(ManifestNs + "media-type", mediaType)
        );
    }
}
