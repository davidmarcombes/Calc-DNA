using System.Text;

namespace CalcDNA.CLI;

/// <summary>
/// Generates the META-INF/manifest.xml file for LibreOffice extension packages.
/// This file lists all files in the extension and their media types.
/// </summary>
internal static class ManifestGenerator
{
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

        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        sb.AppendLine("<manifest:manifest xmlns:manifest=\"http://openoffice.org/2001/manifest\">");

        // Add the XCU configuration file
        sb.AppendLine($"  <manifest:file-entry manifest:full-path=\"{addInName}.xcu\" manifest:media-type=\"application/vnd.sun.star.configuration-data\"/>");

        // Add the RDB type library
        sb.AppendLine($"  <manifest:file-entry manifest:full-path=\"{addInName}.rdb\" manifest:media-type=\"application/vnd.sun.star.uno-typelibrary;type=RDB\"/>");

        // Add all assembly files
        foreach (var assemblyFile in assemblyFiles)
        {
            string fileName = Path.GetFileName(assemblyFile);

            // Only the main add-in DLL is a UNO component; dependency DLLs
            // (CalcDNA.Runtime, CalcDNA.Attributes) are plain assemblies.
            string mediaType = Path.GetExtension(fileName).ToLowerInvariant() switch
            {
                ".dll" when fileName.StartsWith("CalcDNA.", StringComparison.OrdinalIgnoreCase)
                    => "application/octet-stream",
                ".dll" => "application/vnd.sun.star.uno-component;type=.NET",
                ".json" => "application/json",
                _ => "application/octet-stream"
            };

            sb.AppendLine($"  <manifest:file-entry manifest:full-path=\"{fileName}\" manifest:media-type=\"{mediaType}\"/>");
            logger.Debug($"  Added: {fileName} ({mediaType})", true);
        }

        sb.AppendLine("</manifest:manifest>");

        logger.Debug("Manifest.xml generated successfully", true);
        return sb.ToString();
    }
}
