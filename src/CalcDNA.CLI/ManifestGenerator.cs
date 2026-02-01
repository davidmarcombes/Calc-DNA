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
    /// <param name="pythonScriptFile">If set, switches to Python mode: adds the .py entry
    /// with type=Python and marks all DLLs as octet-stream (loaded by pythonnet, not LO).</param>
    /// <returns>XML content for manifest.xml</returns>
    public static string BuildManifest(string addInName, IEnumerable<string> assemblyFiles, Logger logger,
        string? pythonScriptFile = null)
    {
        logger.Debug("Building manifest.xml...", true);

        var sb = new StringBuilder();
        sb.AppendLine("<?xml version=\"1.0\" encoding=\"UTF-8\"?>");
        sb.AppendLine("<manifest:manifest xmlns:manifest=\"http://openoffice.org/2001/manifest\">");

        // Add the XCU configuration file
        sb.AppendLine($"  <manifest:file-entry manifest:full-path=\"{addInName}.xcu\" manifest:media-type=\"application/vnd.sun.star.configuration-data\"/>");

        // Add the RDB type library
        sb.AppendLine($"  <manifest:file-entry manifest:full-path=\"{addInName}.rdb\" manifest:media-type=\"application/vnd.sun.star.uno-typelibrary;type=RDB\"/>");

        // In Python mode, register the .py script as a UNO Python component.
        // LO's pythonloader picks this up and loads the module as a component factory.
        if (pythonScriptFile != null)
        {
            sb.AppendLine($"  <manifest:file-entry manifest:full-path=\"{pythonScriptFile}\" manifest:media-type=\"application/vnd.sun.star.uno-component;type=Python\"/>");
            logger.Debug($"  Added: {pythonScriptFile} (Python UNO component)", true);
        }

        // Add all assembly files
        foreach (var assemblyFile in assemblyFiles)
        {
            string fileName = Path.GetFileName(assemblyFile);

            string mediaType;
            if (pythonScriptFile != null)
            {
                // Python mode: all DLLs are loaded by pythonnet at runtime, not by LO's .NET bridge
                mediaType = Path.GetExtension(fileName).ToLowerInvariant() switch
                {
                    ".dll" => "application/octet-stream",
                    ".json" => "application/json",
                    _ => "application/octet-stream"
                };
            }
            else
            {
                // .NET mode: main add-in DLL is a UNO .NET component;
                // CalcDNA dependency DLLs are plain assemblies.
                mediaType = Path.GetExtension(fileName).ToLowerInvariant() switch
                {
                    ".dll" when fileName.StartsWith("CalcDNA.", StringComparison.OrdinalIgnoreCase)
                        => "application/octet-stream",
                    ".dll" => "application/vnd.sun.star.uno-component;type=.NET",
                    ".json" => "application/json",
                    _ => "application/octet-stream"
                };
            }

            sb.AppendLine($"  <manifest:file-entry manifest:full-path=\"{fileName}\" manifest:media-type=\"{mediaType}\"/>");
            logger.Debug($"  Added: {fileName} ({mediaType})", true);
        }

        sb.AppendLine("</manifest:manifest>");

        logger.Debug("Manifest.xml generated successfully", true);
        return sb.ToString();
    }
}
