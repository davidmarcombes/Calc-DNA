using System.Text.Json;
using System.Text.Json.Nodes;

namespace CalcDNA.CLI;

/// <summary>
/// Generates the .runtimeconfig.json file for .NET assemblies.
/// This file tells the .NET CLR bridge which framework version to use.
/// </summary>
internal static class RuntimeConfigGenerator
{
    /// <summary>
    /// Generates a .runtimeconfig.json file for the specified target framework.
    /// </summary>
    /// <param name="targetFramework">Target framework (e.g., "net10.0", "net8.0")</param>
    /// <param name="logger">Logger instance for output</param>
    /// <returns>JSON content for .runtimeconfig.json</returns>
    public static string BuildRuntimeConfig(string targetFramework, Logger logger)
    {
        logger.Debug($"Building .runtimeconfig.json for {targetFramework}...", true);

        // Parse the target framework to extract version information
        var (frameworkName, version) = ParseTargetFramework(targetFramework);

        var config = new JsonObject
        {
            ["runtimeOptions"] = new JsonObject
            {
                ["tfm"] = targetFramework,
                ["framework"] = new JsonObject
                {
                    ["name"] = frameworkName,
                    ["version"] = version
                },
                ["configProperties"] = new JsonObject
                {
                    ["System.Runtime.Serialization.EnableUnsafeBinaryFormatterSerialization"] = false
                }
            }
        };

        var options = new JsonSerializerOptions
        {
            WriteIndented = true
        };

        string json = JsonSerializer.Serialize(config, options);
        logger.Debug(".runtimeconfig.json generated successfully", true);
        return json;
    }

    /// <summary>
    /// Parses a target framework moniker into framework name and version.
    /// </summary>
    /// <param name="targetFramework">Target framework string (e.g., "net10.0", "net8.0")</param>
    /// <returns>Tuple of (framework name, version string)</returns>
    private static (string frameworkName, string version) ParseTargetFramework(string targetFramework)
    {
        // Handle different TFM formats:
        // net10.0 -> Microsoft.NETCore.App 10.0.0
        // net8.0 -> Microsoft.NETCore.App 8.0.0
        // net6.0 -> Microsoft.NETCore.App 6.0.0

        if (targetFramework.StartsWith("net", StringComparison.OrdinalIgnoreCase))
        {
            // Extract version number
            string versionPart = targetFramework.Substring(3); // Remove "net"

            // Parse version (e.g., "10.0" -> "10.0.0")
            if (double.TryParse(versionPart, out double versionNumber))
            {
                // Modern .NET (5.0+)
                if (versionNumber >= 5.0)
                {
                    return ("Microsoft.NETCore.App", $"{versionNumber}.0");
                }
                // .NET Core 3.x
                else if (versionNumber >= 3.0)
                {
                    return ("Microsoft.NETCore.App", $"{versionNumber}.0");
                }
            }
        }

        // Default fallback
        return ("Microsoft.NETCore.App", "8.0.0");
    }

    /// <summary>
    /// Extracts the target framework from an assembly's metadata.
    /// </summary>
    /// <param name="assemblyPath">Path to the assembly</param>
    /// <param name="logger">Logger instance</param>
    /// <returns>Target framework string, or null if not found</returns>
    public static string? GetTargetFramework(string assemblyPath, Logger logger)
    {
        try
        {
            // Check for .deps.json file first (most reliable)
            string depsJsonPath = Path.ChangeExtension(assemblyPath, ".deps.json");
            if (File.Exists(depsJsonPath))
            {
                string json = File.ReadAllText(depsJsonPath);
                using var doc = JsonDocument.Parse(json);

                if (doc.RootElement.TryGetProperty("runtimeTarget", out var runtimeTarget))
                {
                    if (runtimeTarget.TryGetProperty("name", out var name))
                    {
                        // Extract TFM from ".NETCoreApp,Version=v10.0"
                        string runtimeTargetName = name.GetString() ?? "";
                        if (runtimeTargetName.Contains("Version=v"))
                        {
                            var parts = runtimeTargetName.Split("=v");
                            if (parts.Length > 1)
                            {
                                string version = parts[1];
                                return $"net{version}";
                            }
                        }
                    }
                }
            }

            logger.Debug($"Could not determine target framework from {assemblyPath}, using default", true);
            return null;
        }
        catch (Exception ex)
        {
            logger.Debug($"Error reading target framework: {ex.Message}", true);
            return null;
        }
    }
}
