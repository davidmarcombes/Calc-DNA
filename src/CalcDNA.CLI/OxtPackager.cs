using System.IO.Compression;

namespace CalcDNA.CLI;

/// <summary>
/// Creates LibreOffice extension packages (.oxt files).
/// An OXT file is a ZIP archive with a specific structure containing
/// metadata files, assemblies, and LibreOffice configuration.
/// </summary>
internal static class OxtPackager
{
    /// <summary>
    /// Configuration for OXT package creation.
    /// </summary>
    public class PackageConfig
    {
        /// <summary>
        /// Name of the add-in (used for file naming).
        /// </summary>
        public string AddInName { get; set; } = "";

        /// <summary>
        /// Path to the main add-in assembly.
        /// </summary>
        public string MainAssemblyPath { get; set; } = "";

        /// <summary>
        /// Paths to dependency assemblies (CalcDNA.Runtime, CalcDNA.Attributes, etc.).
        /// </summary>
        public List<string> DependencyAssemblies { get; set; } = new();

        /// <summary>
        /// Path to the generated IDL file.
        /// </summary>
        public string IdlFilePath { get; set; } = "";

        /// <summary>
        /// Path to the generated XCU file.
        /// </summary>
        public string XcuFilePath { get; set; } = "";

        /// <summary>
        /// Path to the generated RDB file.
        /// </summary>
        public string RdbFilePath { get; set; } = "";

        /// <summary>
        /// Extension metadata for description.xml.
        /// </summary>
        public DescriptionGenerator.ExtensionInfo ExtensionInfo { get; set; } = new();

        /// <summary>
        /// Target framework (e.g., "net10.0").
        /// </summary>
        public string TargetFramework { get; set; } = "net8.0";

        /// <summary>
        /// Output path for the .oxt file.
        /// </summary>
        public string OutputPath { get; set; } = "";

        /// <summary>
        /// When true, generates a Python UNO bridge script instead of using the .NET bridge.
        /// DLLs are loaded by pythonnet at runtime; LO invokes the Python service via pythonloader.
        /// </summary>
        public bool PythonMode { get; set; } = false;

        /// <summary>
        /// Add-in classes (required when PythonMode is true, for Python script generation).
        /// </summary>
        public List<AddInClass> AddInClasses { get; set; } = new();
    }

    /// <summary>
    /// Creates an OXT package with all necessary files.
    /// </summary>
    /// <param name="config">Package configuration</param>
    /// <param name="logger">Logger instance for output</param>
    /// <returns>Path to the created .oxt file</returns>
    public static string CreatePackage(PackageConfig config, Logger logger)
    {
        logger.Info($"Creating OXT package for {config.AddInName}...");

        // Validate configuration
        ValidateConfig(config);

        // Determine output path
        string oxtPath = string.IsNullOrEmpty(config.OutputPath)
            ? Path.Combine(Path.GetDirectoryName(config.MainAssemblyPath) ?? ".", $"{config.AddInName}.oxt")
            : config.OutputPath;

        // Ensure output directory exists
        string? outputDir = Path.GetDirectoryName(oxtPath);
        if (!string.IsNullOrEmpty(outputDir) && !Directory.Exists(outputDir))
        {
            Directory.CreateDirectory(outputDir);
        }

        // Delete existing package if it exists
        if (File.Exists(oxtPath))
        {
            logger.Debug($"Removing existing package: {oxtPath}", true);
            File.Delete(oxtPath);
        }

        // Create temporary directory for package contents
        string tempDir = Path.Combine(Path.GetTempPath(), $"calcdna_{Guid.NewGuid():N}");
        try
        {
            Directory.CreateDirectory(tempDir);
            logger.Debug($"Using temporary directory: {tempDir}", true);

            // Generate and write files
            GeneratePackageFiles(config, tempDir, logger);

            // Create ZIP archive with .oxt extension
            logger.Debug("Creating ZIP archive...", true);
            ZipFile.CreateFromDirectory(tempDir, oxtPath, CompressionLevel.Optimal, false);

            logger.Success($"OXT package created: {oxtPath}");
            return oxtPath;
        }
        finally
        {
            // Clean up temporary directory
            if (Directory.Exists(tempDir))
            {
                try
                {
                    Directory.Delete(tempDir, true);
                }
                catch (Exception ex)
                {
                    logger.Warning($"Failed to delete temporary directory: {ex.Message}");
                }
            }
        }
    }

    /// <summary>
    /// Generates all files needed in the OXT package.
    /// </summary>
    private static void GeneratePackageFiles(PackageConfig config, string tempDir, Logger logger)
    {
        // Create META-INF directory
        string metaInfDir = Path.Combine(tempDir, "META-INF");
        Directory.CreateDirectory(metaInfDir);

        // Collect all assembly files (main + dependencies + JSON files)
        var assemblyFiles = new List<string> { config.MainAssemblyPath };
        assemblyFiles.AddRange(config.DependencyAssemblies);

        // Add deps.json if it exists
        string depsJsonPath = Path.ChangeExtension(config.MainAssemblyPath, ".deps.json");
        if (File.Exists(depsJsonPath))
        {
            assemblyFiles.Add(depsJsonPath);
        }

        // In Python mode, generate the UNO service bridge script
        string? pythonScriptFile = null;
        if (config.PythonMode)
        {
            string assemblyFileName = Path.GetFileName(config.MainAssemblyPath);
            string pythonScript = PythonUnoServiceGenerator.BuildPythonScript(
                config.AddInName, assemblyFileName, config.AddInClasses, logger);
            pythonScriptFile = $"{config.AddInName}.py";
            File.WriteAllText(Path.Combine(tempDir, pythonScriptFile), pythonScript);
            logger.Debug($"Generated: {pythonScriptFile}", true);
        }

        // Generate manifest.xml
        logger.Debug("Generating manifest.xml...", true);
        string manifestContent = ManifestGenerator.BuildManifest(
            config.AddInName,
            assemblyFiles.Select(Path.GetFileName).Where(f => f != null).Cast<string>(),
            logger,
            pythonScriptFile
        );
        File.WriteAllText(Path.Combine(metaInfDir, "manifest.xml"), manifestContent);

        // Generate description.xml
        logger.Debug("Generating description.xml...", true);
        string descriptionContent = DescriptionGenerator.BuildDescription(config.ExtensionInfo, logger);
        File.WriteAllText(Path.Combine(tempDir, "description.xml"), descriptionContent);

        // Generate .runtimeconfig.json
        logger.Debug("Generating .runtimeconfig.json...", true);
        string runtimeConfigContent = RuntimeConfigGenerator.BuildRuntimeConfig(config.TargetFramework, logger);
        string runtimeConfigFileName = Path.GetFileNameWithoutExtension(config.MainAssemblyPath) + ".runtimeconfig.json";
        File.WriteAllText(Path.Combine(tempDir, runtimeConfigFileName), runtimeConfigContent);

        // Copy XCU file
        if (File.Exists(config.XcuFilePath))
        {
            string xcuDestPath = Path.Combine(tempDir, Path.GetFileName(config.XcuFilePath));
            File.Copy(config.XcuFilePath, xcuDestPath, true);
            logger.Debug($"Copied: {Path.GetFileName(config.XcuFilePath)}", true);
        }
        else
        {
            throw new FileNotFoundException($"XCU file not found: {config.XcuFilePath}");
        }

        // Copy RDB file
        if (File.Exists(config.RdbFilePath))
        {
            string rdbDestPath = Path.Combine(tempDir, Path.GetFileName(config.RdbFilePath));
            File.Copy(config.RdbFilePath, rdbDestPath, true);
            logger.Debug($"Copied: {Path.GetFileName(config.RdbFilePath)}", true);
        }
        else
        {
            throw new FileNotFoundException($"RDB file not found: {config.RdbFilePath}");
        }

        // Copy all assembly files
        foreach (var assemblyPath in assemblyFiles)
        {
            if (File.Exists(assemblyPath))
            {
                string destPath = Path.Combine(tempDir, Path.GetFileName(assemblyPath));
                File.Copy(assemblyPath, destPath, true);
                logger.Debug($"Copied: {Path.GetFileName(assemblyPath)}", true);
            }
            else
            {
                logger.Warning($"Assembly file not found: {assemblyPath}");
            }
        }
    }

    /// <summary>
    /// Validates the package configuration.
    /// </summary>
    private static void ValidateConfig(PackageConfig config)
    {
        if (string.IsNullOrEmpty(config.AddInName))
            throw new ArgumentException("AddInName is required");

        if (string.IsNullOrEmpty(config.MainAssemblyPath))
            throw new ArgumentException("MainAssemblyPath is required");

        if (!File.Exists(config.MainAssemblyPath))
            throw new FileNotFoundException($"Main assembly not found: {config.MainAssemblyPath}");

        if (string.IsNullOrEmpty(config.XcuFilePath))
            throw new ArgumentException("XcuFilePath is required");

        if (string.IsNullOrEmpty(config.RdbFilePath))
            throw new ArgumentException("RdbFilePath is required");

        if (string.IsNullOrEmpty(config.ExtensionInfo.Identifier))
            config.ExtensionInfo.Identifier = $"org.calcdna.{config.AddInName.ToLowerInvariant()}";

        if (string.IsNullOrEmpty(config.ExtensionInfo.DisplayName))
            config.ExtensionInfo.DisplayName = config.AddInName;
    }
}
