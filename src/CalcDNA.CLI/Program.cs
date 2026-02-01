using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using CommandLine;
using CommandLine.Text;

namespace CalcDNA.CLI;

/// <summary>
/// Record to hold an Add-In class and its methods
/// </summary>
public record AddInClass(Type Type, List<MethodInfo> Methods);

class Program
{
    /// <summary>
    /// Command line options
    /// </summary>
    public class Options
    {
        [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
        public bool Verbose { get; set; } = false;

        [Value(0, MetaName = "assembly", HelpText = "Path to the CalcDNA Add-In assembly (.dll) to process.", Required = true)]
        public string AssemblyPath { get; set; } = "";

        [Value(1, MetaName = "output", HelpText = "Output directory for generated files. Defaults to the assembly directory.", Required = false)]
        public string? OutputPath { get; set; } = null;

        [Value(2, MetaName = "name", HelpText = "Name of the addin to build. Must be a valid identifier (letters, digits, underscores).", Required = false)]
        public string? Name { get; set; } = null;

        [Value(3, MetaName = "sdk", HelpText = "Path of the LibreOffice SDK", Required = false)]
        public string? SDKPath { get; set; }

        [Option('p', "package", Required = false, HelpText = "Create an .oxt package file.")]
        public bool CreatePackage { get; set; } = false;

        [Option("version", Required = false, HelpText = "Extension version (e.g., 1.0.0). Used when creating package.", Default = "1.0.0")]
        public string Version { get; set; } = "1.0.0";

        [Option("publisher", Required = false, HelpText = "Publisher name for the extension. Used when creating package.")]
        public string? Publisher { get; set; } = null;

        [Option("description", Required = false, HelpText = "Extension description. Used when creating package.")]
        public string? Description { get; set; } = null;

        [Option("min-lo-version", Required = false, HelpText = "Minimum LibreOffice version (e.g., 4.0).", Default = "4.0")]
        public string MinLibreOfficeVersion { get; set; } = "4.0";

        [Option("update-url", Required = false, HelpText = "URL to update.xml file for automatic update checking.")]
        public string? UpdateUrl { get; set; } = null;

        [Option("release-notes-url", Required = false, HelpText = "URL to release notes for this version.")]
        public string? ReleaseNotesUrl { get; set; } = null;
    }

    /// <summary>
    /// Main entry point
    /// </summary>
    /// <param name="args">Command line arguments</param>
    static void Main(string[] args)
    {
        Parser.Default.ParseArguments<Options>(args)
                   .WithParsed<Options>(o =>
                   {
                       var logger = new Logger();
                       try
                       {
                           Run(o, logger);
                       }
                       catch (Exception ex)
                       {
                           logger.Error($"Fatal error: {ex.Message}");
                           if (o.Verbose)
                           {
                               logger.Debug(ex.StackTrace ?? "", true);
                           }
                           Environment.Exit(1);
                       }
                   });
    }

    /// <summary>
    /// Run the program
    /// </summary>
    /// <param name="o">Command line options</param>
    /// <param name="logger">Logger instance</param>
    static void Run(Options o, Logger logger)
    {
        bool verbose = o.Verbose;

        // Retrieve and resolve assembly path
        string assemblyPath = Path.GetFullPath(o.AssemblyPath);
        if (!File.Exists(assemblyPath))
        {
            logger.Error($"Assembly not found: {assemblyPath}");
            return;
        }

        // Use specified output directory or default to assembly directory
        string outputDir = !string.IsNullOrEmpty(o.OutputPath)
            ? Path.GetFullPath(o.OutputPath)
            : Path.GetDirectoryName(assemblyPath) ?? Environment.CurrentDirectory;

        try
        {
            if (!Directory.Exists(outputDir))
            {
                logger.Debug($"Creating output directory: {outputDir}", verbose);
                Directory.CreateDirectory(outputDir);
            }
        }
        catch (Exception ex)
        {
            logger.Error($"Could not create or access output directory '{outputDir}': {ex.Message}");
            return;
        }

        // 1. Setup MetadataLoadContext to avoid file locking
        // We need to resolve all dependencies. MetadataLoadContext doesn't do this automatically.
        var paths = new List<string>();

        // Add the target assembly
        paths.Add(assemblyPath);

        // Add all DLLs from the assembly directory to help resolution of dependencies
        var assemblyDir = Path.GetDirectoryName(assemblyPath);
        if (!string.IsNullOrEmpty(assemblyDir) && Directory.Exists(assemblyDir))
        {
            foreach (var dll in Directory.GetFiles(assemblyDir, "*.dll"))
            {
                if (!paths.Contains(dll, StringComparer.OrdinalIgnoreCase))
                {
                    paths.Add(dll);
                }
            }
        }

        // Add runtime assemblies (core .NET libraries)
        var runtimeDir = RuntimeEnvironment.GetRuntimeDirectory();
        if (Directory.Exists(runtimeDir))
        {
            foreach (var dll in Directory.GetFiles(runtimeDir, "*.dll"))
            {
                if (!paths.Contains(dll, StringComparer.OrdinalIgnoreCase))
                {
                    paths.Add(dll);
                }
            }
        }

        logger.Debug($"MetadataLoadContext initialized with {paths.Count} potential assembly paths.", verbose);

        var resolver = new PathAssemblyResolver(paths);
        using var mlc = new MetadataLoadContext(resolver);

        // 2. Load the assembly
        Assembly assembly;
        try
        {
            assembly = mlc.LoadFromAssemblyPath(assemblyPath);
        }
        catch (Exception ex)
        {
            logger.Error($"Failed to load assembly '{assemblyPath}': {ex.Message}");
            return;
        }

        // 3. Collect all [CalcAddIn] classes with their [CalcFunction] methods
        var addInClasses = new List<AddInClass>();
        try
        {
            var types = assembly.GetTypes();
            foreach (var type in types)
            {
                var attrs = type.GetCustomAttributesData();
                bool isAddIn = attrs.Any(a => a.AttributeType.Name == "CalcAddInAttribute" || a.AttributeType.Name == "CalcAddIn");

                if (isAddIn)
                {
                    var methods = type.GetMethods(BindingFlags.Public | BindingFlags.Instance | BindingFlags.Static)
                        .Where(m => m.GetCustomAttributesData().Any(a => a.AttributeType.Name == "CalcFunctionAttribute"))
                        .ToList();

                    if (methods.Any())
                    {
                        addInClasses.Add(new AddInClass(type, methods));
                    }
                }
            }
        }
        catch (ReflectionTypeLoadException ex)
        {
            logger.Error("Failed to load some types from the assembly. This usually means a dependency is missing.");
            foreach (var loaderException in ex.LoaderExceptions)
            {
                if (loaderException != null)
                {
                    logger.Error($"  - {loaderException.Message}");
                }
            }
            return;
        }

        if (addInClasses.Count == 0)
        {
            logger.Warning("No [CalcAddIn] classes with [CalcFunction] methods found.");
            return;
        }

        // 4. Determine Add-In Name
        // Use command-line name if provided, otherwise use assembly name
        string? addinName = o.Name;
        if (string.IsNullOrEmpty(addinName))
        {
            addinName = assembly.GetName().Name ?? "AddIn";
        }

        // Sanitize and validate addinName for use in file paths and IDL module names
        addinName = SanitizeName(addinName);
        if (string.IsNullOrWhiteSpace(addinName))
        {
            logger.Error("The resulting Add-In name is invalid.");
            return;
        }

        logger.Debug($"Generating metadata for '{addinName}'...", verbose);
        foreach (var addIn in addInClasses)
        {
            logger.Debug($"  - {addIn.Type.FullName}: {addIn.Methods.Count} function(s)", verbose);
        }

        // 5. Generate IDL file
        string idlPath = Path.Combine(outputDir, $"{addinName}.idl");
        try
        {
            string idlContent = IdlGenerator.BuildIdl(addinName, addInClasses, logger);
            File.WriteAllText(idlPath, idlContent);
            logger.Success($"Generated {addinName}.idl");
        }
        catch (Exception ex)
        {
            logger.Error($"Failed to generate IDL: {ex.Message}");
            return; // Cannot continue without IDL
        }

        // 6. Generate XCU file
        try
        {
            string xcuContent = XcuGenerator.BuildXcu(addinName, addInClasses, logger);
            File.WriteAllText(Path.Combine(outputDir, $"{addinName}.xcu"), xcuContent);
            logger.Success($"Generated {addinName}.xcu");
        }
        catch (Exception ex)
        {
            logger.Error($"Failed to generate XCU: {ex.Message}");
        }

        // 7. Generate RDB file
        string rdbPath = Path.Combine(outputDir, $"{addinName}.rdb");
        bool rdbExists = false;
        try
        {
            RdbGenerator.WriteRdb(idlPath, rdbPath, o.SDKPath, logger);
            logger.Success($"Generated {addinName}.rdb");
            rdbExists = true;
        }
        catch (Exception ex)
        {
            logger.Error($"Failed to generate RDB: {ex.Message}");

            // Check if RDB file already exists from a previous run
            if (File.Exists(rdbPath))
            {
                logger.Warning($"Using existing RDB file: {rdbPath}");
                rdbExists = true;
            }
            else if (o.CreatePackage)
            {
                logger.Error("Cannot create package without RDB file.");
                logger.Info("To create an RDB file, either:");
                logger.Info("  1. Install the LibreOffice SDK and ensure it's in your PATH");
                logger.Info("  2. Specify the SDK path with --sdk-path option");
                logger.Info("  3. Generate the RDB file manually using the LibreOffice SDK tools");
                return;
            }
        }

        // 8. Create OXT package if requested
        if (o.CreatePackage && rdbExists)
        {
            try
            {
                CreateOxtPackage(assemblyPath, outputDir, addinName, rdbPath, o, logger, verbose, assembly);
            }
            catch (Exception ex)
            {
                logger.Error($"Failed to create OXT package: {ex.Message}");
                if (verbose)
                {
                    logger.Debug(ex.StackTrace ?? "", true);
                }
            }
        }
    }

    /// <summary>
    /// Extracts extension metadata from assembly attributes.
    /// </summary>
    private static DescriptionGenerator.ExtensionInfo ExtractExtensionMetadata(
        Assembly assembly,
        string addinName,
        Options options,
        Logger logger,
        bool verbose)
    {
        var metadata = new DescriptionGenerator.ExtensionInfo
        {
            // Set defaults
            Identifier = $"org.calcdna.{addinName.ToLowerInvariant()}",
            Version = "1.0.0",
            DisplayName = addinName,
            Publisher = "",
            Description = $"{addinName} - LibreOffice Calc Add-in created with Calc-DNA",
            MinLibreOfficeVersion = "4.0",
            MaxLibreOfficeVersion = ""
        };

        // Look for CalcExtensionMetadataAttribute
        var metadataAttr = assembly.GetCustomAttributesData()
            .FirstOrDefault(a => a.AttributeType.Name == "CalcExtensionMetadataAttribute");

        if (metadataAttr != null)
        {
            logger.Debug("Found CalcExtensionMetadata attribute", verbose);

            foreach (var namedArg in metadataAttr.NamedArguments)
            {
                var value = namedArg.TypedValue.Value?.ToString();
                if (string.IsNullOrEmpty(value)) continue;

                switch (namedArg.MemberName)
                {
                    case "Version":
                        metadata.Version = value;
                        logger.Debug($"  Version: {value}", verbose);
                        break;
                    case "DisplayName":
                        metadata.DisplayName = value;
                        logger.Debug($"  DisplayName: {value}", verbose);
                        break;
                    case "Publisher":
                        metadata.Publisher = value;
                        logger.Debug($"  Publisher: {value}", verbose);
                        break;
                    case "Description":
                        metadata.Description = value;
                        logger.Debug($"  Description: {value}", verbose);
                        break;
                    case "Identifier":
                        metadata.Identifier = value;
                        logger.Debug($"  Identifier: {value}", verbose);
                        break;
                    case "MinLibreOfficeVersion":
                        metadata.MinLibreOfficeVersion = value;
                        logger.Debug($"  MinLibreOfficeVersion: {value}", verbose);
                        break;
                    case "MaxLibreOfficeVersion":
                        metadata.MaxLibreOfficeVersion = value;
                        logger.Debug($"  MaxLibreOfficeVersion: {value}", verbose);
                        break;
                    case "UpdateUrl":
                        metadata.UpdateUrl = value;
                        logger.Debug($"  UpdateUrl: {value}", verbose);
                        break;
                    case "ReleaseNotesUrl":
                        metadata.ReleaseNotesUrl = value;
                        logger.Debug($"  ReleaseNotesUrl: {value}", verbose);
                        break;
                    case "IconUrl":
                        metadata.IconUrl = value;
                        logger.Debug($"  IconUrl: {value}", verbose);
                        break;
                }
            }
        }
        else
        {
            logger.Debug("No CalcExtensionMetadata attribute found, using defaults", verbose);
        }

        // Command-line options override assembly attributes
        if (!string.IsNullOrEmpty(options.Version) && options.Version != "1.0.0")
        {
            metadata.Version = options.Version;
            logger.Debug($"Version overridden by command line: {options.Version}", verbose);
        }

        if (!string.IsNullOrEmpty(options.Name))
        {
            metadata.DisplayName = options.Name;
            logger.Debug($"DisplayName overridden by command line: {options.Name}", verbose);
        }

        if (!string.IsNullOrEmpty(options.Publisher))
        {
            metadata.Publisher = options.Publisher;
            logger.Debug($"Publisher overridden by command line: {options.Publisher}", verbose);
        }

        if (!string.IsNullOrEmpty(options.Description))
        {
            metadata.Description = options.Description;
            logger.Debug($"Description overridden by command line: {options.Description}", verbose);
        }

        if (!string.IsNullOrEmpty(options.MinLibreOfficeVersion) && options.MinLibreOfficeVersion != "4.0")
        {
            metadata.MinLibreOfficeVersion = options.MinLibreOfficeVersion;
            logger.Debug($"MinLibreOfficeVersion overridden by command line: {options.MinLibreOfficeVersion}", verbose);
        }

        if (!string.IsNullOrEmpty(options.UpdateUrl))
        {
            metadata.UpdateUrl = options.UpdateUrl;
            logger.Debug($"UpdateUrl overridden by command line: {options.UpdateUrl}", verbose);
        }

        if (!string.IsNullOrEmpty(options.ReleaseNotesUrl))
        {
            metadata.ReleaseNotesUrl = options.ReleaseNotesUrl;
            logger.Debug($"ReleaseNotesUrl overridden by command line: {options.ReleaseNotesUrl}", verbose);
        }

        return metadata;
    }

    /// <summary>
    /// Creates an OXT package from the generated files.
    /// </summary>
    private static void CreateOxtPackage(
        string assemblyPath,
        string outputDir,
        string addinName,
        string rdbPath,
        Options options,
        Logger logger,
        bool verbose,
        Assembly assembly)
    {
        logger.Info("Creating OXT package...");

        // Extract extension metadata from assembly attributes
        var extensionInfo = ExtractExtensionMetadata(assembly, addinName, options, logger, verbose);
        logger.Info($"Extension: {extensionInfo.DisplayName} v{extensionInfo.Version}");
        if (!string.IsNullOrEmpty(extensionInfo.Publisher))
        {
            logger.Info($"Publisher: {extensionInfo.Publisher}");
        }

        // Collect dependency assemblies
        var dependencyAssemblies = new List<string>();
        string? assemblyDir = Path.GetDirectoryName(assemblyPath);

        if (!string.IsNullOrEmpty(assemblyDir))
        {
            // Add CalcDNA.Runtime.dll
            string runtimeDll = Path.Combine(assemblyDir, "CalcDNA.Runtime.dll");
            if (File.Exists(runtimeDll))
            {
                dependencyAssemblies.Add(runtimeDll);
                logger.Debug($"Found dependency: CalcDNA.Runtime.dll", verbose);
            }
            else
            {
                logger.Warning("CalcDNA.Runtime.dll not found in assembly directory");
            }

            // Add CalcDNA.Attributes.dll
            string attributesDll = Path.Combine(assemblyDir, "CalcDNA.Attributes.dll");
            if (File.Exists(attributesDll))
            {
                dependencyAssemblies.Add(attributesDll);
                logger.Debug($"Found dependency: CalcDNA.Attributes.dll", verbose);
            }
            else
            {
                logger.Warning("CalcDNA.Attributes.dll not found in assembly directory");
            }
        }

        // Determine target framework
        string? targetFramework = RuntimeConfigGenerator.GetTargetFramework(assemblyPath, logger);
        if (string.IsNullOrEmpty(targetFramework))
        {
            targetFramework = "net8.0"; // Default fallback
            logger.Debug($"Using default target framework: {targetFramework}", verbose);
        }
        else
        {
            logger.Debug($"Detected target framework: {targetFramework}", verbose);
        }

        // Create package configuration
        var packageConfig = new OxtPackager.PackageConfig
        {
            AddInName = addinName,
            MainAssemblyPath = assemblyPath,
            DependencyAssemblies = dependencyAssemblies,
            IdlFilePath = Path.Combine(outputDir, $"{addinName}.idl"),
            XcuFilePath = Path.Combine(outputDir, $"{addinName}.xcu"),
            RdbFilePath = rdbPath,
            TargetFramework = targetFramework,
            OutputPath = Path.Combine(outputDir, $"{addinName}.oxt"),
            ExtensionInfo = extensionInfo
        };

        // Create the package
        string oxtPath = OxtPackager.CreatePackage(packageConfig, logger);
        logger.Success($"OXT package created successfully: {oxtPath}");
    }

    /// <summary>
    /// Sanitizes a string to be used as a valid identifier and file name.
    /// </summary>
    private static string SanitizeName(string name)
    {
        if (string.IsNullOrEmpty(name)) return "AddIn";

        // Remove invalid filename characters
        char[] invalidChars = Path.GetInvalidFileNameChars();
        var sb = new System.Text.StringBuilder();

        foreach (char c in name)
        {
            if (char.IsLetterOrDigit(c) || c == '_')
            {
                sb.Append(c);
            }
            else if (c == ' ' || c == '.' || c == '-')
            {
                sb.Append('_');
            }
        }

        string result = sb.ToString().Trim('_');

        // Ensure it doesn't start with a digit (IDL requirement)
        if (result.Length > 0 && char.IsDigit(result[0]))
        {
            result = "_" + result;
        }

        return string.IsNullOrEmpty(result) ? "AddIn" : result;
    }
}
