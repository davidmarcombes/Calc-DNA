using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using CommandLine;
using CommandLine.Text;

namespace CalcDNA.CLI;

public record AddInClass(Type Type, List<MethodInfo> Methods);

class Program
{
    public class Options
    {
        [Option('v', "verbose", Required = false, HelpText = "Set output to verbose messages.")]
        public bool Verbose { get; set; } = false;

        [Value(0, MetaName = "assembly", HelpText = "Path to the CalcDNA Add-In assembly (.dll) to process.", Required = true)]
        public string AssemblyPath { get; set; } = "";

        [Value(1, MetaName = "output", HelpText = "Output directory for generated files. Defaults to the assembly directory.", Required = false)]
        public string? OutputPath { get; set; } = null;

        [Value(2, MetaName = "name", HelpText = "Name of the addin to build", Required = false)]
        public string? Name { get; set; } = null;

        [Value(3, MetaName = "sdk", HelpText = "Path of the LibreOffice SDK", Required = false)]
        public string? SDKPath { get; set; }

    }

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
                       }
                   });
    }

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
            : Path.GetDirectoryName(assemblyPath) ?? "";

        if (!Directory.Exists(outputDir))
        {
            logger.Debug($"Creating output directory: {outputDir}", verbose);
            Directory.CreateDirectory(outputDir);
        }

        // 1. Setup MetadataLoadContext to avoid file locking
        var paths = new List<string> { assemblyPath };

        // Add assemblies from the same directory as the target (for dependencies like CalcDNA.Attributes)
        var assemblyDir = Path.GetDirectoryName(assemblyPath);
        if (!string.IsNullOrEmpty(assemblyDir))
        {
            paths.AddRange(Directory.GetFiles(assemblyDir, "*.dll"));
        }

        // Add runtime assemblies
        paths.AddRange(Directory.GetFiles(RuntimeEnvironment.GetRuntimeDirectory(), "*.dll"));

        var resolver = new PathAssemblyResolver(paths);
        using var mlc = new MetadataLoadContext(resolver);

        // 2. Load the assembly
        var assembly = mlc.LoadFromAssemblyPath(assemblyPath);

        // 3. Collect all [CalcAddIn] classes with their [CalcFunction] methods
        var addInClasses = assembly.GetTypes()
            .Where(t => t.GetCustomAttributesData().Any(
                a => a.AttributeType.Name == "CalcAddInAttribute" || a.AttributeType.Name == "CalcAddIn"))
            .Select(type => new AddInClass(
                type,
                type.GetMethods()
                    .Where(m => m.GetCustomAttributesData().Any(a => a.AttributeType.Name == "CalcFunctionAttribute"))
                    .ToList()))
            .Where(c => c.Methods.Count > 0)
            .ToList();

        if (addInClasses.Count == 0)
        {
            logger.Warning("No [CalcAddIn] classes with [CalcFunction] methods found.");
            return;
        }

        // Use specified name or default to assembly name
        string addinName = !string.IsNullOrEmpty(o.Name)
            ? o.Name
            : assembly.GetName().Name ?? "AddIn";

        logger.Debug($"Generating metadata for '{addinName}'...", verbose);
        foreach (var addIn in addInClasses)
        {
            logger.Debug($"  - {addIn.Type.FullName}: {addIn.Methods.Count} function(s)", verbose);
        }

        string idlPath = Path.Combine(outputDir, $"{addinName}.idl");

        // 4. Generate IDL file
        try
        {
            string idlContent = IdlGenerator.BuildIdl(addinName, addInClasses, logger);
            File.WriteAllText(idlPath, idlContent);
            logger.Success($"Generated {addinName}.idl");
        }
        catch (Exception ex)
        {
            logger.Error($"Failed to generate IDL: {ex.Message}");
        }

        // 5. Generate XCU file
        try
        {
            string xcuContent = XcuGenerator.BuildXcu(addInClasses, logger);
            File.WriteAllText(Path.Combine(outputDir, $"{addinName}.xcu"), xcuContent);
            logger.Success($"Generated {addinName}.xcu");
        }
        catch (Exception ex)
        {
            logger.Error($"Failed to generate XCU: {ex.Message}");
        }

        // 6. Generate RDB file
        try
        {
            string rdbPath = Path.Combine(outputDir, $"{addinName}.rdb");
            RdbGenerator.WriteRdb(idlPath, rdbPath, o.SDKPath, logger);
            logger.Success($"Generated {addinName}.rdb");
        }
        catch (Exception ex)
        {
            logger.Error($"Failed to generate RDB: {ex.Message}");
        }
    }
}
