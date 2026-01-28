using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;

namespace CalcDNA.CLI;

class Program
{
    static void Main(string[] args)
    {
        if (args.Length < 1) {
            Console.WriteLine("Usage: CalcDNA.CLI <path_to_your_addon.dll>");
            return;
        }

        string assemblyPath = Path.GetFullPath(args[0]);
        string outputDir = Path.GetDirectoryName(assemblyPath) ?? "";

        // 1. Setup MetadataLoadContext to avoid file locking
        var paths = new List<string> { assemblyPath };
        paths.AddRange(Directory.GetFiles(RuntimeEnvironment.GetRuntimeDirectory(), "*.dll"));

        var resolver = new PathAssemblyResolver(paths);
        using var mlc = new MetadataLoadContext(resolver);

        // 2. Load the assembly
        var assembly = mlc.LoadFromAssemblyPath(assemblyPath);

        // 3. Process classes with [CalcAddIn]
        var addInClasses = assembly.GetTypes()
            .Where(t => t.GetCustomAttributesData().Any(
                a => a.AttributeType.Name == "CalcAddInAttribute" || a.AttributeType.Name == "CalcAddIn"));

        foreach (var type in addInClasses)
        {
            // TODO: Not sure if we need to generate metadata for each class or just once
            GenerateMetadata(type, outputDir);
        }
    }

    static void GenerateMetadata(Type type, string outputDir)
    {
        Console.WriteLine($"Generating metadata for {type.FullName}...");

        var methods = type.GetMethods()
            .Where(m => m.GetCustomAttributesData().Any(a => a.AttributeType.Name == "CalcFunctionAttribute"));

        // Generate IDL file
        string idlContent = BuildIdl(type, methods);
        File.WriteAllText(Path.Combine(outputDir, $"{type.Name}.idl"), idlContent);

        // Generate CalcAddIns.xcu (The Function Wizard XML)
        string xcuContent = BuildXcu(type, methods);
        File.WriteAllText(Path.Combine(outputDir, "CalcAddIns.xcu"), xcuContent);
    }

    static string BuildXcu(Type type, IEnumerable<MethodInfo> methods)
    {
        // Generate the XML structure for the Function Wizard
        return XcuGenerator .Generate(type, methods);
    }

    static string BuildIdl(Type type, IEnumerable<MethodInfo> methods)
    {
        return IdlGenerator.BuildIdl(type, methods);
    }

}