using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace CalcDNA.Generator;

/// <summary>
/// Generates UNO service implementation classes for [CalcAddIn] classes.
/// This replaces the need for uno-skeletonmaker / unoidl-netmaker.
/// Each [CalcAddIn] class gets a corresponding *_UnoService class that implements
/// the required UNO interfaces for LibreOffice integration.
/// </summary>
[Generator]
public class UnoServiceGenerator : IIncrementalGenerator
{
    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        // Find all classes marked with [CalcAddIn]
        var classDeclarations = context.SyntaxProvider
            .CreateSyntaxProvider(
                predicate: (node, _) => IsCalcAddInClass(node),
                transform: (ctx, _) => GetClassWithMethods(ctx))
            .Where(c => c.HasValue)
            .Select((c, _) => c!.Value);

        // Generate the UNO service implementation for each class
        context.RegisterSourceOutput(classDeclarations, (spc, classInfo) =>
        {
            var source = GenerateUnoService(classInfo);
            spc.AddSource($"{classInfo.ClassName}_UnoService.g.cs", SourceText.From(source, Encoding.UTF8));
        });
    }

    private static bool IsCalcAddInClass(SyntaxNode node)
    {
        if (node is not ClassDeclarationSyntax classDecl)
            return false;

        // Must have attributes and be public static partial
        if (classDecl.AttributeLists.Count == 0)
            return false;

        var modifiers = classDecl.Modifiers;
        return modifiers.Any(m => m.ValueText == "public") &&
               modifiers.Any(m => m.ValueText == "static") &&
               modifiers.Any(m => m.ValueText == "partial");
    }

    private static AddInClassInfo? GetClassWithMethods(GeneratorSyntaxContext context)
    {
        var classDecl = (ClassDeclarationSyntax)context.Node;
        var classSymbol = context.SemanticModel.GetDeclaredSymbol(classDecl) as INamedTypeSymbol;

        if (classSymbol == null)
            return null;

        // Check for [CalcAddIn] attribute
        var addInAttr = classSymbol.GetAttributes()
            .FirstOrDefault(a => a.AttributeClass?.Name == "CalcAddInAttribute" ||
                                 a.AttributeClass?.Name == "CalcAddIn");

        if (addInAttr == null)
            return null;

        // Get attribute values
        string addInName = classSymbol.Name;
        string description = "";

        foreach (var namedArg in addInAttr.NamedArguments)
        {
            switch (namedArg.Key)
            {
                case "Name":
                    addInName = namedArg.Value.Value?.ToString() ?? addInName;
                    break;
                case "Description":
                    description = namedArg.Value.Value?.ToString() ?? "";
                    break;
                case "Namespace":
                    // We could use this to override something if needed
                    break;
            }
        }

        // Get all [CalcFunction] methods
        var methods = classSymbol.GetMembers()
            .OfType<IMethodSymbol>()
            .Where(m => m.GetAttributes().Any(a =>
                a.AttributeClass?.Name == "CalcFunctionAttribute" ||
                a.AttributeClass?.Name == "CalcFunction"))
            .Select(m => GetMethodInfo(m))
            .ToList();

        if (methods.Count == 0)
            return null;

        return new AddInClassInfo
        {
            ClassName = classSymbol.Name,
            Namespace = classSymbol.ContainingNamespace.ToDisplayString(),
            AssemblyName = context.SemanticModel.Compilation.AssemblyName ?? "AddIn",
            AddInName = addInName,
            Description = description,
            Methods = methods
        };
    }

    private static MethodInfo GetMethodInfo(IMethodSymbol method)
    {
        var funcAttr = method.GetAttributes()
            .FirstOrDefault(a => a.AttributeClass?.Name == "CalcFunctionAttribute" ||
                                 a.AttributeClass?.Name == "CalcFunction");

        string displayName = method.Name;
        string description = "";
        string category = "Add-In";
        string helpUrl = "";
        string compatibilityName = "";
        bool isVolatile = false;

        if (funcAttr != null)
        {
            foreach (var namedArg in funcAttr.NamedArguments)
            {
                switch (namedArg.Key)
                {
                    case "Name":
                        displayName = namedArg.Value.Value?.ToString() ?? displayName;
                        break;
                    case "Description":
                        description = namedArg.Value.Value?.ToString() ?? "";
                        break;
                    case "Category":
                        category = namedArg.Value.Value?.ToString() ?? "Add-In";
                        break;
                    case "HelpUrl":
                        helpUrl = namedArg.Value.Value?.ToString() ?? "";
                        break;
                    case "CompatibilityName":
                        compatibilityName = namedArg.Value.Value?.ToString() ?? "";
                        break;
                    case "IsVolatile":
                        isVolatile = (bool)(namedArg.Value.Value ?? false);
                        break;
                }
            }
        }

        var parameters = method.Parameters.Select(p => GetParameterInfo(p)).ToList();

        return new MethodInfo
        {
            Name = method.Name,
            DisplayName = displayName,
            Description = description,
            Category = category,
            HelpUrl = helpUrl,
            CompatibilityName = compatibilityName,
            IsVolatile = isVolatile,
            ReturnType = WrapperTypeMapping.MapReturnTypeToWrapper(method.ReturnType),
            Parameters = parameters
        };
    }

    private static ParameterInfo GetParameterInfo(IParameterSymbol param)
    {
        var paramAttr = param.GetAttributes()
            .FirstOrDefault(a => a.AttributeClass?.Name == "CalcParameterAttribute" ||
                                 a.AttributeClass?.Name == "CalcParameter");

        string displayName = param.Name;
        string description = "";

        if (paramAttr != null)
        {
            foreach (var namedArg in paramAttr.NamedArguments)
            {
                switch (namedArg.Key)
                {
                    case "Name":
                        displayName = namedArg.Value.Value?.ToString() ?? displayName;
                        break;
                    case "Description":
                        description = namedArg.Value.Value?.ToString() ?? "";
                        break;
                }
            }
        }

        return new ParameterInfo
        {
            Name = param.Name,
            DisplayName = displayName,
            Description = description,
            Type = WrapperTypeMapping.MapTypeToWrapper(param.Type, WrapperTypeMapping.IsOptionalParameter(param))
        };
    }

    private static string GenerateUnoService(AddInClassInfo classInfo)
    {
        var sb = new StringBuilder();

        sb.AppendLine("// <auto-generated/>");
        sb.AppendLine("#nullable enable");
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Collections.Generic;");
        sb.AppendLine("using CalcDNA.Runtime;");
        sb.AppendLine("using CalcDNA.Runtime.Uno;");
        sb.AppendLine();

        // Namespace must match the UNO module name so the .NET bridge can
        // find the service class by its fully-qualified type name.
        // E.g. service "Demo_App.Functions" → type "Demo_App.Functions".
        string moduleName = SanitizeModuleName(classInfo.AssemblyName);
        sb.AppendLine($"namespace {moduleName}");
        sb.AppendLine("{");

        // Generate interface for the functions
        GenerateInterface(sb, classInfo);

        // Generate the service implementation
        GenerateServiceClass(sb, classInfo);

        sb.AppendLine("}");

        return sb.ToString();
    }

    private static void GenerateInterface(StringBuilder sb, AddInClassInfo classInfo)
    {
        sb.AppendLine($"    /// <summary>");
        sb.AppendLine($"    /// UNO interface for {classInfo.ClassName} functions.");
        sb.AppendLine($"    /// </summary>");
        sb.AppendLine($"    public interface IX{classInfo.ClassName}");
        sb.AppendLine("    {");

        foreach (var method in classInfo.Methods)
        {
            var paramList = string.Join(", ", method.Parameters.Select(p => $"{p.Type} {p.Name}"));
            sb.AppendLine($"        {method.ReturnType} {method.Name}({paramList});");
        }

        sb.AppendLine("    }");
        sb.AppendLine();
    }

    private static void GenerateServiceClass(StringBuilder sb, AddInClassInfo classInfo)
    {
        string interfaceName = $"IX{classInfo.ClassName}";
        // Must match the module name used by the CLI's IDL/XCU generators —
        // those sanitize the assembly name (dots → underscores) to form the UNO module.
        string moduleName = SanitizeModuleName(classInfo.AssemblyName);
        string serviceName = $"{moduleName}.{classInfo.ClassName}";

        sb.AppendLine($"    /// <summary>");
        sb.AppendLine($"    /// UNO Service implementation for {classInfo.ClassName}.");
        sb.AppendLine($"    /// Implements the required UNO interfaces for LibreOffice integration.");
        sb.AppendLine($"    /// </summary>");
        sb.AppendLine($"    public sealed class {classInfo.ClassName} : {interfaceName}, IXServiceInfo, IXLocalizable, IXAddIn");
        sb.AppendLine("    {");

        // Constants
        sb.AppendLine($"        private const string ImplementationName = \"{serviceName}\";");
        sb.AppendLine();
        sb.AppendLine("        private static readonly string[] SupportedServiceNames = new[]");
        sb.AppendLine("        {");
        sb.AppendLine("            \"com.sun.star.sheet.AddIn\",");
        sb.AppendLine($"            \"{serviceName}\"");
        sb.AppendLine("        };");
        sb.AppendLine();

        // Locale field
        sb.AppendLine("        private Locale _locale = Locale.Default;");
        sb.AppendLine();

        // Constructor with diagnostic logging
        sb.AppendLine($"        public {classInfo.ClassName}()");
        sb.AppendLine("        {");
        sb.AppendLine($"            try {{ System.IO.File.AppendAllText(\"/tmp/calcdna_debug.log\", \"{classInfo.ClassName} instantiated\\n\"); }} catch {{ }}");
        sb.AppendLine("        }");
        sb.AppendLine();

        // Function metadata dictionary
        GenerateFunctionMetadata(sb, classInfo);

        // Interface implementation (delegates to static wrapper methods)
        sb.AppendLine($"        #region {interfaceName} Implementation");
        sb.AppendLine();

        foreach (var method in classInfo.Methods)
        {
            var paramList = string.Join(", ", method.Parameters.Select(p => $"{p.Type} {p.Name}"));
            var argList = string.Join(", ", method.Parameters.Select(p => p.Name));

            sb.AppendLine($"        public {method.ReturnType} {method.Name}({paramList})");
            sb.AppendLine("        {");
            sb.AppendLine($"            try {{ System.IO.File.AppendAllText(\"/tmp/calcdna_debug.log\", \"{method.Name} called\\n\"); }} catch {{ }}");
            sb.AppendLine($"            return {classInfo.Namespace}.{classInfo.ClassName}.{method.Name}_UNOWrapper({argList});");
            sb.AppendLine("        }");
            sb.AppendLine();
        }

        sb.AppendLine("        #endregion");
        sb.AppendLine();

        // XServiceInfo implementation
        GenerateXServiceInfo(sb);

        // XLocalizable implementation
        GenerateXLocalizable(sb);

        // XAddIn implementation
        GenerateXAddIn(sb, classInfo);

        // Factory method
        GenerateFactoryMethod(sb, classInfo);

        sb.AppendLine("    }");
    }

    private static void GenerateFunctionMetadata(StringBuilder sb, AddInClassInfo classInfo)
    {
        sb.AppendLine("        // Function metadata for XAddIn interface");
        sb.AppendLine("        private static readonly Dictionary<string, FunctionMetadata> _functions = new()");
        sb.AppendLine("        {");

        foreach (var method in classInfo.Methods)
        {
            sb.AppendLine($"            [\"{method.Name}\"] = new FunctionMetadata");
            sb.AppendLine("            {");
            sb.AppendLine($"                DisplayName = \"{EscapeString(method.DisplayName)}\",");
            sb.AppendLine($"                Description = \"{EscapeString(method.Description)}\",");
            sb.AppendLine($"                Category = \"{EscapeString(method.Category)}\",");
            sb.AppendLine($"                HelpUrl = \"{EscapeString(method.HelpUrl)}\",");
            sb.AppendLine($"                CompatibilityName = \"{EscapeString(method.CompatibilityName)}\",");
            sb.AppendLine($"                IsVolatile = {method.IsVolatile.ToString().ToLower()},");
            sb.AppendLine("                Parameters = new ParameterMetadata[]");
            sb.AppendLine("                {");
            foreach (var param in method.Parameters)
            {
                sb.AppendLine($"                    new ParameterMetadata {{ Name = \"{EscapeString(param.DisplayName)}\", Description = \"{EscapeString(param.Description)}\" }},");
            }
            sb.AppendLine("                }");
            sb.AppendLine("            },");
        }

        sb.AppendLine("        };");
        sb.AppendLine();

        // Metadata classes
        sb.AppendLine("        private class FunctionMetadata");
        sb.AppendLine("        {");
        sb.AppendLine("            public string DisplayName { get; init; } = \"\";");
        sb.AppendLine("            public string Description { get; init; } = \"\";");
        sb.AppendLine("            public string Category { get; init; } = \"Add-In\";");
        sb.AppendLine("            public string HelpUrl { get; init; } = \"\";");
        sb.AppendLine("            public string CompatibilityName { get; init; } = \"\";");
        sb.AppendLine("            public bool IsVolatile { get; init; }");
        sb.AppendLine("            public ParameterMetadata[] Parameters { get; init; } = Array.Empty<ParameterMetadata>();");
        sb.AppendLine("        }");
        sb.AppendLine();
        sb.AppendLine("        private class ParameterMetadata");
        sb.AppendLine("        {");
        sb.AppendLine("            public string Name { get; init; } = \"\";");
        sb.AppendLine("            public string Description { get; init; } = \"\";");
        sb.AppendLine("        }");
        sb.AppendLine();
    }

    private static void GenerateXServiceInfo(StringBuilder sb)
    {
        sb.AppendLine("        #region IXServiceInfo Implementation");
        sb.AppendLine();
        sb.AppendLine("        public string getImplementationName() => ImplementationName;");
        sb.AppendLine();
        sb.AppendLine("        public bool supportsService(string serviceName) => Array.IndexOf(SupportedServiceNames, serviceName) >= 0;");
        sb.AppendLine();
        sb.AppendLine("        public string[] getSupportedServiceNames() => SupportedServiceNames;");
        sb.AppendLine();
        sb.AppendLine("        #endregion");
        sb.AppendLine();
    }

    private static void GenerateXLocalizable(StringBuilder sb)
    {
        sb.AppendLine("        #region IXLocalizable Implementation");
        sb.AppendLine();
        sb.AppendLine("        public void setLocale(Locale locale) => _locale = locale;");
        sb.AppendLine();
        sb.AppendLine("        public Locale getLocale() => _locale;");
        sb.AppendLine();
        sb.AppendLine("        #endregion");
        sb.AppendLine();
    }

    private static void GenerateXAddIn(StringBuilder sb, AddInClassInfo classInfo)
    {
        sb.AppendLine("        #region IXAddIn Implementation");
        sb.AppendLine();

        // getProgrammaticFuntionName (note: typo is intentional - matches UNO API)
        sb.AppendLine("        public string getProgrammaticFuntionName(string displayName)");
        sb.AppendLine("        {");
        sb.AppendLine("            foreach (var kvp in _functions)");
        sb.AppendLine("            {");
        sb.AppendLine("                if (kvp.Value.DisplayName == displayName)");
        sb.AppendLine("                    return kvp.Key;");
        sb.AppendLine("            }");
        sb.AppendLine("            return displayName;");
        sb.AppendLine("        }");
        sb.AppendLine();

        // getDisplayFunctionName
        sb.AppendLine("        public string getDisplayFunctionName(string programmaticName)");
        sb.AppendLine("        {");
        sb.AppendLine("            return _functions.TryGetValue(programmaticName, out var meta) ? meta.DisplayName : programmaticName;");
        sb.AppendLine("        }");
        sb.AppendLine();

        // getFunctionDescription
        sb.AppendLine("        public string getFunctionDescription(string programmaticName)");
        sb.AppendLine("        {");
        sb.AppendLine("            return _functions.TryGetValue(programmaticName, out var meta) ? meta.Description : \"\";");
        sb.AppendLine("        }");
        sb.AppendLine();

        // getDisplayArgumentName
        sb.AppendLine("        public string getDisplayArgumentName(string programmaticFunctionName, int argumentIndex)");
        sb.AppendLine("        {");
        sb.AppendLine("            if (_functions.TryGetValue(programmaticFunctionName, out var meta) &&");
        sb.AppendLine("                argumentIndex >= 0 && argumentIndex < meta.Parameters.Length)");
        sb.AppendLine("            {");
        sb.AppendLine("                return meta.Parameters[argumentIndex].Name;");
        sb.AppendLine("            }");
        sb.AppendLine("            return \"\";");
        sb.AppendLine("        }");
        sb.AppendLine();

        // getArgumentDescription
        sb.AppendLine("        public string getArgumentDescription(string programmaticFunctionName, int argumentIndex)");
        sb.AppendLine("        {");
        sb.AppendLine("            if (_functions.TryGetValue(programmaticFunctionName, out var meta) &&");
        sb.AppendLine("                argumentIndex >= 0 && argumentIndex < meta.Parameters.Length)");
        sb.AppendLine("            {");
        sb.AppendLine("                return meta.Parameters[argumentIndex].Description;");
        sb.AppendLine("            }");
        sb.AppendLine("            return \"\";");
        sb.AppendLine("        }");
        sb.AppendLine();

        // getProgrammaticCategoryName
        sb.AppendLine("        public string getProgrammaticCategoryName(string programmaticFunctionName)");
        sb.AppendLine("        {");
        sb.AppendLine("            return _functions.TryGetValue(programmaticFunctionName, out var meta) ? meta.Category : \"Add-In\";");
        sb.AppendLine("        }");
        sb.AppendLine();

        // getDisplayCategoryName
        sb.AppendLine("        public string getDisplayCategoryName(string programmaticFunctionName)");
        sb.AppendLine("        {");
        sb.AppendLine("            return getProgrammaticCategoryName(programmaticFunctionName);");
        sb.AppendLine("        }");
        sb.AppendLine();

        sb.AppendLine("        #endregion");
        sb.AppendLine();
    }

    private static void GenerateFactoryMethod(StringBuilder sb, AddInClassInfo classInfo)
    {
        sb.AppendLine("        #region Factory");
        sb.AppendLine();
        sb.AppendLine($"        /// <summary>");
        sb.AppendLine($"        /// Creates a new instance of the {classInfo.ClassName} UNO service.");
        sb.AppendLine($"        /// </summary>");
        sb.AppendLine($"        public static {classInfo.ClassName} Create() => new();");
        sb.AppendLine();
        sb.AppendLine("        #endregion");
    }

    /// <summary>
    /// Sanitizes an assembly name into a valid UNO module name.
    /// Must produce the same result as Program.cs SanitizeName so that the
    /// ImplementationName here matches the oor:name in the XCU and the
    /// module name in the IDL.
    /// </summary>
    private static string SanitizeModuleName(string name)
    {
        if (string.IsNullOrEmpty(name)) return "AddIn";

        var sb = new StringBuilder();
        foreach (char c in name)
        {
            if (char.IsLetterOrDigit(c) || c == '_')
                sb.Append(c);
            else
                sb.Append('_');
        }

        string result = sb.ToString().Trim('_');
        if (result.Length > 0 && char.IsDigit(result[0]))
            result = "_" + result;

        return string.IsNullOrEmpty(result) ? "AddIn" : result;
    }

    private static string EscapeString(string value)
    {
        return value
            .Replace("\\", "\\\\")
            .Replace("\"", "\\\"")
            .Replace("\n", "\\n")
            .Replace("\r", "\\r")
            .Replace("\t", "\\t");
    }

    // Internal types for collecting class information
    private struct AddInClassInfo
    {
        public string ClassName;
        public string Namespace;
        public string AssemblyName;
        public string AddInName;
        public string Description;
        public List<MethodInfo> Methods;
    }

    private struct MethodInfo
    {
        public string Name;
        public string DisplayName;
        public string Description;
        public string Category;
        public string HelpUrl;
        public string CompatibilityName;
        public bool IsVolatile;
        public string ReturnType;
        public List<ParameterInfo> Parameters;
    }

    private struct ParameterInfo
    {
        public string Name;
        public string DisplayName;
        public string Description;
        public string Type;
    }
}
