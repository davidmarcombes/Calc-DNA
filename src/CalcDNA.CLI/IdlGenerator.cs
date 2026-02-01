using System.Reflection;
using System.Text;
using System.Collections.Generic;

namespace CalcDNA.CLI;

/// <summary>
/// IDL generator for LibreOffice Calc add-ins.
/// </summary>
internal static class IdlGenerator
{
    private static readonly HashSet<string> ReservedWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "module", "interface", "service", "any", "boolean", "byte", "char",
        "double", "enum", "exception", "FALSE", "float", "hyper", "long",
        "octet", "sequence", "short", "string", "struct", "TRUE", "type",
        "typedef", "union", "unsigned", "void", "in", "out", "inout"
    };

    /// <summary>
    /// Build the IDL file content.
    /// </summary>
    public static string BuildIdl(string assemblyName, IEnumerable<AddInClass> addInClasses, Logger logger)
    {
        var sb = new StringBuilder();
        sb.AppendLine("#include <com/sun/star/uno/XInterface.idl>");
        sb.AppendLine("#include <com/sun/star/sheet/XAddIn.idl>");
        sb.AppendLine("#include <com/sun/star/lang/XServiceInfo.idl>");
        sb.AppendLine();
        
        string moduleName = SanitizeIdentifier(assemblyName);
        sb.AppendLine($"module {moduleName} {{");

        foreach (var addIn in addInClasses)
        {
            string interfaceName = "X" + SanitizeIdentifier(addIn.Type.Name);
            
            sb.AppendLine();
            sb.AppendLine($"    interface {interfaceName} {{");

            foreach (var method in addIn.Methods)
            {
                try
                {
                    var unoReturn = UnoTypeMapping.MapTypeToUno(method.ReturnType);
                    var parameters = method.GetParameters()
                        .Select(p => $"[in] {MapParameterType(p, logger)} {SanitizeIdentifier(p.Name ?? "arg")}");

                    string methodName = SanitizeIdentifier(method.Name);
                    sb.AppendLine($"        {unoReturn} {methodName}( {string.Join(", ", parameters)} );");
                }
                catch (NotSupportedException ex)
                {
                    logger.Warning($"Skipping method '{method.Name}' in {addIn.Type.Name}: {ex.Message}");
                }
                catch (Exception ex)
                {
                    logger.Error($"Error processing method '{method.Name}' in {addIn.Type.Name}: {ex.Message}");
                }
            }

            sb.AppendLine("    };");

            // Modern UNO IDL uses single-interface services
            string implementationName = SanitizeIdentifier(addIn.Type.Name);
            sb.AppendLine();
            sb.AppendLine($"    service {implementationName} : {interfaceName};");
        }

        sb.AppendLine("};");
        return sb.ToString();
    }

    private static string MapParameterType(ParameterInfo param, Logger logger)
    {
        try
        {
            return UnoTypeMapping.MapTypeToUno(param.ParameterType);
        }
        catch (NotSupportedException ex)
        {
            logger.Warning($"Parameter '{param.Name}' has unsupported type: {ex.Message}");
            throw;
        }
    }

    /// <summary>
    /// Sanitizes a string to be a valid IDL identifier.
    /// IDL identifiers must start with a letter and contain only letters, digits, and underscores.
    /// </summary>
    private static string SanitizeIdentifier(string name)
    {
        if (string.IsNullOrEmpty(name)) return "id";

        var sb = new StringBuilder();
        foreach (char c in name)
        {
            if (char.IsLetterOrDigit(c) || c == '_')
            {
                sb.Append(c);
            }
            else
            {
                sb.Append('_');
            }
        }

        string result = sb.ToString().Trim('_');
        
        if (string.IsNullOrEmpty(result)) result = "id";

        // IDL identifiers cannot start with a digit
        if (char.IsDigit(result[0]))
        {
            result = "_" + result;
        }

        // Avoid reserved words
        if (ReservedWords.Contains(result))
        {
            result = "_" + result;
        }

        return result;
    }
}