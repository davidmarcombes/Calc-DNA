using System.Reflection;
using System.Text;

namespace CalcDNA.CLI;

internal static class IdlGenerator
{
    public static string BuildIdl(string assemblyName, IEnumerable<AddInClass> addInClasses, Logger logger)
    {
        var sb = new StringBuilder();
        sb.AppendLine("#include <com/sun/star/uno/XInterface.idl>");
        sb.AppendLine();
        sb.AppendLine($"module {SanitizeModuleName(assemblyName)} {{");

        foreach (var addIn in addInClasses)
        {
            sb.AppendLine();
            sb.AppendLine($"    interface X{addIn.Type.Name} {{");

            foreach (var method in addIn.Methods)
            {
                try
                {
                    var unoReturn = UNOTypeMapping.MapTypeToUno(method.ReturnType);
                    var parameters = method.GetParameters()
                        .Select(p => $"[in] {MapParameterType(p, logger)} {p.Name}");

                    sb.AppendLine($"        {unoReturn} {method.Name}( {string.Join(", ", parameters)} );");
                }
                catch (NotSupportedException ex)
                {
                    logger.Warning($"Skipping method '{method.Name}': {ex.Message}");
                }
            }

            sb.AppendLine("    };");
        }

        sb.AppendLine("};");
        return sb.ToString();
    }

    private static string MapParameterType(ParameterInfo param, Logger logger)
    {
        try
        {
            return UNOTypeMapping.MapTypeToUno(param.ParameterType);
        }
        catch (NotSupportedException ex)
        {
            logger.Warning($"Parameter '{param.Name}' has unsupported type: {ex.Message}");
            throw;
        }
    }

    private static string SanitizeModuleName(string name)
    {
        // IDL module names must be valid identifiers - replace dots and dashes
        return name.Replace(".", "_").Replace("-", "_");
    }
}