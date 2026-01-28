namespace CalcDNA.CLI;

internal static class IdlGenerator
{
    public static string BuildIdl(Type type, IEnumerable<MethodInfo> methods)
    {
        var sb = new StringBuilder();
        sb.AppendLine("#include <com/sun/star/uno/XInterface.idl>");
        sb.AppendLine("module my_namespace {");
        sb.AppendLine($"    interface X{type.Name} {{");

        foreach (var method in methods)
        {
            var unoReturn = UnoTypeMapper.MapTypeToUno(method.ReturnType);
            var parameters = method.GetParameters()
                .Select(p => $"[in] {UnoTypeMapper.MapTypeToUno(p.ParameterType)} {p.Name}");

            sb.AppendLine($"        {unoReturn} {method.Name}( {string.Join(", ", parameters)} );");
        }

        sb.AppendLine("    };    
        sb.AppendLine("};");
        return sb.ToString();
    }
}