
namespace CalcDNA.CLI;

public static class UNOTypeMapping
{
       public static string MapTypeToUno(Type type)
    {
        // Handle nullable types
        if (type.IsGenericType && type.GetGenericTypeDefinition() == typeof(Nullable<>))
        {
            var nullableType = type.GetGenericArguments()[0];
            return MapTypeToUno(nullableType);
        }

        // Handle 1D Arrays
        if (type.IsArray && type.GetArrayRank() == 1)
        {
            var elementType = type.GetElementType();
            if (elementType is null)
                throw new NotSupportedException($"Type {type.Name} not supported for Calc UDFs.");
            return $"sequence< {MapTypeToUno(elementType)} >";
        }

        // Handle 2D Arrays (Calc Ranges)
        if (type.IsArray && type.GetArrayRank() == 2)
        {
            var elementType = type.GetElementType();
            if (elementType is null)
                throw new NotSupportedException($"Type {type.Name} not supported for Calc UDFs.");
            return $"sequence< sequence< {MapTypeToUno(elementType)} > >";
        }

        // Standard mappings
        return type.Name switch
        {
            "Boolean" => "boolean",
            "Byte"    => "byte",
            "Short"   => "short",
            "Int32"   => "long",
            "Int64"   => "hyper",
            "Float"   => "float",
            "Double"  => "double",
            "Char"    => "char",
            "String"  => "string",
            "Object"  => "any",
            _         => throw new NotSupportedException($"Type {type.Name} not supported for Calc UDFs.")
        };
    }
}