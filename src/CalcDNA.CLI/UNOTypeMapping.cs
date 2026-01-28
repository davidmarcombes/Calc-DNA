namespace CalcDNA.CLI;

/// <summary>
/// Maps .NET types to UNO IDL types for LibreOffice Calc add-in generation.
/// </summary>
public static class UNOTypeMapping
{
    public static string MapTypeToUno(Type type)
    {
        // Note: We use name-based checks because MetadataLoadContext loads types
        // from a different context, so direct typeof() comparisons fail.

        // Handle nullable types - map to "any" since UNO doesn't have nullable primitives
        if (type.IsGenericType && IsGenericTypeDefinition(type, "System.Nullable`1"))
        {
            return "any";
        }

        // Handle 1D Arrays -> sequence<T>
        if (type.IsArray && type.GetArrayRank() == 1)
        {
            var elementType = type.GetElementType();
            if (elementType is null)
                throw new NotSupportedException($"Type {type.Name} not supported for Calc UDFs.");
            return $"sequence< {MapTypeToUno(elementType)} >";
        }

        // Handle 2D Arrays -> sequence<sequence<T>>
        if (type.IsArray && type.GetArrayRank() == 2)
        {
            var elementType = type.GetElementType();
            if (elementType is null)
                throw new NotSupportedException($"Type {type.Name} not supported for Calc UDFs.");
            return $"sequence< sequence< {MapTypeToUno(elementType)} > >";
        }

        // Handle List<T> -> sequence<T>
        if (type.IsGenericType && IsGenericTypeDefinition(type, "System.Collections.Generic.List`1"))
        {
            var elementType = type.GetGenericArguments()[0];
            return $"sequence< {MapTypeToUno(elementType)} >";
        }

        // Handle CalcRange -> sequence<sequence<any>>
        if (type.Name == "CalcRange" && type.Namespace == "CalcDNA.Runtime")
        {
            return "sequence< sequence< any > >";
        }

        // Handle DateTime -> double (OLE Automation date)
        if (type.FullName == "System.DateTime" || type.Name == "DateTime")
        {
            return "double";
        }

        // Standard type mappings (.NET type name -> UNO IDL type)
        return type.Name switch
        {
            "Boolean" => "boolean",
            "Byte"    => "byte",        // UNO: unsigned 8-bit
            "SByte"   => "short",       // UNO: no signed 8-bit, map to short
            "Int16"   => "short",       // UNO: signed 16-bit
            "Int32"   => "long",        // UNO: signed 32-bit (confusingly named)
            "Int64"   => "hyper",       // UNO: signed 64-bit
            "Single"  => "float",       // UNO: 32-bit float
            "Double"  => "double",      // UNO: 64-bit float
            "Char"    => "char",        // UNO: 16-bit Unicode
            "String"  => "string",      // UNO: Unicode string
            "Object"  => "any",         // UNO: variant type
            _         => throw new NotSupportedException($"Type '{type.FullName ?? type.Name}' is not supported for Calc UDFs.")
        };
    }

    /// <summary>
    /// Checks if a type's generic type definition matches the expected full name.
    /// This is needed because MetadataLoadContext types can't be compared directly with typeof().
    /// </summary>
    private static bool IsGenericTypeDefinition(Type type, string expectedFullName)
    {
        if (!type.IsGenericType) return false;
        var genericDef = type.GetGenericTypeDefinition();
        return genericDef.FullName == expectedFullName;
    }
}