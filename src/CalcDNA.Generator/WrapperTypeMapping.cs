using Microsoft.CodeAnalysis;

namespace CalcDNA.Generator;
public static class WrapperTypeMapping
{
    public static string MapTypeToWrapper(ITypeSymbol typeSymbol, bool optional = false)
    {
        // Any optional parameter must be mapped to an any type
        if (optional)
            return "object";

        // Get the full type name
        var typeName = typeSymbol.Name;
        var typeNamespace = typeSymbol.ContainingNamespace?.ToDisplayString();

        if (typeSymbol is INamedTypeSymbol namedType)
        {
            // Handle nullable types
            // TODO: Consider if we want to map to any and treat as optional
            if (namedType.IsGenericType && namedType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
            {
                var nullableType = namedType.TypeArguments[0];
                return MapTypeToWrapper(nullableType, optional);
            }

            // Handle Generic Lists
            if (namedType.IsGenericType && typeName == "List" && typeNamespace == "System.Collections.Generic")
            {
                var elementType = namedType.TypeArguments[0];
                return $"{MapTypeToWrapper(elementType, optional)}[]";
            }
        }

        // Handle arrays
        if (typeSymbol is IArrayTypeSymbol arrayType)
        {
            var elementType = arrayType.ElementType;

            // Handle 1D arrays
            if (arrayType.Rank == 1)
                return $"{MapTypeToWrapper(elementType, optional)}[]";

            // Handle 2D arrays
            if (arrayType.Rank == 2)
                return $"{MapTypeToWrapper(elementType, optional)}[][]";
        }

        // Handle special CalcDNA types
        if (typeName == "CalcRange" && typeNamespace == "CalcDNA.Runtime")
            return "object[][]";

        // Handle primitive and common types
        return typeSymbol.SpecialType switch
        {
            SpecialType.System_Boolean => "bool",
            SpecialType.System_Byte => "byte",
            SpecialType.System_Int16 => "short",
            SpecialType.System_Int32 => "int",
            SpecialType.System_Int64 => "long",
            SpecialType.System_Single => "float",
            SpecialType.System_Double => "double",
            SpecialType.System_Char => "char",
            SpecialType.System_String => "string",
            SpecialType.System_Object => "object",
            _ => throw new NotSupportedException($"Type {typeName} not supported for Calc UDFs.")
        };
    }

    static bool NeedsMarshaling(ITypeSymbol typeSymbol)
    {
        // Determine if the type requires marshaling
        if (typeSymbol is IArrayTypeSymbol arrayType)
        {
            return true; // Arrays need marshaling
        }
        if (typeSymbol is INamedTypeSymbol namedType)
        {
            if (namedType.IsGenericType && namedType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
            {
                return NeedsMarshaling(namedType.TypeArguments[0]);
            }
            if (namedType.IsGenericType && namedType.Name == "List" && namedType.ContainingNamespace?.ToDisplayString() == "System.Collections.Generic")
            {
                return true; // Generic Lists need marshaling
            }
        }
        if (typeSymbol.Name == "CalcRange" && typeSymbol.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
        {
            return true; // CalcRange needs marshaling
        }
        return false; // Primitive types do not need marshaling
    }

    public static bool IsOptionalParameter(IParameterSymbol parameter)
    {
        // Check if parameter has [CalcParameter(Optional=true)] or if it has a default value
        var isOptional = parameter.GetAttributes()
            .Where(attr => attr.AttributeClass?.Name == "CalcParameterAttribute" || attr.AttributeClass?.Name == "CalcParameter")
            .Any(attr => attr.NamedArguments.Any(arg => arg.Key == "Optional" && arg.Value.Value is true)) ||
            parameter.IsOptional;
        return isOptional;
    }

    public static bool NeedsMarshaling(IParameterSymbol parameter)
    {
        if (IsOptionalParameter(parameter))
            return true; // Optional parameters need unwrapping

        var specialType = parameter.Type.SpecialType;

        return !(
           specialType == SpecialType.System_Boolean ||
           specialType == SpecialType.System_Byte ||
           specialType == SpecialType.System_Int16 ||
           specialType == SpecialType.System_Int32 ||
           specialType == SpecialType.System_Int64 ||
           specialType == SpecialType.System_Single ||
           specialType == SpecialType.System_Double ||
           specialType == SpecialType.System_Char ||
           specialType == SpecialType.System_String ||
           specialType == SpecialType.System_Object
           );
    }

    public static string GetMarshalingCode(IParameterSymbol parameter)
    {
        var typeSymbol = parameter.Type;
        var specialType = parameter.Type.SpecialType;

        if (IsOptionalParameter(parameter))
        {
            var defaultValue = parameter.ExplicitDefaultValue; 

            // Get the default value for the underlying type
            if (specialType == SpecialType.System_Double)
                return $"UnoMarshal.UnwrapOptionalDouble({parameter.Name}, {defaultValue?? 0.0})";
            else
                throw new NotSupportedException($"Optional parameter of type {typeSymbol} is not supported.");
        }

        throw new NotSupportedException($"Marshaling for type {typeSymbol.Name} is not supported.");
    }

}