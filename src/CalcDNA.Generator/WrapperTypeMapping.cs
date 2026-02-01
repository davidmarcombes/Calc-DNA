using Microsoft.CodeAnalysis;

namespace CalcDNA.Generator;

public static class WrapperTypeMapping
{
    public static string MapTypeToWrapper(ITypeSymbol typeSymbol, bool optional)
    {
        // 1. Mandatory "Any" (object) cases
        if (optional)
            return "object?";

        // 2. Handle Nullables (System.Nullable<T>)
        if (typeSymbol.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
            return "object?";

        // 3. Handle Arrays
        if (typeSymbol is IArrayTypeSymbol arrayType)
        {
            if (arrayType.Rank > 1 || arrayType.ElementType is IArrayTypeSymbol)
                return "object[][]";
            return "object[]";
        }

        // 4. Handle CalcRange
        if (typeSymbol.Name == "CalcRange" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
        {
            return "object[][]";
        }

        // 5. Handle List<T> / IEnumerable<T> -> object[]
        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } listType &&
            listType.ContainingNamespace?.ToDisplayString() == "System.Collections.Generic" &&
            (listType.Name is "List" or "IEnumerable" or "ICollection" or "IList"))
        {
            return "object[]";
        }

        // 6. Handle DateTime -> double (OLE Automation date)
        if (typeSymbol.Name == "DateTime" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "System")
        {
            return "double";
        }

        // 7. POD Mapping with UNO-Specific adjustments
        return typeSymbol.SpecialType switch
        {
            SpecialType.System_Boolean => "bool",
            SpecialType.System_Byte => "byte",   // UNO byte (unsigned 8)
            SpecialType.System_SByte => "short",  // Map up to avoid signedness issues
            SpecialType.System_Int16 => "short",  // UNO short (signed 16)
            SpecialType.System_Int32 => "int",    // UNO long (signed 32)
            SpecialType.System_Int64 => "long",   // UNO hyper (signed 64) - BE CAREFUL HERE
            SpecialType.System_Single => "float",
            SpecialType.System_Double => "double",
            SpecialType.System_String => "string",
            _ => "object?"  // Default safety net
        };
    }

    public static string MapReturnTypeToWrapper(ITypeSymbol typeSymbol)
    {
        if (typeSymbol.SpecialType == SpecialType.System_Void)
            return "void";

        if (typeSymbol.SpecialType == SpecialType.System_DateTime)
            return "double";

        if (typeSymbol is IArrayTypeSymbol arrayType)
        {
            if (arrayType.Rank > 1 || arrayType.ElementType is IArrayTypeSymbol)
                return "object[][]";
            return "object[]";
        }

        if (typeSymbol.Name == "CalcRange" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
        {
            return "object[][]";
        }

        // Generic collections
        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } namedType)
        {
            var name = namedType.Name;
            var ns = namedType.ContainingNamespace?.ToDisplayString();
            if (ns == "System.Collections.Generic" && (name is "List" or "IEnumerable" or "ICollection" or "IList"))
                return "object[]";
        }

        return MapTypeToWrapper(typeSymbol, false);
    }

    public static bool IsTypeMappedToObject(ITypeSymbol typeSymbol, bool optional)
    {
        // 1. Mandatory "Any" (object) cases
        if (optional)
            return true;

        // 2. Handle Nullables (System.Nullable<T>)
        if (typeSymbol.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
            return true;

        // 3. POD Mapping with UNO-Specific adjustments
        return typeSymbol.SpecialType switch
        {
            SpecialType.System_Boolean => false,
            SpecialType.System_Byte => false,   
            SpecialType.System_SByte => false,  
            SpecialType.System_Int16 => false,  
            SpecialType.System_Int32 => false,  
            SpecialType.System_Int64 => false,  
            SpecialType.System_Single => false,
            SpecialType.System_Double => false,
            SpecialType.System_String => false,
            _ => true  // Default safety net
        };
    }

    public static bool IsOptionalParameter(IParameterSymbol parameter)
    {
        // Check if parameter has [CalcParameter(Optional=true)] or has a default value
        var hasOptionalAttribute = parameter.GetAttributes()
            .Where(attr => attr.AttributeClass?.Name is "CalcParameterAttribute" or "CalcParameter")
            .Any(attr => attr.NamedArguments.Any(arg => arg.Key == "Optional" && arg.Value.Value is true));

        return hasOptionalAttribute || parameter.IsOptional;
    }

    public static bool NeedsMarshaling(IParameterSymbol parameter)
    {
        return IsTypeMappedToObject(parameter.Type, IsOptionalParameter(parameter));
    }

    public static string GetMarshalingCode(IParameterSymbol parameter)
    {
        // Optional parameter marshaling 
        if (IsOptionalParameter(parameter))
        {
            return GetOptionalMarshalingCode(parameter);
        }

        // Nullable parameter marshalling
        if (parameter.Type.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
        {
            return GetNullableMarshalingCode(parameter);
        }

        // Complex data structure
        return GetComplexTypeMarshalingCode(parameter);

    }

    public static string GetReturnMarshalingCode(IMethodSymbol method, string callCode)
    {
        var type = method.ReturnType;

        if (type.SpecialType == SpecialType.System_Void)
            return $"{callCode};";

        if (type.SpecialType == SpecialType.System_DateTime)
            return $"return {callCode}.ToOADate();";

        if (type is IArrayTypeSymbol arrayType)
        {
            if (arrayType.Rank == 1)
                return $"return UnoMarshal.ToUno1DArray({callCode});";
            else
                return $"return UnoMarshal.ToUno2DArray({callCode});";
        }

        if (type.Name == "CalcRange" &&
            type.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
        {
            return $"return UnoMarshal.ToUno2DArray({callCode});";
        }

        if (type is INamedTypeSymbol { IsGenericType: true } namedType)
        {
            var name = namedType.Name;
            var ns = namedType.ContainingNamespace?.ToDisplayString();
            if (ns == "System.Collections.Generic" && (name is "List" or "IEnumerable" or "ICollection" or "IList"))
                return $"return UnoMarshal.ToUno1DArray({callCode});";
        }

        return $"return {callCode};";
    }

    // ─── Python Wrapper Mapping ────────────────────────────────────────
    // Sequence types become 'object' because pythonnet passes Python
    // lists/tuples as generic IEnumerable, not typed arrays.
    // Scalar optional/nullable unwrappers reuse UnoMarshal (identical logic).
    // Return marshaling uses PyMarshal.ToPy* (preserves null, no DBNull).

    public static string MapTypeToPyWrapper(ITypeSymbol typeSymbol, bool optional)
    {
        if (optional)
            return "object?";

        if (typeSymbol.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
            return "object?";

        // Sequences: 'object' — pythonnet iterables
        if (typeSymbol is IArrayTypeSymbol)
            return "object";

        if (typeSymbol.Name == "CalcRange" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
            return "object";

        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } listType &&
            listType.ContainingNamespace?.ToDisplayString() == "System.Collections.Generic" &&
            (listType.Name is "List" or "IEnumerable" or "ICollection" or "IList"))
            return "object";

        // DateTime, PODs, fallback: identical to UNO
        if (typeSymbol.Name == "DateTime" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "System")
            return "double";

        return typeSymbol.SpecialType switch
        {
            SpecialType.System_Boolean => "bool",
            SpecialType.System_Byte => "byte",
            SpecialType.System_SByte => "short",
            SpecialType.System_Int16 => "short",
            SpecialType.System_Int32 => "int",
            SpecialType.System_Int64 => "long",
            SpecialType.System_Single => "float",
            SpecialType.System_Double => "double",
            SpecialType.System_String => "string",
            _ => "object?"
        };
    }

    public static string GetPyMarshalingCode(IParameterSymbol parameter)
    {
        if (IsOptionalParameter(parameter))
            return GetPyOptionalMarshalingCode(parameter);

        if (parameter.Type.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
            return GetPyNullableMarshalingCode(parameter);

        return GetPyComplexTypeMarshalingCode(parameter);
    }

    public static string GetPyReturnMarshalingCode(IMethodSymbol method, string callCode)
    {
        var type = method.ReturnType;

        if (type.SpecialType == SpecialType.System_Void)
            return $"{callCode};";

        if (type.SpecialType == SpecialType.System_DateTime)
            return $"return {callCode}.ToOADate();";

        if (type is IArrayTypeSymbol arrayType)
        {
            if (arrayType.Rank == 1)
                return $"return PyMarshal.ToPy1DArray({callCode});";
            else
                return $"return PyMarshal.ToPy2DArray({callCode});";
        }

        if (type.Name == "CalcRange" &&
            type.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
            return $"return PyMarshal.ToPy2DArray({callCode});";

        if (type is INamedTypeSymbol { IsGenericType: true } namedType)
        {
            var name = namedType.Name;
            var ns = namedType.ContainingNamespace?.ToDisplayString();
            if (ns == "System.Collections.Generic" && (name is "List" or "IEnumerable" or "ICollection" or "IList"))
                return $"return PyMarshal.ToPy1DArray({callCode});";
        }

        return $"return {callCode};";
    }

    private static string GetPyOptionalMarshalingCode(IParameterSymbol parameter)
    {
        var code = GetPyComplexOptionalTypeMarshalingCode(parameter);
        if (code != null) return code;
        return GetOptionalMarshalingCode(parameter);
    }

    private static string GetPyNullableMarshalingCode(IParameterSymbol parameter)
    {
        var code = GetPyComplexNullableTypeMarshalingCode(parameter);
        if (code != null) return code;
        return GetNullableMarshalingCode(parameter);
    }

    private static string GetPyComplexTypeMarshalingCode(IParameterSymbol parameter)
    {
        var typeSymbol = parameter.Type;
        var paramName = parameter.Name;

        if (typeSymbol.SpecialType == SpecialType.System_DateTime)
            return $"UnoMarshal.ToDateTime({paramName})";

        if (typeSymbol is IArrayTypeSymbol arrayType)
        {
            var elementTypeName = GetFullTypeName(arrayType.ElementType);
            return arrayType.Rank switch
            {
                1 => $"PyMarshal.To1DArray<{elementTypeName}>({paramName})",
                2 => $"PyMarshal.To2DArray<{elementTypeName}>({paramName})",
                _ => throw new NotSupportedException($"Arrays with rank {arrayType.Rank} are not supported.")
            };
        }

        if (typeSymbol.Name == "CalcRange" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
            return $"PyMarshal.ToCalcRange({paramName})";

        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } listType &&
            listType.Name == "List" &&
            listType.ContainingNamespace?.ToDisplayString() == "System.Collections.Generic")
        {
            var elementTypeName = GetFullTypeName(listType.TypeArguments[0]);
            return $"PyMarshal.ToList<{elementTypeName}>({paramName})";
        }

        throw new NotSupportedException($"Marshaling for type '{typeSymbol.ToDisplayString()}' is not supported.");
    }

    private static string? GetPyComplexOptionalTypeMarshalingCode(IParameterSymbol parameter)
    {
        var typeSymbol = parameter.Type;
        var paramName = parameter.Name;
        var defaultLiteral = FormatDefaultValue(parameter);

        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } nullableType &&
            nullableType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
            typeSymbol = nullableType.TypeArguments[0];

        if (typeSymbol is IArrayTypeSymbol arrayType)
        {
            var elementTypeName = GetFullTypeName(arrayType.ElementType);
            return arrayType.Rank switch
            {
                1 => $"PyMarshal.UnwrapOptional1DArray<{elementTypeName}>({paramName}, {defaultLiteral})",
                2 => $"PyMarshal.UnwrapOptional2DArray<{elementTypeName}>({paramName}, {defaultLiteral})",
                _ => throw new NotSupportedException($"Optional arrays with rank {arrayType.Rank} are not supported.")
            };
        }

        if (typeSymbol.Name == "CalcRange" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
            return $"PyMarshal.UnwrapOptionalCalcRange({paramName}, {defaultLiteral})";

        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } listType &&
            listType.Name == "List" &&
            listType.ContainingNamespace?.ToDisplayString() == "System.Collections.Generic")
        {
            var elementTypeName = GetFullTypeName(listType.TypeArguments[0]);
            return $"PyMarshal.UnwrapOptionalList<{elementTypeName}>({paramName}, {defaultLiteral})";
        }

        return null;
    }

    private static string? GetPyComplexNullableTypeMarshalingCode(IParameterSymbol parameter)
    {
        var typeSymbol = parameter.Type;
        var paramName = parameter.Name;

        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } nullableType &&
            nullableType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
            typeSymbol = nullableType.TypeArguments[0];

        if (typeSymbol is IArrayTypeSymbol arrayType)
        {
            var elementTypeName = GetFullTypeName(arrayType.ElementType);
            return arrayType.Rank switch
            {
                1 => $"PyMarshal.UnwrapNullable1DArray<{elementTypeName}>({paramName})",
                2 => $"PyMarshal.UnwrapNullable2DArray<{elementTypeName}>({paramName})",
                _ => throw new NotSupportedException($"Nullable arrays with rank {arrayType.Rank} are not supported.")
            };
        }

        if (typeSymbol.Name == "CalcRange" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
            return $"PyMarshal.UnwrapNullableCalcRange({paramName})";

        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } listType &&
            listType.Name == "List" &&
            listType.ContainingNamespace?.ToDisplayString() == "System.Collections.Generic")
        {
            var elementTypeName = GetFullTypeName(listType.TypeArguments[0]);
            return $"PyMarshal.UnwrapNullableList<{elementTypeName}>({paramName})";
        }

        return null;
    }

    // ─── UNO Wrapper Private Helpers ────────────────────────────────────

    private static string GetOptionalMarshalingCode(IParameterSymbol parameter)
    {
        // Check for complex optional types first
        var code = GetComplexOptionalTypeMarshalingCode(parameter);
        if (code != null) {
            return code;
        }

        // Handle primitive optional types
        var typeSymbol = parameter.Type;
        var paramName = parameter.Name;

        // This will correctly check if type is nullable and return null if no default is specified
        var defaultLiteral = FormatDefaultValue(parameter);

        // Handle primitives
        return typeSymbol.SpecialType switch
        {
            SpecialType.System_Boolean => $"UnoMarshal.UnwrapOptionalBool({paramName}, {defaultLiteral})",
            SpecialType.System_Byte => $"UnoMarshal.UnwrapOptionalByte({paramName}, {defaultLiteral})",
            SpecialType.System_SByte => $"UnoMarshal.UnwrapOptionalSByte({paramName}, {defaultLiteral})",
            SpecialType.System_Int16 => $"UnoMarshal.UnwrapOptionalShort({paramName}, {defaultLiteral})",
            SpecialType.System_Int32 => $"UnoMarshal.UnwrapOptionalInt({paramName}, {defaultLiteral})",
            SpecialType.System_Int64 => $"UnoMarshal.UnwrapOptionalLong({paramName}, {defaultLiteral})",
            SpecialType.System_Single => $"UnoMarshal.UnwrapOptionalFloat({paramName}, {defaultLiteral})",
            SpecialType.System_Double => $"UnoMarshal.UnwrapOptionalDouble({paramName}, {defaultLiteral})",
            SpecialType.System_Char => $"UnoMarshal.UnwrapOptionalChar({paramName}, {defaultLiteral})",
            SpecialType.System_String => $"UnoMarshal.UnwrapOptionalString({paramName}, {defaultLiteral})",
            SpecialType.System_Object => $"UnoMarshal.UnwrapOptionalObject({paramName}, {defaultLiteral})",
            _ => throw new NotSupportedException($"Optional parameter of type '{typeSymbol.ToDisplayString()}' is not supported.")
        };
    }

    private static string GetNullableMarshalingCode(IParameterSymbol parameter)
    {
        // Check for complex nullable types first
        var code = GetComplexNullableTypeMarshalingCode(parameter);

        if (code != null) {
            return code;
        }

        var typeSymbol = parameter.Type;
        var paramName = parameter.Name;

        // Unwrap Nullable<T> to get the underlying type
        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } nullableType &&
            nullableType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
        {
            typeSymbol = nullableType.TypeArguments[0];
        }

        // Handle primitives based on underlying type
        return typeSymbol.SpecialType switch
        {
            SpecialType.System_Boolean => $"UnoMarshal.UnwrapNullableBool({paramName})",
            SpecialType.System_Byte => $"UnoMarshal.UnwrapNullableByte({paramName})",
            SpecialType.System_SByte => $"UnoMarshal.UnwrapNullableSByte({paramName})",
            SpecialType.System_Int16 => $"UnoMarshal.UnwrapNullableShort({paramName})",
            SpecialType.System_Int32 => $"UnoMarshal.UnwrapNullableInt({paramName})",
            SpecialType.System_Int64 => $"UnoMarshal.UnwrapNullableLong({paramName})",
            SpecialType.System_Single => $"UnoMarshal.UnwrapNullableFloat({paramName})",
            SpecialType.System_Double => $"UnoMarshal.UnwrapNullableDouble({paramName})",
            SpecialType.System_Char => $"UnoMarshal.UnwrapNullableChar({paramName})",
            SpecialType.System_String => $"UnoMarshal.UnwrapNullableString({paramName})",
            SpecialType.System_Object => $"UnoMarshal.UnwrapNullableObject({paramName})",
            SpecialType.System_DateTime => $"UnoMarshal.UnwrapNullableDateTime({paramName})",
            _ => throw new NotSupportedException($"Nullable parameter of type '{typeSymbol.ToDisplayString()}' is not supported.")
        };
    }

    private static string GetComplexTypeMarshalingCode(IParameterSymbol parameter)
    {
        var typeSymbol = parameter.Type;
        var paramName = parameter.Name;

        // DateTime
        if (typeSymbol.SpecialType == SpecialType.System_DateTime)
        {
            return $"UnoMarshal.ToDateTime({paramName})";
        }

        // Arrays
        if (typeSymbol is IArrayTypeSymbol arrayType)
        {
            var elementTypeName = GetFullTypeName(arrayType.ElementType);
            return arrayType.Rank switch
            {
                1 => $"UnoMarshal.To1DArray<{elementTypeName}>({paramName})",
                2 => $"UnoMarshal.To2DArray<{elementTypeName}>({paramName})",
                _ => throw new NotSupportedException($"Arrays with rank {arrayType.Rank} are not supported.")
            };
        }

        // CalcRange
        if (typeSymbol.Name == "CalcRange" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
        {
            return $"UnoMarshal.ToCalcRange({paramName})";
        }

        // List<T>
        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } listType &&
            listType.Name == "List" &&
            listType.ContainingNamespace?.ToDisplayString() == "System.Collections.Generic")
        {
            var elementTypeName = GetFullTypeName(listType.TypeArguments[0]);
            return $"UnoMarshal.ToList<{elementTypeName}>({paramName})";
        }

        // TODO: Would be good to support DataTable, Dictionnary, etc.

        throw new NotSupportedException($"Marshaling for type '{typeSymbol.ToDisplayString()}' is not supported.");
    }

    private static string? GetComplexOptionalTypeMarshalingCode(IParameterSymbol parameter)
    {
        var typeSymbol = parameter.Type;
        var paramName = parameter.Name;

        // This will correctly check if type is nullable and return null if no default is specified
        var defaultLiteral = FormatDefaultValue(parameter);

        // Handle Nullable<T> by unwrapping
        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } nullableType &&
            nullableType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
        {
            typeSymbol = nullableType.TypeArguments[0];
        }

        // DateTime
        if (typeSymbol.SpecialType == SpecialType.System_DateTime)
        {
            return $"UnoMarshal.UnwrapOptionalDateTime({paramName}, {defaultLiteral})";
        }

        // Handle arrays
        if (typeSymbol is IArrayTypeSymbol arrayType)
        {
            var elementTypeName = GetFullTypeName(arrayType.ElementType);
            return arrayType.Rank switch
            {
                1 => $"UnoMarshal.UnwrapOptional1DArray<{elementTypeName}>({paramName}, {defaultLiteral})",
                2 => $"UnoMarshal.UnwrapOptional2DArray<{elementTypeName}>({paramName}, {defaultLiteral})",
                _ => throw new NotSupportedException($"Optional arrays with rank {arrayType.Rank} are not supported.")
            };
        }

        // Handle CalcRange
        if (typeSymbol.Name == "CalcRange" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
        {
            return $"UnoMarshal.UnwrapOptionalCalcRange({paramName}, {defaultLiteral})";
        }

        // Handle List<T>
        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } listType &&
            listType.Name == "List" &&
            listType.ContainingNamespace?.ToDisplayString() == "System.Collections.Generic")
        {
            var elementTypeName = GetFullTypeName(listType.TypeArguments[0]);
            return $"UnoMarshal.UnwrapOptionalList<{elementTypeName}>({paramName}, {defaultLiteral})";
        }
        // TODO: Would be good to support DataTable, Dictionnary, etc.

        return null;
    }   

    private static string? GetComplexNullableTypeMarshalingCode(IParameterSymbol parameter)
    {

        var typeSymbol = parameter.Type;
        var paramName = parameter.Name;

        // This will correctly check if type is nullable and return null if no default is specified
        var defaultLiteral = FormatDefaultValue(parameter);

        // Handle Nullable<T> by unwrapping
        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } nullableType &&
            nullableType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
        {
            typeSymbol = nullableType.TypeArguments[0];
        }

        // DateTime
        if (typeSymbol.SpecialType == SpecialType.System_DateTime)
        {
            return $"UnoMarshal.UnwrapNullableDateTime({paramName})";
        }

        // Handle arrays
        if (typeSymbol is IArrayTypeSymbol arrayType)
        {
            var elementTypeName = GetFullTypeName(arrayType.ElementType);
            return arrayType.Rank switch
            {
                1 => $"UnoMarshal.UnwrapNullable1DArray<{elementTypeName}>({paramName})",
                2 => $"UnoMarshal.UnwrapNullable2DArray<{elementTypeName}>({paramName})",
                _ => throw new NotSupportedException($"Nullable arrays with rank {arrayType.Rank} are not supported.")
            };
        }

        // Handle CalcRange
        if (typeSymbol.Name == "CalcRange" &&
            typeSymbol.ContainingNamespace?.ToDisplayString() == "CalcDNA.Runtime")
        {
            return $"UnoMarshal.UnwrapNullableCalcRange({paramName})";
        }

        // Handle List<T>
        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } listType &&
            listType.Name == "List" &&
            listType.ContainingNamespace?.ToDisplayString() == "System.Collections.Generic")
        {
            var elementTypeName = GetFullTypeName(listType.TypeArguments[0]);
            return $"UnoMarshal.UnwrapNullableList<{elementTypeName}>({paramName})";
        }
        // TODO: Would be good to support DataTable, Dictionnary, etc.

        return null;
    }

    private static string FormatDefaultValue(IParameterSymbol parameter)
    {
        // If no explicit default value, return type default string
        // Note that his handles nullables by returning "null"
        if (!parameter.HasExplicitDefaultValue)
            return GetTypeDefault(parameter.Type);

        // Actual provided default value
        var value = parameter.ExplicitDefaultValue;

        // Early exit for null
        if (value is null)
            return "null";

        // String version of the default value for the generated code
        // Only handle common types here and default to value.ToString() otherwise
        return parameter.Type.SpecialType switch
        {
            SpecialType.System_Boolean => value is true ? "true" : "false",
            SpecialType.System_String => $"\"{EscapeString((string)value)}\"",
            SpecialType.System_Char => $"'{EscapeChar((char)value)}'",
            SpecialType.System_Single => $"{value}f",
            SpecialType.System_Double => FormattableString.Invariant($"{value}"),
            SpecialType.System_Int64 => $"{value}L",
            _ => value.ToString() ?? GetTypeDefault(parameter.Type)
        };
    }

    private static string GetTypeDefault(ITypeSymbol typeSymbol)
    {
        // Handle Nullable<T>
        if (typeSymbol is INamedTypeSymbol { IsGenericType: true } nullableType &&
            nullableType.OriginalDefinition.SpecialType == SpecialType.System_Nullable_T)
        {
            return "null";
        }

        // Reference types default to null
        if (typeSymbol.IsReferenceType)
            return "null";

        // Arrays default to null
        if (typeSymbol is IArrayTypeSymbol)
            return "null";

        return typeSymbol.SpecialType switch
        {
            SpecialType.System_Boolean => "false",
            SpecialType.System_Byte => "(byte)0",
            SpecialType.System_SByte => "(sbyte)0",
            SpecialType.System_Int16 => "(short)0",
            SpecialType.System_Int32 => "0",
            SpecialType.System_Int64 => "0L",
            SpecialType.System_Single => "0f",
            SpecialType.System_Double => "0.0",
            SpecialType.System_Char => "'\\0'",
            _ => "default"
        };
    }

    // Use for generic types and arrays where we need the type in <> or before []
    private static string GetFullTypeName(ITypeSymbol typeSymbol)
    {
        // Use keyword names for primitives, full name for others
        return typeSymbol.SpecialType switch
        {
            SpecialType.System_Boolean => "bool",
            SpecialType.System_Byte => "byte",
            SpecialType.System_SByte => "sbyte",
            SpecialType.System_Int16 => "short",
            SpecialType.System_Int32 => "int",
            SpecialType.System_Int64 => "long",
            SpecialType.System_Single => "float",
            SpecialType.System_Double => "double",
            SpecialType.System_Char => "char",
            SpecialType.System_String => "string",
            SpecialType.System_Object => "object",
            _ => typeSymbol.ToDisplayString(SymbolDisplayFormat.FullyQualifiedFormat)
        };
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

    private static string EscapeChar(char value)
    {
        return value switch
        {
            '\\' => "\\\\",
            '\'' => "\\'",
            '\n' => "\\n",
            '\r' => "\\r",
            '\t' => "\\t",
            '\0' => "\\0",
            _ => value.ToString()
        };
    }
}
