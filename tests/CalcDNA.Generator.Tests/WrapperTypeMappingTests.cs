using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp;

namespace CalcDNA.Generator.Tests;

public class WrapperTypeMappingTests
{
    #region Helper Methods

    private static ITypeSymbol GetTypeSymbol(string typeName)
    {
        var compilation = CSharpCompilation.Create("Test",
            references: new[] { MetadataReference.CreateFromFile(typeof(object).Assembly.Location) });

        return typeName switch
        {
            "bool" => compilation.GetSpecialType(SpecialType.System_Boolean),
            "byte" => compilation.GetSpecialType(SpecialType.System_Byte),
            "sbyte" => compilation.GetSpecialType(SpecialType.System_SByte),
            "short" => compilation.GetSpecialType(SpecialType.System_Int16),
            "int" => compilation.GetSpecialType(SpecialType.System_Int32),
            "long" => compilation.GetSpecialType(SpecialType.System_Int64),
            "float" => compilation.GetSpecialType(SpecialType.System_Single),
            "double" => compilation.GetSpecialType(SpecialType.System_Double),
            "char" => compilation.GetSpecialType(SpecialType.System_Char),
            "string" => compilation.GetSpecialType(SpecialType.System_String),
            "object" => compilation.GetSpecialType(SpecialType.System_Object),
            _ => throw new ArgumentException($"Unknown type: {typeName}")
        };
    }

    #endregion

    #region MapTypeToWrapper Tests

    [Theory]
    [InlineData("bool", false, "bool")]
    [InlineData("byte", false, "byte")]
    [InlineData("short", false, "short")]
    [InlineData("int", false, "int")]
    [InlineData("long", false, "long")]
    [InlineData("float", false, "float")]
    [InlineData("double", false, "double")]
    [InlineData("string", false, "string")]
    public void MapTypeToWrapper_WithPrimitives_ReturnsPrimitiveType(string inputType, bool optional, string expectedOutput)
    {
        var typeSymbol = GetTypeSymbol(inputType);
        var result = WrapperTypeMapping.MapTypeToWrapper(typeSymbol, optional);
        Assert.Equal(expectedOutput, result);
    }

    [Theory]
    [InlineData("bool")]
    [InlineData("int")]
    [InlineData("double")]
    [InlineData("string")]
    public void MapTypeToWrapper_WithOptionalPrimitives_ReturnsObjectNullable(string inputType)
    {
        var typeSymbol = GetTypeSymbol(inputType);
        var result = WrapperTypeMapping.MapTypeToWrapper(typeSymbol, optional: true);
        Assert.Equal("object?", result);
    }

    [Fact]
    public void MapTypeToWrapper_WithSByte_MapsToShort()
    {
        // SByte maps to short to avoid UNO signedness issues
        var typeSymbol = GetTypeSymbol("sbyte");
        var result = WrapperTypeMapping.MapTypeToWrapper(typeSymbol, optional: false);
        Assert.Equal("short", result);
    }

    #endregion

    #region IsTypeMappedToObject Tests

    [Theory]
    [InlineData("bool", false, false)]
    [InlineData("int", false, false)]
    [InlineData("double", false, false)]
    [InlineData("string", false, false)]
    public void IsTypeMappedToObject_WithNonOptionalPrimitives_ReturnsFalse(string inputType, bool optional, bool expected)
    {
        var typeSymbol = GetTypeSymbol(inputType);
        var result = WrapperTypeMapping.IsTypeMappedToObject(typeSymbol, optional);
        Assert.Equal(expected, result);
    }

    [Theory]
    [InlineData("bool", true, true)]
    [InlineData("int", true, true)]
    [InlineData("double", true, true)]
    public void IsTypeMappedToObject_WithOptionalPrimitives_ReturnsTrue(string inputType, bool optional, bool expected)
    {
        var typeSymbol = GetTypeSymbol(inputType);
        var result = WrapperTypeMapping.IsTypeMappedToObject(typeSymbol, optional);
        Assert.Equal(expected, result);
    }

    #endregion
}
