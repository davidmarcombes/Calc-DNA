using Microsoft.CodeAnalysis;
using Microsoft.CodeAnalysis.CSharp.Syntax;
using Microsoft.CodeAnalysis.Text;
using System.Linq;
using System.Text;

namespace CalcDNA.Generator;

[Generator]
public class CalcWrapperGenerator : IIncrementalGenerator
{
    const string sStatic = "static";
    const string sPublic = "public";

    public void Initialize(IncrementalGeneratorInitializationContext context)
    {
        // STEP 1: Filter - Find methods with [CalcFunction] in classes with [CalcAddIn]
        var methodDeclarations = context.SyntaxProvider
            .CreateSyntaxProvider(
                predicate: (node, _) => IsTargetMethod(node),
                transform: (ctx, _) => GetSemanticTarget(ctx))
            .Where(m => m is not null);

        // STEP 2: Output - Generate the wrapper code
        context.RegisterSourceOutput(methodDeclarations, (spc, method) => {
            if (method is null)
                return;

            var source = GenerateWrapperSource(method);
            spc.AddSource($"{method.ContainingType.Name}_{method.Name}_Wrapper.g.cs", SourceText.From(source, Encoding.UTF8));
        });
    }

    private static bool IsTargetMethod(SyntaxNode node)
    {
        // Method used for syntactic filter
        // Fast, lightweight check on the syntax tree

        // Check if node is a method with any attributes
        if (node is not MethodDeclarationSyntax method || method.AttributeLists.Count == 0)
            return false;

        // Check if containing class exists and has any attributes
        if (method.Parent is not ClassDeclarationSyntax classDeclaration || classDeclaration.AttributeLists.Count == 0)
            return false;

        // Check if method is public and static
        var modifiers = method.Modifiers;
        if (!modifiers.Any(m => m.ValueText == sPublic) || !modifiers.Any(m => m.ValueText == sStatic))
            return false;

        // Check if class is public and static
        var classModifiers = classDeclaration.Modifiers;
        if (!classModifiers.Any(m => m.ValueText == sPublic) || !classModifiers.Any(m => m.ValueText == sStatic))
            return false;

        return true;
    }

    private static IMethodSymbol? GetSemanticTarget(GeneratorSyntaxContext context)
    {
        // Method used for semantic filter
        // More expensive, but provides full semantic information

        var method = (MethodDeclarationSyntax)context.Node;
        if (context.SemanticModel.GetDeclaredSymbol(method) is not IMethodSymbol methodSymbol)
            return null;

        // Check if method has [CalcFunction] attribute
        var hasCalcFunction = methodSymbol.GetAttributes()
            .Any(attr => attr.AttributeClass?.Name == "CalcFunctionAttribute" ||
                        attr.AttributeClass?.Name == "CalcFunction");

        if (!hasCalcFunction)
            return null;

        // Check if containing class has [CalcAddIn] attribute
        var containingClass = methodSymbol.ContainingType;
        var hasCalcAddIn = containingClass.GetAttributes()
            .Any(attr => attr.AttributeClass?.Name == "CalcAddInAttribute" ||
                        attr.AttributeClass?.Name == "CalcAddIn");

        if (!hasCalcAddIn)
            return null;

        return methodSymbol;
    }

    private static string GenerateWrapperSource(IMethodSymbol method)
    {

        var returnType = WrapperTypeMapping.MapTypeToWrapper(method.ReturnType, false);
        var parameters = string.Join(", ", method.Parameters.Select(p =>
        {
            return $"{WrapperTypeMapping.MapTypeToWrapper(p.Type, WrapperTypeMapping.IsOptionalParameter(p))} {p.Name}";
        }));

        // Marshal parameters if needed
        StringBuilder paramMarschal = new StringBuilder();
        StringBuilder paramLList = new StringBuilder();
        string sep = "";
        string tab = "";
        foreach ( var parameter in method.Parameters ) {
            // Optional parmeters must be unwrapped
            if( WrapperTypeMapping.NeedsMarshaling( parameter ))
            {
                paramMarschal.AppendLine( $"{tab}var marshaled_{parameter.Name} = {WrapperTypeMapping.GetMarshalingCode( parameter )};" );
                paramLList.Append( $"{sep}marshaled_{parameter.Name}" );
                tab = "                ";
            } else {
                paramLList.Append( $"{sep}{parameter.Name}" );
            }
            sep= ", ";
            


        }


        // TODO: Should throw a UNO exception instead of a C# exception

        // This is where you write the C# text for the wrapper class
        return $@"
using CalcDNA.Runtime;
namespace {method.ContainingNamespace} {{
    public partial class {method.ContainingType.Name} {{
        // Your generated UNO-compatible wrapper method goes here
        // Method: {method.Name}
        public static {returnType} {method.Name}_UNOWrapper({parameters})
        {{
            try 
            {{
                {paramMarschal}
                return {method.Name}({paramLList});
            }}
            catch (Exception ex)
            {{
                    throw new Exception($"Error calling {method.Name}: {ex.Message}");
            }}
        }}
    }}
}}";
    }
}