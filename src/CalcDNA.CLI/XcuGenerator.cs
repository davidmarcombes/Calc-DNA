using System.Reflection;
using System.Xml.Linq;

public static class XcuGenerator
{
    private static readonly XNamespace Oor = "http://openoffice.org/2001/registry";

    public static string BuildXcu(Type type, IEnumerable<MethodInfo> methods)
    {
        // Use the implementation name (this must match your C# class full name)
        string implementationName = type.FullName ?? type.Name;

        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(Oor + "component-data",
                new XAttribute(Oor + "name", "CalcAddIns"),
                new XAttribute(Oor + "package", "org.openoffice.Office"),
                new XElement("node", new XAttribute(Oor + "name", "AddInInfo"),
                    new XElement("node",
                        new XAttribute(Oor + "name", implementationName),
                        new XAttribute(Oor + "op", "replace"),
                        new XElement("node", new XAttribute(Oor + "name", "AddInFunctions"),
                            methods.Select(CreateFunctionNode)
                        )
                    )
                )
            )
        );

        return doc.ToString();
    }

    private static XElement CreateFunctionNode(MethodInfo method)
    {
        // Extract attribute data (using dynamic or casting since MLC is used)
        var funcAttr = method.GetCustomAttributesData()
            .FirstOrDefault(a => a.AttributeType.Name == "CalcFunctionAttribute");

        string desc = funcAttr?.ConstructorArguments[0].Value?.ToString() ?? "";
        string category = funcAttr?.NamedArguments
            .FirstOrDefault(a => a.MemberName == "Category").TypedValue.Value?.ToString() ?? "Add-In";

        return new XElement("node",
            new XAttribute(Oor + "name", method.Name),
            new XAttribute(Oor + "op", "replace"),
            new XElement("prop", new XAttribute(Oor + "name", "DisplayName"),
                new XElement("value", method.Name)),
            new XElement("prop", new XAttribute(Oor + "name", "Description"),
                new XElement("value", desc)),
            new XElement("prop", new XAttribute(Oor + "name", "Category"),
                new XElement("value", category)),
            new XElement("node", new XAttribute(Oor + "name", "Parameters"),
                method.GetParameters().Select(CreateParameterNode))
        );
    }

    private static XElement CreateParameterNode(ParameterInfo param)
    {
        var paramAttr = param.GetCustomAttributesData()
            .FirstOrDefault(a => a.AttributeType.Name == "CalcParameterAttribute");

        string desc = paramAttr?.ConstructorArguments[0].Value?.ToString() ?? "";

        return new XElement("node",
            new XAttribute(Oor + "name", param.Name),
            new XAttribute(Oor + "op", "replace"),
            new XElement("prop", new XAttribute(Oor + "name", "DisplayName"),
                new XElement("value", param.Name)),
            new XElement("prop", new XAttribute(Oor + "name", "Description"),
                new XElement("value", desc))
        );
    }
}