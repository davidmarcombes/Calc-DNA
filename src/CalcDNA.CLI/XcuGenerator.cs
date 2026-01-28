using System.Reflection;
using System.Xml.Linq;
using CalcDNA.CLI;

public static class XcuGenerator
{
    private static readonly XNamespace Oor = "http://openoffice.org/2001/registry";

    public static string BuildXcu(IEnumerable<AddInClass> addInClasses, Logger logger)
    {
        var addInNodes = new List<XElement>();

        foreach (var addIn in addInClasses)
        {
            try
            {
                addInNodes.Add(CreateAddInNode(addIn, logger));
            }
            catch (Exception ex)
            {
                logger.Warning($"Skipping class '{addIn.Type.Name}': {ex.Message}");
            }
        }

        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(Oor + "component-data",
                new XAttribute(Oor + "name", "CalcAddIns"),
                new XAttribute(Oor + "package", "org.openoffice.Office"),
                new XElement("node", new XAttribute(Oor + "name", "AddInInfo"),
                    addInNodes
                )
            )
        );

        return doc.ToString();
    }

    private static XElement CreateAddInNode(AddInClass addIn, Logger logger)
    {
        string implementationName = addIn.Type.FullName ?? addIn.Type.Name;

        var functionNodes = new List<XElement>();
        foreach (var method in addIn.Methods)
        {
            try
            {
                functionNodes.Add(CreateFunctionNode(method, logger));
            }
            catch (Exception ex)
            {
                logger.Warning($"Skipping method '{method.Name}': {ex.Message}");
            }
        }

        return new XElement("node",
            new XAttribute(Oor + "name", implementationName),
            new XAttribute(Oor + "op", "replace"),
            new XElement("node", new XAttribute(Oor + "name", "AddInFunctions"),
                functionNodes
            )
        );
    }

    private static XElement CreateFunctionNode(MethodInfo method, Logger logger)
    {
        // Extract attribute data (using dynamic or casting since MLC is used)
        var funcAttr = method.GetCustomAttributesData()
            .FirstOrDefault(a => a.AttributeType.Name == "CalcFunctionAttribute");

        string displayName = method.Name;
        string desc = "";
        string category = "Add-In";

        if (funcAttr != null)
        {
            // Check named arguments for all properties
            foreach (var namedArg in funcAttr.NamedArguments)
            {
                switch (namedArg.MemberName)
                {
                    case "Name":
                        displayName = namedArg.TypedValue.Value?.ToString() ?? method.Name;
                        break;
                    case "Description":
                        desc = namedArg.TypedValue.Value?.ToString() ?? "";
                        break;
                    case "Category":
                        category = namedArg.TypedValue.Value?.ToString() ?? "Add-In";
                        break;
                }
            }

            // Also check constructor arguments (if name was passed via constructor)
            if (funcAttr.ConstructorArguments.Count > 0)
            {
                displayName = funcAttr.ConstructorArguments[0].Value?.ToString() ?? displayName;
            }
        }

        return new XElement("node",
            new XAttribute(Oor + "name", method.Name),
            new XAttribute(Oor + "op", "replace"),
            new XElement("prop", new XAttribute(Oor + "name", "DisplayName"),
                new XElement("value", displayName)),
            new XElement("prop", new XAttribute(Oor + "name", "Description"),
                new XElement("value", desc)),
            new XElement("prop", new XAttribute(Oor + "name", "Category"),
                new XElement("value", category)),
            new XElement("node", new XAttribute(Oor + "name", "Parameters"),
                method.GetParameters().Select(p => CreateParameterNode(p, logger)))
        );
    }

    private static XElement CreateParameterNode(ParameterInfo param, Logger logger)
    {
        var paramAttr = param.GetCustomAttributesData()
            .FirstOrDefault(a => a.AttributeType.Name == "CalcParameterAttribute");

        string desc = "";
        if (paramAttr != null && paramAttr.ConstructorArguments.Count > 0)
        {
            desc = paramAttr.ConstructorArguments[0].Value?.ToString() ?? "";
        }

        string paramName = param.Name ?? $"arg{param.Position}";

        return new XElement("node",
            new XAttribute(Oor + "name", paramName),
            new XAttribute(Oor + "op", "replace"),
            new XElement("prop", new XAttribute(Oor + "name", "DisplayName"),
                new XElement("value", paramName)),
            new XElement("prop", new XAttribute(Oor + "name", "Description"),
                new XElement("value", desc))
        );
    }
}