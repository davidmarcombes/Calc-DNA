using System.Reflection;
using System.Xml.Linq;
using CalcDNA.CLI;
using System.Collections.Generic;
using System.Linq;

/// <summary>
/// XCU (XML Configuration) generator for LibreOffice Calc add-ins.
/// </summary>
internal static class XcuGenerator
{
    private static readonly XNamespace Oor = "http://openoffice.org/2001/registry";

    /// <summary>
    /// Build the XCU file content.
    /// </summary>
    public static string BuildXcu(IEnumerable<AddInClass> addInClasses, Logger logger)
    {
        var addInNodes = new List<XElement>();

        foreach (var addIn in addInClasses)
        {
            try
            {
                var node = CreateAddInNode(addIn, logger);
                if (node != null)
                {
                    addInNodes.Add(node);
                }
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

    private static XElement? CreateAddInNode(AddInClass addIn, Logger logger)
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
                logger.Warning($"Skipping method '{method.Name}' in {addIn.Type.Name}: {ex.Message}");
            }
        }

        if (!functionNodes.Any()) return null;

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
        string compatName = "";

        if (funcAttr != null)
        {
            // Check constructor arguments (if name was passed via constructor)
            if (funcAttr.ConstructorArguments.Count > 0)
            {
                displayName = funcAttr.ConstructorArguments[0].Value?.ToString() ?? displayName;
            }

            // Check named arguments for all properties
            foreach (var namedArg in funcAttr.NamedArguments)
            {
                switch (namedArg.MemberName)
                {
                    case "Name":
                        displayName = namedArg.TypedValue.Value?.ToString() ?? displayName;
                        break;
                    case "Description":
                        desc = namedArg.TypedValue.Value?.ToString() ?? "";
                        break;
                    case "Category":
                        category = namedArg.TypedValue.Value?.ToString() ?? "Add-In";
                        break;
                    case "CompatibilityName":
                        compatName = namedArg.TypedValue.Value?.ToString() ?? "";
                        break;
                }
            }
        }

        var node = new XElement("node",
            new XAttribute(Oor + "name", method.Name),
            new XAttribute(Oor + "op", "replace"),
            new XElement("prop", new XAttribute(Oor + "name", "DisplayName"),
                new XElement("value", displayName)),
            new XElement("prop", new XAttribute(Oor + "name", "Description"),
                new XElement("value", desc)),
            new XElement("prop", new XAttribute(Oor + "name", "Category"),
                new XElement("value", category))
        );

        if (!string.IsNullOrEmpty(compatName))
        {
            node.Add(new XElement("prop", new XAttribute(Oor + "name", "CompatibilityName"),
                new XElement("value", compatName)));
        }

        node.Add(new XElement("node", new XAttribute(Oor + "name", "Parameters"),
                method.GetParameters().Select(p => CreateParameterNode(p, logger))));

        return node;
    }

    private static XElement CreateParameterNode(ParameterInfo param, Logger logger)
    {
        var paramAttr = param.GetCustomAttributesData()
            .FirstOrDefault(a => a.AttributeType.Name == "CalcParameterAttribute");

        string displayName = param.Name ?? $"arg{param.Position}";
        string desc = "";

        if (paramAttr != null)
        {
             // Check constructor arguments
            if (paramAttr.ConstructorArguments.Count > 0)
            {
                displayName = paramAttr.ConstructorArguments[0].Value?.ToString() ?? displayName;
            }

            // Check named arguments
            foreach (var namedArg in paramAttr.NamedArguments)
            {
                switch (namedArg.MemberName)
                {
                    case "Name":
                        displayName = namedArg.TypedValue.Value?.ToString() ?? displayName;
                        break;
                    case "Description":
                        desc = namedArg.TypedValue.Value?.ToString() ?? "";
                        break;
                }
            }
        }

        string paramInternalName = param.Name ?? $"arg{param.Position}";

        return new XElement("node",
            new XAttribute(Oor + "name", paramInternalName),
            new XAttribute(Oor + "op", "replace"),
            new XElement("prop", new XAttribute(Oor + "name", "DisplayName"),
                new XElement("value", displayName)),
            new XElement("prop", new XAttribute(Oor + "name", "Description"),
                new XElement("value", desc))
        );
    }
}