using System.Reflection;
using System.Xml.Linq;
using CalcDNA.CLI;
using System.Collections.Generic;
using System.Linq;
using System.Text;

/// <summary>
/// XCU (XML Configuration) generator for LibreOffice Calc add-ins.
/// </summary>
internal static class XcuGenerator
{
    private static readonly XNamespace Oor = "http://openoffice.org/2001/registry";
    private static readonly XNamespace Xml = XNamespace.Xml;

    /// <summary>
    /// Build the XCU file content.
    /// </summary>
    public static string BuildXcu(string moduleName, IEnumerable<AddInClass> addInClasses, Logger logger)
    {
        var addInNodes = new List<XElement>();

        foreach (var addIn in addInClasses)
        {
            try
            {
                var node = CreateAddInNode(moduleName, addIn, logger);
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

        // IMPORTANT: LibreOffice .xcu files expect elements (node, prop, value) 
        // to be in the empty namespace, but attributes (name, op) to be in the 'oor' namespace.
        // The root component-data element IS in the 'oor' namespace.
        var doc = new XDocument(
            new XDeclaration("1.0", "UTF-8", null),
            new XElement(Oor + "component-data",
                new XAttribute(XNamespace.Xmlns + "oor", Oor.NamespaceName),
                new XAttribute(Oor + "name", "CalcAddIns"),
                new XAttribute(Oor + "package", "org.openoffice.Office"),
                new XElement("node", new XAttribute(Oor + "name", "AddInInfo"),
                    addInNodes
                )
            )
        );

        return doc.ToString();
    }

    private static XElement? CreateAddInNode(string moduleName, AddInClass addIn, Logger logger)
    {
        // Use the UNO module name (sanitized) + class name (sanitized)
        string sanitizedClassName = SanitizeIdentifier(addIn.Type.Name);
        string implementationName = $"{moduleName}.{sanitizedClassName}";

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
        // Extract attribute data
        var funcAttr = method.GetCustomAttributesData()
            .FirstOrDefault(a => a.AttributeType.Name == "CalcFunctionAttribute");

        string displayName = method.Name;
        string desc = "";
        string category = "Add-In";
        string compatName = "";

        if (funcAttr != null)
        {
            if (funcAttr.ConstructorArguments.Count > 0)
            {
                displayName = funcAttr.ConstructorArguments[0].Value?.ToString() ?? displayName;
            }

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

        string sanitizedMethodName = SanitizeIdentifier(method.Name);

        var node = new XElement("node",
            new XAttribute(Oor + "name", sanitizedMethodName),
            new XAttribute(Oor + "op", "replace"),
            new XElement("prop", new XAttribute(Oor + "name", "DisplayName"),
                new XElement("value", new XAttribute(Xml + "lang", "en"), displayName)),
            new XElement("prop", new XAttribute(Oor + "name", "Description"),
                new XElement("value", new XAttribute(Xml + "lang", "en"), desc)),
            new XElement("prop", new XAttribute(Oor + "name", "Category"),
                new XElement("value", new XAttribute(Xml + "lang", "en"), category))
        );

        if (!string.IsNullOrEmpty(compatName))
        {
            node.Add(new XElement("prop", new XAttribute(Oor + "name", "CompatibilityName"),
                new XElement("value", new XAttribute(Xml + "lang", "en"), compatName)));
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
            if (paramAttr.ConstructorArguments.Count > 0)
            {
                displayName = paramAttr.ConstructorArguments[0].Value?.ToString() ?? displayName;
            }

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

        string sanitizedParamName = SanitizeIdentifier(param.Name ?? $"arg{param.Position}");

        return new XElement("node",
            new XAttribute(Oor + "name", sanitizedParamName),
            new XAttribute(Oor + "op", "replace"),
            new XElement("prop", new XAttribute(Oor + "name", "DisplayName"),
                new XElement("value", new XAttribute(Xml + "lang", "en"), displayName)),
            new XElement("prop", new XAttribute(Oor + "name", "Description"),
                new XElement("value", new XAttribute(Xml + "lang", "en"), desc))
        );
    }

    private static readonly HashSet<string> ReservedWords = new(StringComparer.OrdinalIgnoreCase)
    {
        "module", "interface", "service", "any", "boolean", "byte", "char",
        "double", "enum", "exception", "FALSE", "float", "hyper", "long",
        "octet", "sequence", "short", "string", "struct", "TRUE", "type",
        "typedef", "union", "unsigned", "void", "in", "out", "inout"
    };

    private static string SanitizeIdentifier(string name)
    {
        if (string.IsNullOrEmpty(name)) return "id";

        var sb = new StringBuilder();
        foreach (char c in name)
        {
            if (char.IsLetterOrDigit(c) || c == '_')
            {
                sb.Append(c);
            }
            else
            {
                sb.Append('_');
            }
        }

        string result = sb.ToString().Trim('_');
        if (string.IsNullOrEmpty(result)) result = "id";

        if (char.IsDigit(result[0]))
        {
            result = "_" + result;
        }

        if (ReservedWords.Contains(result))
        {
            result = "_" + result;
        }

        return result;
    }
}