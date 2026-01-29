# Agent Instructions for Calc-DNA Development

This document provides guidelines for AI coding assistants and developers working on the Calc-DNA project.

## Project Context

Calc-DNA is a framework for building LibreOffice Calc add-ins using C#, inspired by Excel-DNA. The goal is to provide a simple, attribute-based API that hides the complexity of LibreOffice's UNO component model.

**Key Principles:**
- Keep the user-facing API simple and familiar (Excel-DNA-like)
- Hide UNO complexity in the framework
- Support cross-platform development (Windows, Linux, macOS)
- Prioritize clarity and maintainability over cleverness

## Technology Stack

- **Language**: C#
- **Target Platform**: LibreOffice 7.0+ via UNO (Universal Network Objects)
- **Build System**: MSBuild (.NET SDK-style projects)
- **IDE**: VS Code with C# Dev Kit or Visual Studio 2022+
- **Testing**: xUnit 

## Code Style and Conventions

### C# Style Guidelines

Follow modern C# conventions.

**File Organization:**
- One primary type per file
- File name matches the primary type name
- Organize using folders/namespaces that reflect purpose

### Documentation

**XML Documentation:**
All public APIs must have XML documentation

**Inline Comments:**
Use for non-obvious logic, UNO-specific quirks, or workarounds

### Async/Threading

**Current Scope:**
- Initial version will be **synchronous only**
- Calc's calculation model is primarily single-threaded
- UNO components have specific threading requirements

## UNO-Specific Guidelines

### UNO Naming Conventions

UNO uses different conventions than .NET:

```csharp
// UNO interfaces start with X
using unoidl.com.sun.star.sheet.XAddIn;
using unoidl.com.sun.star.lang.XServiceInfo;

// Wrap UNO calls with try-catch (UNO can throw exceptions)
public string GetFunctionName(string programmaticName)
{
    try
    {
        // UNO method names are camelCase
        return _unoComponent.getDisplayFunctionName(programmaticName);
    }
    catch (unoidl.com.sun.star.uno.Exception ex)
    {
        throw new UnoException("Failed to get function name", ex);
    }
}
```

### Type Marshalling

Be explicit about .NET ↔ UNO type conversions:

### Component Registration

UNO registration is non-intuitive; document it well

## Testing Guidelines

### Unit Tests

Test business logic without UNO dependencies

### Integration Tests

Test with actual UNO components

### Version Compatibility

Plan for future changes:

```csharp
[AttributeUsage(AttributeTargets.Method)]
public class CalcFunctionAttribute : Attribute
{
    // Version 1 properties
    public string? Name { get; set; }
    public string? Description { get; set; }
    public string? Category { get; set; }

    // Reserved for future use - don't implement yet
    // public bool IsVolatile { get; set; }
    // public bool IsAsync { get; set; }
    // public bool IsThreadSafe { get; set; }
}
```

## Common Pitfalls

### UNO String Handling
```csharp
// ❌ Wrong - UNO strings are not null-terminated like C strings
string GetFunctionName() => _name + "\0";

// ✅ Correct - UNO handles strings properly
string GetFunctionName() => _name;
```

### Case Sensitivity
```csharp
// UNO is case-sensitive for service/interface names
// ❌ Wrong
const string SERVICE = "com.sun.star.sheet.addin";

// ✅ Correct - exact casing matters
const string SERVICE = "com.sun.star.sheet.AddIn";
```

### Assembly Loading
```csharp
// ❌ Wrong - may not find assemblies in UNO context
var asm = Assembly.Load("MyFunctions");

// ✅ Correct - use executing assembly or full path
var asm = Assembly.GetExecutingAssembly();
```

## Resources

- [LibreOffice UNO/CLI](https://wiki.openoffice.org/wiki/Uno/Cli)

