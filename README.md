# Calc-DNA

[![Build Status](https://img.shields.io/github/actions/workflow/status/davidmarcombes/Calc-DNA/dotnet.yml?branch=main)](https://github.com/davidmarcombes/Calc-DNA/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/version-0.1.0-blue.svg)](https://github.com/davidmarcombes/Calc-DNA/releases)

A framework for building LibreOffice Calc add-ins in C#, inspired by Excel-DNA.

## Overview

Calc-DNA enables C# developers to create user-defined functions (UDFs) for LibreOffice Calc using a simple, attribute-based API. The framework abstracts the complexity of LibreOffice's UNO (Universal Network Objects) component model, providing a development experience similar to Excel-DNA.

The framework consists of three key components:
- **Compile-time code generation** using Roslyn source generators
- **Runtime type marshalling** between .NET and UNO types
- **CLI tooling** for generating LibreOffice extension metadata (IDL, XCU, RDB)

## Key Features

- **Simple API**: Attribute-based function decoration with `[CalcFunction]` and `[CalcParameter]`
- **Type Safety**: Automatic marshalling between .NET types (double, string, List<T>) and UNO types
- **Source Generation**: Roslyn-based code generators eliminate boilerplate and provide compile-time validation
- **CLI Tooling**: Automated generation of IDL, XCU, and RDB files required for LibreOffice extensions
- **Rich Interoperability**: Support for `CalcRange` objects to handle cell ranges efficiently
- **Cross-Platform**: Works on Windows, Linux, and macOS

## Architecture

Calc-DNA is designed around a three-stage pipeline that transforms simple C# methods into fully functional LibreOffice add-in components:

### 1. Compile-Time Code Generation (Roslyn Source Generators)

Two incremental source generators run during compilation:

- **CalcWrapperGenerator**: Generates wrapper methods that handle UNO type marshalling. For each `[CalcFunction]` method, it creates a corresponding wrapper that:
  - Accepts UNO-compatible types (`object[][]`, `double`, `string`)
  - Marshals inputs from UNO types to .NET types using `UnoMarshal`
  - Invokes the user's original method
  - Marshals the return value back to UNO types

- **UnoServiceGenerator**: Generates UNO service classes that implement the required LibreOffice interfaces (`XAddIn`, `XServiceInfo`, `XLocalizable`). These generated classes:
  - Expose function metadata (names, descriptions, categories)
  - Route function calls to the appropriate wrapper methods
  - Handle UNO component lifecycle and registration

### 2. Runtime Type Marshalling (CalcDNA.Runtime)

The runtime library provides:

- **UnoMarshal**: Bidirectional type conversion between .NET and UNO types:
  - Primitives: `double`, `int`, `string`, `bool`
  - Collections: `List<T>`, `T[]`
  - Cell ranges: `CalcRange` (wraps UNO `object[][]` with typed accessors)

- **CalcRange**: A strongly-typed wrapper around LibreOffice cell ranges that provides:
  - Enumeration over cell values
  - Row/column access
  - Type-safe value extraction

### 3. CLI Metadata Generation (CalcDNA.CLI)

The CLI tool processes compiled assemblies to generate LibreOffice extension metadata:

- **IDL Generator**: Creates `.idl` (Interface Definition Language) files that define the UNO component interfaces
- **XCU Generator**: Produces `.xcu` (Configuration Registry) files that register functions, specify categories, and provide localized descriptions
- **RDB Generator**: Invokes the LibreOffice SDK's `idlc` and `regmerge` tools to compile IDL into `.rdb` (Registry Database) files

The CLI tool uses reflection to discover `[CalcAddIn]` classes and `[CalcFunction]` methods, then generates all necessary metadata for LibreOffice to recognize and invoke the add-in.

### Data Flow

```
User's C# Code
    ↓
[CalcAddIn] + [CalcFunction] Attributes
    ↓
Roslyn Source Generators (compile-time)
    ├─→ Generated Wrapper Methods (type marshalling)
    └─→ Generated UNO Service Classes (XAddIn implementation)
    ↓
Compiled Assembly (.dll)
    ↓
CLI Tool (post-build)
    ├─→ .idl (interface definitions)
    ├─→ .xcu (function registry)
    └─→ .rdb (compiled type database)
    ↓
LibreOffice Extension (.oxt)
    ↓
User Functions Available in Calc
```

## Quick Start

Defining user-defined functions (UDFs) is straightforward. Calc-DNA handles complex types like ranges and lists automatically:

```csharp
using System.Linq;
using CalcDNA.Attributes;
using CalcDNA.Runtime;

[CalcAddIn(Name = "MyTools")]
public static class MathFunctions
{
    [CalcFunction(Description = "Calculates the sum of a range plus a list of extra values")]
    public static double ComplexSum(
        [CalcParameter(Description = "A range of cells")] CalcRange range,
        [CalcParameter(Description = "Extra values to add")] List<double> extras)
    {
        double sum = extras.Sum();
        
        foreach (var cell in range.Values())
        {
            if (cell is double d) sum += d;
        }
        
        return sum;
    }
}
```

## Repository Structure

The repository is organized into distinct projects, each with a specific responsibility:

```
Calc-DNA/
├── src/
│   ├── CalcDNA.Attributes/          # Attribute definitions
│   │   ├── CalcAddInAttribute.cs    # Marks a class as containing UDFs
│   │   ├── CalcFunctionAttribute.cs # Marks a method as a UDF
│   │   └── CalcParameterAttribute.cs # Provides parameter metadata
│   │
│   ├── CalcDNA.Generator/           # Roslyn source generators
│   │   ├── CalcWrapperGenerator.cs  # Generates type-marshalling wrappers
│   │   ├── UnoServiceGenerator.cs   # Generates UNO service classes
│   │   └── WrapperTypeMapping.cs    # Maps .NET types to UNO types
│   │
│   ├── CalcDNA.Runtime/             # Runtime support library
│   │   ├── CalcRange.cs             # Strongly-typed range wrapper
│   │   ├── UnoMarshal.cs            # Type conversion utilities
│   │   └── Uno/                     # UNO interface definitions
│   │       ├── IXAddIn.cs           # Add-in component interface
│   │       ├── IXServiceInfo.cs     # Service metadata interface
│   │       └── IXLocalizable.cs     # Localization interface
│   │
│   ├── CalcDNA.CLI/                 # Command-line metadata generator
│   │   ├── Program.cs               # CLI entry point
│   │   ├── IdlGenerator.cs          # Generates .idl files
│   │   ├── XcuGenerator.cs          # Generates .xcu files
│   │   ├── RdbGenerator.cs          # Invokes SDK tools to create .rdb
│   │   ├── UnoTypeMapping.cs        # Maps .NET types to UNO IDL types
│   │   └── Logger.cs                # CLI logging utilities
│   │
│   └── Demo.App/                    # Example add-in
│       ├── Functions.cs             # Basic function examples
│       ├── MoreFunctions.cs         # Range and list examples
│       ├── AdvancedFunctions.cs     # Advanced usage examples
│       └── Generated/               # Auto-generated wrapper code
│
├── tests/
│   ├── CalcDNA.Generator.Tests/     # Source generator tests
│   │   ├── CalcWrapperGeneratorTests.cs
│   │   └── WrapperTypeMappingTests.cs
│   │
│   └── CalcDNA.Runtime.Tests/       # Runtime library tests
│       ├── UnoMarshalTests.cs       # Type marshalling tests
│       └── CalcRangeTests.cs        # Range wrapper tests
│
├── README.md
├── DEVELOPMENT.md                   # Development environment setup
├── AGENTS.md                        # Guidelines for contributors
└── CalcDNA.slnx                     # Solution file
```

### Project Dependencies

```
Demo.App
  ├─→ CalcDNA.Attributes (reference)
  ├─→ CalcDNA.Runtime (reference)
  └─→ CalcDNA.Generator (analyzer/source generator)

CalcDNA.CLI
  ├─→ CalcDNA.Attributes (reference, for reflection)
  └─→ LibreOffice SDK tools (external, invoked via Process)

CalcDNA.Generator
  ├─→ CalcDNA.Attributes (reference, for symbol matching)
  ├─→ CalcDNA.Runtime (reference, for type information)
  └─→ Microsoft.CodeAnalysis.* (Roslyn APIs)

CalcDNA.Runtime
  └─→ (no internal dependencies, only LibreOffice CLI assemblies)

CalcDNA.Attributes
  └─→ (no dependencies)
```

## Building the Project

Build the entire solution:

```bash
dotnet build
```

This will:
1. Compile all projects in dependency order
2. Run Roslyn source generators during `Demo.App` compilation (generating wrapper and service code)
3. Produce assemblies in each project's `bin/` directory

## Using the CLI Tool

The CLI tool generates the LibreOffice extension metadata required to register your add-in:

```bash
dotnet run --project src/CalcDNA.CLI -- <AssemblyPath> [OutputPath] [AddInName] [SDKPath]
```

### Arguments

- `AssemblyPath` (required): Path to your compiled add-in assembly (.dll)
- `OutputPath` (optional): Directory for generated files (defaults to assembly directory)
- `AddInName` (optional): Name for the add-in (defaults to assembly name)
- `SDKPath` (optional): Path to LibreOffice SDK (auto-detected if not specified)

### Options

- `-v, --verbose`: Enable detailed output for debugging

### Example

```bash
dotnet run --project src/CalcDNA.CLI -- ./Demo.App/bin/Debug/net10.0/Demo.App.dll --output ./output --verbose
```

### Generated Files

The CLI generates three types of files:

- **`.idl`**: Interface Definition Language file defining the UNO component interfaces
- **`.xcu`**: XML Configuration Unit file registering functions and metadata with LibreOffice
- **`.rdb`**: Registry Database (compiled IDL) used by LibreOffice's component loader

These files are required for LibreOffice to discover and invoke your add-in functions.

## Testing

The project uses xUnit for testing with two test suites:

### CalcDNA.Generator.Tests

Tests for Roslyn source generators using the `Microsoft.CodeAnalysis.Testing` framework:

- **CalcWrapperGeneratorTests**: Verifies that source generators produce correct wrapper code for various scenarios (simple types, ranges, lists, error conditions)
- **WrapperTypeMappingTests**: Validates type mapping logic between .NET and UNO types

These tests use snapshot-style verification, comparing generated source code against expected output.

### CalcDNA.Runtime.Tests

Tests for runtime type marshalling and range handling:

- **UnoMarshalTests**: Validates bidirectional marshalling for primitives, collections, and complex types
- **CalcRangeTests**: Tests range wrapper functionality (enumeration, indexing, type conversion)

### Running Tests

```bash
# Run all tests
dotnet test

# Run with detailed output
dotnet test --logger "console;verbosity=detailed"

# Run specific test project
dotnet test tests/CalcDNA.Generator.Tests
dotnet test tests/CalcDNA.Runtime.Tests
```

### Testing Strategy

- **Unit tests**: Core logic without external dependencies (type mapping, marshalling)
- **Generator tests**: Verify source generation produces correct code
- **Integration tests**: (planned) End-to-end tests with actual LibreOffice instance

Integration testing with LibreOffice requires a running LibreOffice instance and is not yet automated. Manual integration testing is performed using the [Demo.App](src/Demo.App) project.

## Project Status and Roadmap

### Completed

- [x] Core attribute API (`[CalcAddIn]`, `[CalcFunction]`, `[CalcParameter]`)
- [x] Roslyn source generators for wrapper and service code generation
- [x] Runtime type marshalling (primitives, strings, lists, ranges)
- [x] CLI tool with IDL, XCU, and RDB generation
- [x] `CalcRange` wrapper for cell ranges
- [x] Unit and generator test suites

### In Progress

- [ ] Error handling and diagnostics improvements
- [ ] Enhanced logging and debugging support
- [ ] Documentation and inline code comments

### Planned

- [ ] **Automated `.oxt` packaging**: Generate complete LibreOffice extension packages
- [ ] **Enhanced type support**: Dictionaries, nullable types, custom objects
- [ ] **Attribute enhancements**: Category, volatile functions, help URLs
- [ ] **Integration test automation**: Automated testing against LibreOffice
- [ ] **Performance optimization**: Caching, lazy initialization, reduced allocations
- [ ] **Async support**: Asynchronous function execution (requires UNO threading investigation)
- [ ] **Developer tooling**: VS Code extension, debugging helpers, live reload

### Known Limitations

- **No async/await support**: All functions must be synchronous
- **Limited error context**: UNO exceptions lose some context during marshalling
- **Manual packaging**: `.oxt` extension packages must be created manually
- **No COM interop**: Unlike Excel-DNA, there is no equivalent to Excel's COM automation model in LibreOffice

## Documentation

- [DEVELOPMENT.md](DEVELOPMENT.md) - Development environment setup, dependencies, and build instructions
- [AGENTS.md](AGENTS.md) - Guidelines for contributors and AI assistants
- [LibreOffice SDK Documentation](https://api.libreoffice.org/) - Official LibreOffice API reference
- [UNO/CLI Language Binding](https://wiki.openoffice.org/wiki/Uno/Cli) - Information on C# UNO bindings

## Contributing

Contributions are welcome in the form of bug reports, feature requests, and pull requests. When contributing:

1. Fork the repository and create a feature branch
2. Follow the coding conventions outlined in [AGENTS.md](AGENTS.md)
3. Add tests for new functionality
4. Ensure all tests pass with `dotnet test`
5. Submit a pull request with a clear description of changes

### Areas Needing Attention

- Integration testing with LibreOffice instances
- Error handling and diagnostic improvements
- Documentation and code examples
- Cross-platform testing (Windows, Linux, macOS)
- Performance profiling and optimization

## License

This project is licensed under the MIT License. See the LICENSE file for details (pending).

## Acknowledgments

- **Excel-DNA** - Inspiration for the attribute-based API design and developer experience
- **LibreOffice Project** - For the UNO component framework and extensive SDK
- **.NET Foundation** - For Roslyn compiler APIs that enable source generation
- The open-source community for making cross-platform development possible

