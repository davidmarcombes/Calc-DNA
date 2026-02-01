# Calc-DNA

[![Build Status](https://img.shields.io/github/actions/workflow/status/davidmarcombes/Calc-DNA/dotnet.yml?branch=main)](https://github.com/davidmarcombes/Calc-DNA/actions)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Version](https://img.shields.io/badge/version-0.1.0-blue.svg)](https://github.com/davidmarcombes/Calc-DNA/releases)

A framework for building LibreOffice Calc add-ins in C#, inspired by Excel-DNA.

## Overview

Calc-DNA enables C# developers to create user-defined functions (UDFs) for LibreOffice Calc using a simple, attribute-based API. The framework abstracts the complexity of LibreOffice's UNO (Universal Network Objects) component model, providing a development experience similar to Excel-DNA.

The framework consists of four key components:
- **Compile-time code generation** using Roslyn source generators
- **Runtime type marshalling** between .NET and UNO types
- **CLI tooling** for generating LibreOffice extension metadata and packaging
- **Python UNO bridge** for cross-platform deployment without a .NET UNO bridge

## Key Features

- **Simple API**: Attribute-based function decoration with `[CalcFunction]` and `[CalcParameter]`
- **Type Safety**: Automatic marshalling between .NET types (`double`, `string`, `List<T>`) and UNO types via `UnoMarshal` and `PyMarshal`
- **Source Generation**: Roslyn-based code generators eliminate boilerplate and provide compile-time validation
- **CLI Tooling**: Automated generation of IDL, XCU, RDB files, and complete `.oxt` extension packages
- **Rich Interoperability**: Support for `CalcRange` objects to handle cell ranges efficiently
- **Python UNO Bridge**: On Linux (and other platforms where the .NET UNO bridge is unavailable), Calc-DNA generates a Python script that uses [pythonnet](https://github.com/pythonnet/pythonnet) to load your .NET assembly and expose functions via LibreOffice's Python component loader
- **Cross-Platform**: Works on Windows, Linux, and macOS

## Architecture

Calc-DNA is designed around a pipeline that transforms simple C# methods into fully functional LibreOffice add-in components. Two deployment paths are supported:

- **UNO .NET bridge** (Windows / source-built LibreOffice): LibreOffice loads the .NET assembly directly via its built-in .NET UNO bridge.
- **Python UNO bridge** (Linux and other platforms): A generated Python script bootstraps pythonnet, loads your .NET assembly, and presents your functions to LibreOffice through the Python component loader. Activated with the `--python` CLI flag.

### 1. Compile-Time Code Generation (Roslyn Source Generators)

Two incremental source generators run during compilation:

- **CalcWrapperGenerator**: For each `[CalcFunction]` method, generates two sets of wrapper methods:
  - `_UNOWrapper` — accepts UNO-compatible types (`object[][]`, `double`, `string`), marshals inputs via `UnoMarshal`, invokes the user method, and marshals the return value back. Used by the .NET UNO bridge path.
  - `_PyWrapper` — accepts pythonnet-compatible types (`object` for sequences, since pythonnet wraps Python lists/tuples as `IEnumerable`), marshals via `PyMarshal`, and invokes the user method. Used by the Python UNO bridge path.

- **UnoServiceGenerator**: Generates UNO service classes that implement the required LibreOffice interfaces (`XAddIn`, `XServiceInfo`, `XLocalizable`). These generated classes:
  - Expose function metadata (names, descriptions, categories)
  - Route function calls to the appropriate wrapper methods
  - Handle UNO component lifecycle and registration

### 2. Runtime Type Marshalling (CalcDNA.Runtime)

The runtime library provides:

- **UnoMarshal**: Bidirectional type conversion for the .NET UNO bridge path:
  - Primitives: `double`, `int`, `string`, `bool`
  - Collections: `List<T>`, `T[]`
  - Cell ranges: `CalcRange` (wraps `object[][]` with typed accessors)
  - Null handling: replaces nulls with `DBNull.Value` for UNO compatibility

- **PyMarshal**: Bidirectional type conversion for the Python UNO bridge path:
  - Input: accepts any `IEnumerable` (pythonnet wraps Python lists/tuples as generic `IEnumerable` in .NET), delegates scalar conversions to `UnoMarshal.ConvertValue<T>`
  - Output: preserves .NET nulls (does not substitute `DBNull`), returning plain `object[]` / `object[][]` that pythonnet can convert back to Python lists
  - Provides `UnwrapOptional*` / `UnwrapNullable*` helpers for complex optional and nullable parameter types

- **CalcRange**: A strongly-typed wrapper around LibreOffice cell ranges that provides:
  - Enumeration over cell values
  - Row/column access
  - Type-safe value extraction

### 3. CLI Metadata Generation & Packaging (CalcDNA.CLI)

The CLI tool processes compiled assemblies to generate LibreOffice extension metadata and package everything into a `.oxt` file:

- **IDL Generator**: Creates `.idl` files defining the UNO component interfaces
- **XCU Generator**: Produces `.xcu` files that register functions, specify categories, and provide localized descriptions
- **RDB Generator**: Invokes the LibreOffice SDK's `idlc` and `regmerge` tools to compile IDL into `.rdb` type database files
- **PythonUnoServiceGenerator** *(--python mode)*: Generates a Python script that:
  - Bootstraps pythonnet and loads your .NET assembly at runtime via `clr.AddReference`
  - Defines a UNO service class per `[CalcAddIn]` class, delegating each function to the corresponding `_PyWrapper` method
  - Registers itself via `g_ImplementationHelper` so LibreOffice's `pythonloader` discovers it automatically
- **OxtPackager**: Assembles all generated files into a `.oxt` ZIP archive. In Python mode, the generated `.py` script and all DLLs are included; `manifest.xml` declares the `.py` file with media type `application/vnd.sun.star.uno-component;type=Python` and marks DLLs as `application/octet-stream` (loaded by pythonnet, not LO's .NET bridge).

### Data Flow

#### .NET Bridge (default)

```
User's C# Code
    ↓
[CalcAddIn] + [CalcFunction] Attributes
    ↓
Roslyn Source Generators (compile-time)
    ├─→ _UNOWrapper methods (UnoMarshal)
    └─→ UNO Service Classes (XAddIn implementation)
    ↓
Compiled Assembly (.dll)
    ↓
CLI Tool (post-build)
    ├─→ .idl (interface definitions)
    ├─→ .xcu (function registry)
    ├─→ .rdb (compiled type database)
    └─→ .oxt (packaged extension)
    ↓
LibreOffice loads .dll via .NET UNO bridge
    ↓
User Functions Available in Calc
```

#### Python Bridge (--python)

```
User's C# Code
    ↓
[CalcAddIn] + [CalcFunction] Attributes
    ↓
Roslyn Source Generators (compile-time)
    ├─→ _PyWrapper methods (PyMarshal)
    └─→ UNO Service Classes (XAddIn implementation)
    ↓
Compiled Assembly (.dll)
    ↓
CLI Tool --python (post-build)
    ├─→ .idl → .rdb (type database)
    ├─→ .xcu (function registry)
    ├─→ Generated .py (UNO service + pythonnet bootstrap)
    └─→ .oxt (manifest declares .py as Python UNO component)
    ↓
LibreOffice pythonloader imports .py
    → pythonnet loads .dll into the same process
    → Python UNO service delegates calls to _PyWrapper methods
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

The same code produces a working extension under both deployment paths — no changes required when switching between .NET bridge and Python bridge.

## Repository Structure

```
Calc-DNA/
├── src/
│   ├── CalcDNA.Attributes/          # Attribute definitions
│   │   ├── CalcAddInAttribute.cs    # Marks a class as containing UDFs
│   │   ├── CalcFunctionAttribute.cs # Marks a method as a UDF
│   │   └── CalcParameterAttribute.cs # Provides parameter metadata
│   │
│   ├── CalcDNA.Generator/           # Roslyn source generators
│   │   ├── CalcWrapperGenerator.cs  # Generates _UNOWrapper and _PyWrapper methods
│   │   ├── UnoServiceGenerator.cs   # Generates UNO service classes
│   │   └── WrapperTypeMapping.cs    # Maps .NET types to UNO and Python wrapper types
│   │
│   ├── CalcDNA.Runtime/             # Runtime support library
│   │   ├── CalcRange.cs             # Strongly-typed range wrapper
│   │   ├── UnoMarshal.cs            # Type conversion for .NET UNO bridge
│   │   ├── PyMarshal.cs             # Type conversion for Python UNO bridge
│   │   └── Uno/                     # UNO interface definitions
│   │       ├── IXAddIn.cs           # Add-in component interface
│   │       ├── IXServiceInfo.cs     # Service metadata interface
│   │       └── IXLocalizable.cs     # Localization interface
│   │
│   ├── CalcDNA.CLI/                 # Command-line metadata generator & packager
│   │   ├── Program.cs               # CLI entry point (supports --python flag)
│   │   ├── IdlGenerator.cs          # Generates .idl files
│   │   ├── XcuGenerator.cs          # Generates .xcu files
│   │   ├── RdbGenerator.cs          # Invokes SDK tools to create .rdb
│   │   ├── OxtPackager.cs           # Assembles .oxt extension packages
│   │   ├── PythonUnoServiceGenerator.cs # Generates Python UNO bridge script
│   │   ├── ManifestGenerator.cs     # Generates META-INF/manifest.xml
│   │   ├── DescriptionGenerator.cs  # Generates description.xml
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
  └─→ (no internal dependencies)

CalcDNA.Attributes
  └─→ (no dependencies)
```

## Prerequisites

- **.NET 10 SDK** (or .NET 8+) for building the framework and user add-ins
- **LibreOffice SDK** with `idlc` and `regmerge` tools (required for RDB generation)
- **pythonnet** *(Python bridge only)*: LibreOffice on the target platform must have pythonnet available in its Python environment. On Linux this typically means installing pythonnet into the Python that LibreOffice ships with (see [DEVELOPMENT.md](DEVELOPMENT.md) for details)

## Building the Project

Build the entire solution:

```bash
dotnet build
```

This will:
1. Compile all projects in dependency order
2. Run Roslyn source generators during `Demo.App` compilation (generating both `_UNOWrapper` and `_PyWrapper` methods, plus UNO service code)
3. Produce assemblies in each project's `bin/` directory

## Using the CLI Tool

The CLI tool generates the LibreOffice extension metadata and packages everything into a `.oxt` file:

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
- `--python`: Generate a Python UNO bridge script instead of relying on the .NET UNO bridge. Use this on Linux or any platform where the .NET bridge is unavailable.

### Examples

Generate a standard .NET bridge extension:
```bash
dotnet run --project src/CalcDNA.CLI -- ./Demo.App/bin/Debug/net10.0/Demo.App.dll --output ./output --verbose
```

Generate a Python bridge extension (for Linux):
```bash
dotnet run --project src/CalcDNA.CLI -- ./Demo.App/bin/Debug/net10.0/Demo.App.dll --output ./output --python --verbose
```

### Generated Files

| File | Description |
|------|-------------|
| `.idl` | Interface Definition Language file defining UNO component interfaces |
| `.xcu` | XML Configuration Unit registering functions and metadata |
| `.rdb` | Registry Database (compiled IDL) for LibreOffice's component loader |
| `.oxt` | Complete LibreOffice extension package (ZIP archive) |
| `.py` *(--python only)* | Python UNO service script that bootstraps pythonnet and delegates to `_PyWrapper` methods |

### How the Python Bridge Works

When `--python` is specified, the CLI generates a Python script (`<AddInName>.py`) included in the `.oxt` package. At runtime in LibreOffice:

1. LibreOffice's `pythonloader` discovers the `.py` file via its `application/vnd.sun.star.uno-component;type=Python` manifest entry
2. The script uses `clr.AddReference` (pythonnet) to load your .NET assembly into the same process
3. A Python UNO service class is instantiated for each `[CalcAddIn]` class
4. When Calc invokes a function, the Python service calls the corresponding `_PyWrapper` method on your .NET class
5. `PyMarshal` converts between Python sequences (tuples/lists, exposed as `IEnumerable` by pythonnet) and .NET types

Your .NET code runs in-process alongside LibreOffice — no inter-process marshalling, no separate .NET runtime. State in your .NET classes persists for the lifetime of the LibreOffice process.

## Testing

The project uses xUnit for testing with two test suites:

### CalcDNA.Generator.Tests

Tests for Roslyn source generators using the `Microsoft.CodeAnalysis.Testing` framework:

- **CalcWrapperGeneratorTests**: Verifies that source generators produce correct wrapper code for various scenarios (simple types, ranges, lists, error conditions). Tests validate both `_UNOWrapper` and `_PyWrapper` output.
- **WrapperTypeMappingTests**: Validates type mapping logic between .NET and UNO/Python wrapper types

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
- **Generator tests**: Verify source generation produces correct code for both bridge paths
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
- [x] Automated `.oxt` packaging (`OxtPackager`)
- [x] Python UNO bridge (`--python` mode: `PyMarshal`, `_PyWrapper` generation, `PythonUnoServiceGenerator`)

### In Progress

- [ ] Error handling and diagnostics improvements
- [ ] Enhanced logging and debugging support
- [ ] Documentation and inline code comments

### Planned

- [ ] **Enhanced type support**: Dictionaries, nullable types, custom objects
- [ ] **Attribute enhancements**: Category, volatile functions, help URLs
- [ ] **Integration test automation**: Automated testing against LibreOffice
- [ ] **Performance optimization**: Caching, lazy initialization, reduced allocations
- [ ] **Async support**: Asynchronous function execution (requires UNO threading investigation)
- [ ] **Developer tooling**: VS Code extension, debugging helpers, live reload

### Known Limitations

- **No .NET UNO bridge on Linux**: LibreOffice's prebuilt packages on Linux do not include the .NET UNO bridge (tracked as [LO Bug 165585](https://bugs.documentfoundation.org/show_bug.cgi?id=165585)). Use `--python` mode as the workaround.
- **pythonnet dependency**: Python bridge mode requires pythonnet to be available in LibreOffice's Python environment
- **No async/await support**: All functions must be synchronous
- **Limited error context**: UNO exceptions lose some context during marshalling
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
- **pythonnet** - For enabling Python/.NET interop that powers the cross-platform Python bridge
- The open-source community for making cross-platform development possible
