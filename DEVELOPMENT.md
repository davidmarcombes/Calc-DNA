# Development Environment Setup

This guide walks you through setting up your development environment for Calc-DNA on both Linux and Windows.

## Prerequisites

### Common Requirements
- **LibreOffice** 7.0 or later (24.2+ recommended)
- **.NET SDK** 8.0 or later
- **VS Code** with C# extensions
- **Git**

### Platform-Specific Requirements

#### Windows
- **Visual Studio Build Tools** (or full Visual Studio 2022)
- **Windows SDK** (usually included with VS Build Tools)

#### Linux (Ubuntu/Debian-based)
- **build-essential** package
- **mono-complete** (for some legacy UNO tooling)

## Installation Steps

### 1. Install LibreOffice

#### Windows
1. Download LibreOffice from https://www.libreoffice.org/download/download/
2. Run the installer
3. Install to default location (usually `C:\Program Files\LibreOffice`)

#### Linux (Ubuntu/Debian)
```bash
# Install LibreOffice
sudo apt update
sudo apt install libreoffice libreoffice-dev

# Verify installation
libreoffice --version
```

### 2. Install LibreOffice SDK

#### Windows
1. Download SDK from https://www.libreoffice.org/download/download/
2. Run the SDK installer
3. Note the installation path (e.g., `C:\Program Files\LibreOffice\sdk`)

#### Linux (Ubuntu/Debian)
```bash
# Install SDK
sudo apt install libreoffice-sdk

# SDK is typically installed to:
# /usr/lib/libreoffice/sdk
```

### 3. Install .NET SDK

#### Windows
1. Download from https://dotnet.microsoft.com/download
2. Run installer for .NET 8.0 SDK
3. Verify: `dotnet --version`

#### Linux
```bash
# Add Microsoft package repository
wget https://packages.microsoft.com/config/ubuntu/22.04/packages-microsoft-prod.deb -O packages-microsoft-prod.deb
sudo dpkg -i packages-microsoft-prod.deb
rm packages-microsoft-prod.deb

# Install .NET SDK
sudo apt update
sudo apt install dotnet-sdk-8.0

# Verify installation
dotnet --version
```

### 4. Install VS Code and Extensions

#### Both Platforms
1. Download VS Code from https://code.visualstudio.com/
2. Install the following extensions:
   - **C# Dev Kit** (Microsoft)
   - **C#** (Microsoft)
   - **.NET Extension Pack** (Microsoft)
   - **EditorConfig** (EditorConfig)

Open VS Code and press `Ctrl+Shift+X` (or `Cmd+Shift+X` on Mac), then search for and install each extension.

### 5. Configure Environment Variables

#### Windows
Add these to your system environment variables:

```batch
# LibreOffice program directory
LO_HOME=C:\Program Files\LibreOffice

# LibreOffice SDK directory
OO_SDK_HOME=C:\Program Files\LibreOffice\sdk

# URE (UNO Runtime Environment) directory
OO_SDK_URE_HOME=C:\Program Files\LibreOffice\URE

# Add to PATH
PATH=%PATH%;%LO_HOME%\program;%OO_SDK_HOME%\bin
```

To set these:
1. Right-click "This PC" → Properties → Advanced System Settings
2. Click "Environment Variables"
3. Add each variable under "System variables"

#### Linux
Add to your `~/.bashrc` or `~/.zshrc`:

```bash
# LibreOffice paths
export LO_HOME=/usr/lib/libreoffice
export OO_SDK_HOME=/usr/lib/libreoffice/sdk
export OO_SDK_URE_HOME=/usr/lib/libreoffice/ure

# Update PATH
export PATH=$PATH:$LO_HOME/program:$OO_SDK_HOME/bin

# Apply changes
source ~/.bashrc  # or source ~/.zshrc
```

### 6. Locate UNO CLI Assemblies

The C# UNO bindings are typically located at:

#### Windows
```
C:\Program Files\LibreOffice\program\
  - cli_basetypes.dll
  - cli_cppuhelper.dll
  - cli_ure.dll
  - cli_uretypes.dll
```

#### Linux
```
/usr/lib/libreoffice/program/
  - cli_basetypes.dll
  - cli_cppuhelper.dll
  - cli_ure.dll
  - cli_uretypes.dll
```

You'll need to reference these DLLs in your C# projects.

## Verify Installation

Run these commands to verify everything is set up correctly:

```bash
# Check LibreOffice
libreoffice --version

# Check .NET
dotnet --version

# Check environment variables (Linux/Mac)
echo $LO_HOME
echo $OO_SDK_HOME

# Check environment variables (Windows PowerShell)
echo $env:LO_HOME
echo $env:OO_SDK_HOME

# Verify CLI assemblies exist (Linux/Mac)
ls -la $LO_HOME/program/cli_*.dll

# Verify CLI assemblies exist (Windows PowerShell)
dir "$env:LO_HOME\program\cli_*.dll"
```

## Project Structure

The project follows a modular architecture:

```
Calc-DNA/
├── src/
│   ├── CalcDNA.Attributes/      # Attribute definitions (target for UDFs)
│   ├── CalcDNA.CLI/             # Command-line tool for IDL/XCU generation
│   ├── CalcDNA.Generator/       # Logic for generating IDL and service code
│   ├── CalcDNA.Runtime/         # Runtime support and UNO type marshalling
│   └── Demo.App/                # Sample Calc add-in
├── tests/
│   ├── CalcDNA.Generator.Tests/
│   └── CalcDNA.Runtime.Tests/
├── README.md
├── DEVELOPMENT.md
└── CalcDNA.slnx                 # Solution file (Visual Studio 2022+)
```

## Building the Project

To build the entire solution:

```bash
dotnet build
```

## Using the CLI Tool

The CLI tool is used to process your Add-In assembly and generate the necessary metadata for LibreOffice.

```bash
# Run the CLI tool
dotnet run --project src/CalcDNA.CLI -- <PathToYourAssembly.dll> [OutputPath] [AddInName]
```

### Options:
- `--verbose`: Enable detailed output.
- `assembly`: (Required) Path to the assembly containing your `[CalcAddIn]` classes.
- `output`: (Optional) Directory to save generated files.
- `name`: (Optional) Name for the add-in.
- `--sdk`: (Optional) Explicit path to LibreOffice SDK.

## Creating a New Add-In

1. **Create a new Class Library**:
   ```bash
   dotnet new classlib -n MyCalcAddIn
   ```

2. **Add Reference to CalcDNA.Attributes**:
   ```bash
   dotnet add MyCalcAddIn reference path/to/CalcDNA.Attributes.csproj
   ```

3. **Define your functions**:
   ```csharp
   using CalcDNA.Attributes;

   [CalcAddIn(Name = "MyTools")]
   public class MyFunctions {
       [CalcFunction]
       public double MySum(double a, double b) => a + b;
   }
   ```

4. **Generate Metadata**:
   Use the CLI tool as described above to generate the `.idl`, `.xcu`, and `.rdb` files.

## Testing

Run tests using:
```bash
dotnet test
```

## Resources

- [LibreOffice SDK Documentation](https://api.libreoffice.org/)
- [UNO/CLI Language Binding](https://wiki.openoffice.org/wiki/Uno/Cli)
- [.NET Documentation](https://docs.microsoft.com/en-us/dotnet/)
