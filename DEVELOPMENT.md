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

The recommended project structure:

```
calc-dna/
├── src/
│   ├── CalcDna.Core/              # Core framework (attributes, types)
│   │   ├── CalcDna.Core.csproj
│   │   └── ...
│   ├── CalcDna.Host/              # UNO component implementation
│   │   ├── CalcDna.Host.csproj
│   │   └── ...
│   └── CalcDna.Build/             # Build and packaging tools
│       ├── CalcDna.Build.csproj
│       └── ...
├── samples/
│   └── SimpleFunctions/           # Example add-in
│       ├── SimpleFunctions.csproj
│       └── Functions.cs
├── tests/
│   └── CalcDna.Tests/
│       └── CalcDna.Tests.csproj
├── docs/
├── .editorconfig
├── .gitignore
├── Directory.Build.props           # Shared MSBuild properties
├── README.md
├── DEVELOPMENT.md
└── AGENTS.md
```

## Creating Your First Project

```bash
# Create solution
dotnet new sln -n CalcDna

# Create core library
dotnet new classlib -n CalcDna.Core -o src/CalcDna.Core -f net8.0
dotnet sln add src/CalcDna.Core/CalcDna.Core.csproj

# Create host component
dotnet new classlib -n CalcDna.Host -o src/CalcDna.Host -f net8.0
dotnet sln add src/CalcDna.Host/CalcDna.Host.csproj

# Create sample
dotnet new classlib -n SimpleFunctions -o samples/SimpleFunctions -f net8.0
dotnet sln add samples/SimpleFunctions/SimpleFunctions.csproj

# Add reference from Host to Core
dotnet add src/CalcDna.Host reference src/CalcDna.Core

# Add reference from Sample to Core
dotnet add samples/SimpleFunctions reference src/CalcDna.Core
```

## Adding UNO Assembly References

Edit your `.csproj` files to reference the UNO CLI assemblies:

```xml
<Project Sdk="Microsoft.NET.Sdk">
  <PropertyGroup>
    <TargetFramework>net8.0</TargetFramework>
  </PropertyGroup>

  <ItemGroup>
    <!-- Windows paths -->
    <Reference Include="cli_basetypes" Condition="'$(OS)' == 'Windows_NT'">
      <HintPath>C:\Program Files\LibreOffice\program\cli_basetypes.dll</HintPath>
      <Private>false</Private>
    </Reference>
    <Reference Include="cli_cppuhelper" Condition="'$(OS)' == 'Windows_NT'">
      <HintPath>C:\Program Files\LibreOffice\program\cli_cppuhelper.dll</HintPath>
      <Private>false</Private>
    </Reference>
    <Reference Include="cli_ure" Condition="'$(OS)' == 'Windows_NT'">
      <HintPath>C:\Program Files\LibreOffice\program\cli_ure.dll</HintPath>
      <Private>false</Private>
    </Reference>
    <Reference Include="cli_uretypes" Condition="'$(OS)' == 'Windows_NT'">
      <HintPath>C:\Program Files\LibreOffice\program\cli_uretypes.dll</HintPath>
      <Private>false</Private>
    </Reference>

    <!-- Linux paths -->
    <Reference Include="cli_basetypes" Condition="'$(OS)' != 'Windows_NT'">
      <HintPath>/usr/lib/libreoffice/program/cli_basetypes.dll</HintPath>
      <Private>false</Private>
    </Reference>
    <Reference Include="cli_cppuhelper" Condition="'$(OS)' != 'Windows_NT'">
      <HintPath>/usr/lib/libreoffice/program/cli_cppuhelper.dll</HintPath>
      <Private>false</Private>
    </Reference>
    <Reference Include="cli_ure" Condition="'$(OS)' != 'Windows_NT'">
      <HintPath>/usr/lib/libreoffice/program/cli_ure.dll</HintPath>
      <Private>false</Private>
    </Reference>
    <Reference Include="cli_uretypes" Condition="'$(OS)' != 'Windows_NT'">
      <HintPath>/usr/lib/libreoffice/program/cli_uretypes.dll</HintPath>
      <Private>false</Private>
    </Reference>
  </ItemGroup>
</Project>
```

## Testing Your Setup

Create a simple test file to verify UNO assemblies load:

```csharp
// File: test-uno.cs
using System;
using unoidl.com.sun.star.uno;

class Program
{
    static void Main()
    {
        Console.WriteLine("UNO assemblies loaded successfully!");
        Console.WriteLine($"XInterface type: {typeof(XInterface)}");
    }
}
```

Compile and run:
```bash
dotnet run
```

If this succeeds, your development environment is ready!

## Common Issues

### Issue: Cannot find UNO assemblies

**Solution**: Verify the paths in your `.csproj` match your LibreOffice installation. Use:
- Windows: `where /R "C:\Program Files\LibreOffice" cli_ure.dll`
- Linux: `find /usr/lib/libreoffice -name "cli_ure.dll"`

### Issue: LibreOffice won't load the extension

**Solution**: Check that:
1. Extension is built for correct .NET version
2. All dependencies are included in .oxt
3. manifest.xml is correctly formatted
4. Component registration is correct

### Issue: Functions don't appear in Calc

**Solution**: 
1. Check LibreOffice extension manager (Tools → Extension Manager)
2. Verify extension is enabled
3. Restart LibreOffice completely
4. Check for errors in: Help → About LibreOffice → Show log

## Next Steps

1. Read through the [AGENTS.md](AGENTS.md) file for coding guidelines
2. Review the LibreOffice SDK documentation
3. Start with the proof-of-concept (see README.md roadmap)
4. Join the LibreOffice development community for support

## Resources

- [LibreOffice SDK Documentation](https://api.libreoffice.org/)
- [UNO/CLI Language Binding](https://wiki.openoffice.org/wiki/Uno/Cli)
- [.NET Documentation](https://docs.microsoft.com/en-us/dotnet/)
- [VS Code C# Documentation](https://code.visualstudio.com/docs/languages/csharp)
