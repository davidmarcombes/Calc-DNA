# Extension Metadata

Define all your extension package metadata in code using the `[CalcExtensionMetadata]` attribute. No command-line arguments needed!

## Quick Start

Create a file in your add-in project (e.g., `ExtensionInfo.cs`) with assembly-level metadata:

```csharp
using CalcDNA.Attributes;

[assembly: CalcExtensionMetadata(
    Version = "1.0.0",
    DisplayName = "My Awesome Add-in",
    Publisher = "Your Name",
    Description = "Does amazing calculations in LibreOffice Calc"
)]
```

That's it! Now when you build and package:

```bash
dotnet build
dotnet run --project path/to/CalcDNA.CLI -- MyAddIn.dll --package
```

The CLI will automatically read the metadata from your assembly and use it in the `.oxt` package.

## All Available Properties

```csharp
[assembly: CalcExtensionMetadata(
    // Version string for upgrade detection (REQUIRED for upgrades)
    Version = "1.2.0",

    // Display name shown in Extension Manager
    DisplayName = "My Calc Add-in",

    // Publisher/author name
    Publisher = "Your Company",

    // Short description of functionality
    Description = "Advanced financial calculations for LibreOffice Calc",

    // Unique identifier (default: org.calcdna.{assemblyname})
    Identifier = "com.mycompany.mycalcaddin",

    // Minimum LibreOffice version required
    MinLibreOfficeVersion = "7.0",

    // Maximum LibreOffice version (optional, leave empty for no limit)
    MaxLibreOfficeVersion = "",

    // URL to update.xml for automatic update checking
    UpdateUrl = "https://example.com/updates/MyAddIn.update.xml",

    // URL to release notes for this version
    ReleaseNotesUrl = "https://example.com/releases/v1.2.0.html",

    // URL to extension icon (optional)
    IconUrl = "https://example.com/icon.png"
)]
```

## Complete Example

Here's a complete example from the Demo.App project:

```csharp
using CalcDNA.Attributes;

// Extension metadata - all the package information in one place!
[assembly: CalcExtensionMetadata(
    Version = "1.2.0",
    DisplayName = "Demo Calc Add-in",
    Publisher = "Calc-DNA Project",
    Description = "Demonstrates the power of Calc-DNA with various function examples.",
    Identifier = "org.calcdna.demo",
    MinLibreOfficeVersion = "7.0",
    UpdateUrl = "https://example.com/calcdna/updates/Demo.update.xml",
    ReleaseNotesUrl = "https://example.com/calcdna/releases/v1.2.0.html"
)]

namespace Demo.App
{
    // Your CalcAddIn classes and functions...
}
```

## Where to Put the Attribute

**Recommended approach**: Create a dedicated file like `ExtensionInfo.cs` at the root of your project:

```
MyAddIn/
├── ExtensionInfo.cs         ← Put metadata here
├── MyFunctions.cs
├── MoreFunctions.cs
└── MyAddIn.csproj
```

This keeps all extension metadata in one place, making it easy to find and update.

**Alternative**: You can put it in any `.cs` file since it's an assembly-level attribute, but a dedicated file is cleaner.

## Versioning Best Practices

### Use Semantic Versioning

```csharp
Version = "1.2.3"
//         ┬ ┬ ┬
//         │ │ └─ PATCH: Bug fixes
//         │ └─── MINOR: New features (backward compatible)
//         └───── MAJOR: Breaking changes
```

### Version Workflow

1. **Initial Release**
```csharp
[assembly: CalcExtensionMetadata(Version = "1.0.0")]
```

2. **Bug Fix**
```csharp
[assembly: CalcExtensionMetadata(Version = "1.0.1")]
```

3. **New Features**
```csharp
[assembly: CalcExtensionMetadata(Version = "1.1.0")]
```

4. **Breaking Changes**
```csharp
[assembly: CalcExtensionMetadata(Version = "2.0.0")]
```

When users install a newer version, LibreOffice automatically recognizes it as an upgrade based on the identifier and version number.

## Command-Line Overrides (Optional)

While metadata should be in code, you can still override values via command line if needed:

```bash
dotnet run --project CalcDNA.CLI -- MyAddIn.dll --package \
  --version "1.0.1-beta" \
  --publisher "Different Publisher"
```

Command-line options override assembly attributes. This is useful for:
- Testing different configurations
- CI/CD pipelines that inject version numbers
- Special builds (beta, rc, etc.)

## Integration with Build Process

Since metadata is in code, your build/package workflow becomes simple:

```bash
# Build your add-in
dotnet build

# Package it (metadata read from assembly)
dotnet run --project ../CalcDNA.CLI/CalcDNA.CLI.csproj -- ./bin/Debug/net10.0/MyAddIn.dll --package
```

No need to remember or script version numbers, publisher info, etc. It's all versioned with your code!

## Automatic Updates Setup

To enable automatic update checking:

1. **Set UpdateUrl in metadata:**
```csharp
[assembly: CalcExtensionMetadata(
    Version = "1.0.0",
    UpdateUrl = "https://yoursite.com/updates/MyAddIn.update.xml"
)]
```

2. **Host an update.xml file** at that URL:
```xml
<?xml version="1.0" encoding="UTF-8"?>
<description xmlns="http://openoffice.org/extensions/update/2006"
             xmlns:xlink="http://www.w3.org/1999/xlink">
  <identifier value="org.calcdna.myaddin" />
  <version value="1.1.0" />
  <update-download>
    <src xlink:href="https://yoursite.com/downloads/MyAddIn-1.1.0.oxt" />
  </update-download>
</description>
```

3. Users can now check for updates via Extension Manager!

See [VERSIONING_AND_UPDATES.md](VERSIONING_AND_UPDATES.md) for complete details on the update system.

## FAQ

### Do I need to specify all properties?

No! Only properties you set are used. Defaults are:
- Version: "1.0.0"
- Identifier: Generated from assembly name
- DisplayName: Assembly name
- MinLibreOfficeVersion: "7.0"
- All other properties: Empty

### Can I have multiple [CalcExtensionMetadata] attributes?

No, only one per assembly. The attribute is marked with `AllowMultiple = false`.

### What if I don't use this attribute?

The CLI will use defaults. Your extension will work, but you won't have proper versioning, publisher info, etc.

### Can I read these values in my code at runtime?

Yes! They're standard assembly attributes:

```csharp
var assembly = Assembly.GetExecutingAssembly();
var metadata = assembly.GetCustomAttribute<CalcExtensionMetadataAttribute>();
if (metadata != null)
{
    Console.WriteLine($"Version: {metadata.Version}");
    Console.WriteLine($"Publisher: {metadata.Publisher}");
}
```

### What's the identifier used for?

The identifier uniquely identifies your extension. When LibreOffice sees two extensions with the same identifier:
- It treats the newer version as an upgrade
- It prevents installing both simultaneously

Choose carefully - changing it later creates a "different" extension!

## Examples

### Minimal

```csharp
[assembly: CalcExtensionMetadata(
    Version = "1.0.0"
)]
```

### Typical

```csharp
[assembly: CalcExtensionMetadata(
    Version = "2.1.0",
    DisplayName = "Financial Tools",
    Publisher = "FinanceCorp",
    Description = "Advanced financial calculations and analysis tools"
)]
```

### Full-Featured

```csharp
[assembly: CalcExtensionMetadata(
    Version = "3.0.0",
    DisplayName = "Pro Calc Tools",
    Publisher = "AnalyticsSoft Inc.",
    Description = "Professional data analysis and visualization for Calc",
    Identifier = "com.analyticssoft.procalctools",
    MinLibreOfficeVersion = "24.2",
    UpdateUrl = "https://updates.analyticssoft.com/procalctools.update.xml",
    ReleaseNotesUrl = "https://analyticssoft.com/releases/procalctools/v3.0.0"
)]
```
