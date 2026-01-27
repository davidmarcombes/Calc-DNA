# Calc-DNA

A framework for building LibreOffice Calc add-ins in C#, inspired by Excel-DNA.

## Overview

Calc-DNA enables C# developers to create user-defined functions (UDFs) for LibreOffice Calc using a simple, attribute-based API similar to Excel-DNA. The goal is to make it easy to port existing Excel add-in business logic to LibreOffice Calc, or to write once and deploy to both platforms.

## Motivation

Many Excel add-ins are built using Excel-DNA in C#, with clean separation between:
- **Excel-specific layer**: Registration, marshalling, Excel-DNA attributes
- **Business logic**: Pure C# calculation code

Calc-DNA provides an alternative registration layer for LibreOffice Calc, allowing developers to reuse their calculation implementations with minimal changes.

## Goals

### Primary Goals
1. **Simple API**: Attribute-based function decoration similar to Excel-DNA
2. **Minimal boilerplate**: Hide UNO complexity from developers
3. **Type safety**: Automatic marshalling between .NET and UNO types
4. **Cross-platform**: Work on Windows, Linux, and macOS
5. **Easy packaging**: Simple build process to create .oxt extensions

### Secondary Goals
- Source generation for optimal performance (future)
- Support for async functions (future)
- RTD-like functionality (future)
- Debugging tools (future)

## Project Status

🚧 **Early Development** - This project is in the initial planning and setup phase.

### Current Phase
- [x] Project planning and documentation
- [ ] Development environment setup
- [ ] Basic UNO component in C#
- [ ] Core attribute definitions
- [ ] Type marshalling layer
- [ ] Function registration system
- [ ] Build and packaging tooling

### Roadmap


## Contributing

This project is in early development. Contributions, ideas, and feedback are welcome!

## License

[To be determined]

## Acknowledgments

- **Excel-DNA**: Major inspiration for API design and architecture
- **LibreOffice**: For the UNO component framework
- The LibreOffice and .NET communities
