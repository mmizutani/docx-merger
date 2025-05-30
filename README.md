# DOCX Merger

A C# command-line application that merges multiple Word documents (DOCX files) using the OpenXML PowerTools library.

## Features

- Merges multiple DOCX files into a single document
- Preserves formatting and styles from the first document
- Handles sections, headers, and footers appropriately
- Command-line interface for easy automation
- Comprehensive error handling and validation
- Built-in test document creation for quick testing

## Quick Start

1. **Clone and build the project:**
   ```bash
   git clone <repository-url>
   cd docx-merger
   dotnet build
   ```

2. **Create test documents and try it out:**
   ```bash
   dotnet run -- --create-test
   dotnet run -- test1.docx test2.docx merged.docx
   ```

3. **Or run the demo script:**
   ```bash
   ./demo.sh
   ```

## Requirements

- .NET 9.0 or later
- Valid DOCX files (OpenXML format)

## Installation

1. Clone or download the project
2. Navigate to the project directory
3. Restore NuGet packages and build:
   ```bash
   dotnet restore
   dotnet build
   ```

## Usage

### Command Line Syntax
```bash
DocxMerger <input1.docx> <input2.docx> [input3.docx ...] <output.docx>
```

### Examples

Merge two documents:
```bash
dotnet run document1.docx document2.docx merged.docx
```

Merge three documents:
```bash
dotnet run report1.docx report2.docx appendix.docx final_report.docx
```

### Arguments

- `input*.docx` - Input Word documents to merge (minimum 2 required)
- `output.docx` - Output merged document file

## Technical Details

This application uses the [Open-XML-PowerTools](https://github.com/OpenXmlDev/Open-Xml-PowerTools) library, which provides:

- High-fidelity document merging
- Proper handling of styles, fonts, and formatting
- Support for headers, footers, and sections
- Automatic resolution of style conflicts

### How It Works

1. **Validation**: Checks that all input files exist and are accessible
2. **Source Creation**: Creates `Source` objects for each input document
3. **Merge Process**: Uses `DocumentBuilder.BuildDocument()` to combine documents
4. **Output**: Saves the merged document to the specified output file

### Source Configuration

Each document is configured as a `Source` with:
- `keepSections: true` - Preserves section formatting from the first document
- Full document content is included (no partial document merging)

## Error Handling

The application handles common scenarios:

- **Missing Files**: Clear error messages for non-existent input files
- **Invalid Arguments**: Usage help when incorrect arguments are provided
- **Processing Errors**: Detailed error information for debugging

## Dependencies

- **OpenXmlPowerTools**: Version 4.5.3.2 - Core document processing library
- **.NET 9.0**: Target framework

## Building for Distribution

To create a standalone executable:

```bash
# Windows
dotnet publish -c Release -r win-x64 --self-contained

# macOS
dotnet publish -c Release -r osx-x64 --self-contained

# Linux
dotnet publish -c Release -r linux-x64 --self-contained
```

## Notes

- Documents are merged in the order specified on the command line
- The first document's styles and page setup are used as the template
- All input files must be valid DOCX files (OpenXML format)
- Large documents may take some time to process

## License

This project uses the OpenXML PowerTools library, which is licensed under the MIT License.
