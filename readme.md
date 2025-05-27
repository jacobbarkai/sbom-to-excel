# SBOM to Excel Converter

## Features

- **Multi-Format Support**: 
  - SPDX 2.2/2.3 (tag-value and JSON formats)
  - CycloneDX 1.3/1.4/1.5 (JSON and XML formats)
- **Auto-Detection**: Automatically detects# SBOM to Excel Converter

A universal Python tool for converting Software Bill of Materials (SBOM) files into well-organized Excel spreadsheets. This tool supports both SPDX and CycloneDX formats, helping organizations manage and analyze SBOM data in a familiar tabular format.

## Overview

The SBOM to Excel Converter transforms SPDX and CycloneDX files into structured Excel workbooks, making it easier to:
- Review software component licensing information
- Share SBOM data with non-technical stakeholders
- Perform compliance analysis and reporting
- Track dependencies and vulnerabilities
- Generate reports for audit purposes

## Features

- **Multi-Format Support**: 
  - SPDX 2.2/2.3 (tag-value and JSON formats)
  - CycloneDX 1.3/1.4/1.5 (JSON and XML formats)
- **Auto-Detection**: Automatically detects SBOM format from file content
- **Comprehensive Data Extraction**: Captures all essential fields including packages, licenses, checksums, and metadata
- **Unified Output**: Normalizes data from different SBOM formats into a consistent Excel structure
- **Organized Output**: Creates multi-sheet Excel workbooks with logical data organization
- **License Analysis**: Provides license usage breakdown and statistics
- **Auto-Formatting**: Automatically adjusts column widths for optimal readability
- **Summary Statistics**: Generates summary metrics for quick SBOM analysis
- **Command-Line Interface**: Simple CLI with intuitive options

## Installation

### Prerequisites

- Python 3.7 or higher
- pip package manager

### Install Dependencies

```bash
pip install -r requirements.txt
```

Or install packages directly:

```bash
pip install pandas openpyxl
```

## Usage

### Basic Usage

Convert an SBOM file to Excel format:

```bash
python sbom_to_excel.py input.spdx
python sbom_to_excel.py cyclonedx.json
python sbom_to_excel.py bom.xml
```

This creates an Excel file with the same name in the current directory.

### Specify Output File

```bash
python sbom_to_excel.py sbom.json -o my_report.xlsx
```

### Force Specific Format

```bash
python sbom_to_excel.py input.json -f spdx
python sbom_to_excel.py input.xml -f cyclonedx
```

### Verbose Mode

Get detailed processing information:

```bash
python sbom_to_excel.py input.spdx -v
```

### Command-Line Options

```
usage: sbom_to_excel.py [-h] [-o OUTPUT] [-v] [-f {auto,spdx,cyclonedx}] input

Convert SBOM files (SPDX and CycloneDX) to Excel format

positional arguments:
  input                 Input SBOM file (SPDX or CycloneDX format)

optional arguments:
  -h, --help            show this help message and exit
  -o OUTPUT, --output OUTPUT
                        Output Excel file (default: input_file.xlsx)
  -v, --verbose         Enable verbose output
  -f {auto,spdx,cyclonedx}, --format {auto,spdx,cyclonedx}
                        SBOM format (default: auto-detect)
```

## Output Structure

The generated Excel file contains four worksheets:

### 1. Components Sheet
Contains detailed information for each software component:
- Package Name
- Version
- Type (library, application, framework, etc.)
- License Information
- Package URL (PURL)
- Author and Supplier Information
- Copyright Text
- Checksums
- External References
- Additional metadata specific to each format

### 2. Document Info Sheet
Metadata about the SBOM document:
- Format (SPDX or CycloneDX)
- Format Version
- Document Name
- Creation Date
- Creator/Tool Information
- Document-specific metadata

### 3. Summary Sheet
Statistical overview of the SBOM:
- SBOM Format used
- Total number of components
- Unique licenses count
- Components with versions, PURLs, checksums
- Component type breakdown

### 4. License Summary Sheet
License usage analysis:
- Each unique license found
- Number of components using each license
- Sorted by usage frequency

## Supported SBOM Formats

### SPDX Formats

#### Tag-Value Format (.spdx)
```
SPDXVersion: SPDX-2.3
DataLicense: CC0-1.0
SPDXID: SPDXRef-DOCUMENT
DocumentName: Example Document
PackageName: Example Package
PackageVersion: 1.0.0
...
```

#### JSON Format (.json)
```json
{
  "spdxVersion": "SPDX-2.3",
  "dataLicense": "CC0-1.0",
  "name": "Example Document",
  "packages": [
    {
      "name": "Example Package",
      "versionInfo": "1.0.0",
      ...
    }
  ]
}
```

### CycloneDX Formats

#### JSON Format (.json)
```json
{
  "bomFormat": "CycloneDX",
  "specVersion": "1.5",
  "components": [
    {
      "type": "library",
      "name": "example-lib",
      "version": "1.0.0",
      "purl": "pkg:npm/example-lib@1.0.0"
    }
  ]
}
```

#### XML Format (.xml)
```xml
<?xml version="1.0" encoding="UTF-8"?>
<bom xmlns="http://cyclonedx.org/schema/bom/1.5">
  <components>
    <component type="library">
      <name>example-lib</name>
      <version>1.0.0</version>
      <purl>pkg:npm/example-lib@1.0.0</purl>
    </component>
  </components>
</bom>
```

## Use Cases

### 1. Compliance Review
Export SPDX data to Excel for legal and compliance teams to review licensing information without needing specialized SPDX tools.

### 2. Supply Chain Analysis
Create spreadsheets for tracking software dependencies across your organization's products.

### 3. Vulnerability Management
Export package lists to cross-reference with vulnerability databases and tracking systems.

### 4. Reporting and Documentation
Generate Excel reports for management reviews, audits, or customer deliverables.

### 5. Data Integration
Use Excel as an intermediate format to import SBOM data into other business systems.

## Example Workflow

1. **Generate SBOM** using your preferred tool (e.g., Syft, SPDX tools, CycloneDX tools)
2. **Convert to Excel**:
   ```bash
   python sbom_to_excel.py project_sbom.json -o project_sbom_report.xlsx
   ```
3. **Review and analyze** the data in Excel
4. **Share** with stakeholders or import into other systems

## Format-Specific Features

### SPDX Features
- Preserves SPDXID references
- Maintains both concluded and declared licenses
- Includes package verification codes
- Tracks file analysis status

### CycloneDX Features
- Preserves Package URLs (PURLs)
- Maintains component types (library, application, etc.)
- Includes external references with types
- Supports nested component structures

## Troubleshooting

### Common Issues

1. **File not found error**
   - Ensure the input file path is correct
   - Check file permissions

2. **Format detection failure**
   - Use `-f` flag to explicitly specify format: `-f spdx` or `-f cyclonedx`
   - Verify the file is a valid SBOM format

3. **Invalid format error**
   - Verify the SBOM file is valid (SPDX or CycloneDX)
   - Check for file corruption or incomplete downloads
   - Ensure XML files are well-formed

4. **Missing dependencies**
   - Run `pip install -r requirements.txt`
   - Ensure Python version is 3.7+

### Getting Help

For issues or questions:
1. Check the verbose output: `python sbom_to_excel.py input.json -v`
2. Verify SBOM file validity using format-specific validation tools
3. Try forcing the format type with `-f` option
4. Ensure all dependencies are correctly installed

## Limitations

- Focuses on component/package-level information (file-level details are not included)
- Large SBOM files may take time to process
- Excel row limit (1,048,576) applies to very large SBOMs
- CycloneDX nested dependencies are flattened in the output
- Some format-specific fields may be normalized or omitted for consistency

## Contributing

Contributions are welcome! Consider adding:
- Support for additional SBOM fields
- File-level information extraction
- Dependency tree visualization
- Custom filtering options
- Additional output formats (CSV, HTML)
- SPDX 3.0 support when available
- CycloneDX 1.6+ support
- Vulnerability data integration (VEX)

## License

The SBOM To Excel tool is Copyright (c) Jacob Barkai. All Rights Reserved.

Permission to modify and redistribute is granted under the terms of the Apache 2.0 license. See the [LICENSE] file for the full license.

[License]: https://github.com/jacobbarkai/sbom-to-excel/blob/main/LICENSE

## Related Tools

- [SPDX Tools](https://github.com/spdx/tools): Official SPDX tools
- [CycloneDX CLI](https://github.com/CycloneDX/cyclonedx-cli): Official CycloneDX tools
- [Syft](https://github.com/anchore/syft): SBOM generation tool
- [SPDX Specification](https://spdx.dev/): SPDX format documentation
- [CycloneDX Specification](https://cyclonedx.org/): CycloneDX format documentation
