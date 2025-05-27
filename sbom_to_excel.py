#!/usr/bin/env python3
"""
SBOM to Excel Converter
Converts SPDX and CycloneDX (Software Bill of Materials) files to Excel format
Supports:
- SPDX 2.2/2.3 in tag-value and JSON formats
- CycloneDX 1.3/1.4/1.5 in JSON and XML formats
"""

import argparse
import json
import sys
from pathlib import Path
import pandas as pd
import re
import xml.etree.ElementTree as ET
from typing import Dict, List, Tuple


class SBOMParser:
    """Base class for SBOM parsers"""
    
    def __init__(self, filepath):
        self.filepath = Path(filepath)
        self.packages = []
        self.document_info = {}
        
    def parse(self) -> Tuple[List[Dict], Dict]:
        """Parse the file and return packages and document info"""
        raise NotImplementedError


class SPDXParser(SBOMParser):
    """Parse SPDX files in various formats"""
    
    def parse(self) -> Tuple[List[Dict], Dict]:
        """Determine file format and parse accordingly"""
        with open(self.filepath, 'r', encoding='utf-8') as f:
            content = f.read()
            
        # Check if JSON format
        if content.strip().startswith('{'):
            self._parse_json(content)
        else:
            self._parse_tag_value(content)
            
        return self.packages, self.document_info
    
    def _parse_json(self, content: str):
        """Parse JSON format SPDX file"""
        try:
            data = json.loads(content)
            
            # Extract document information
            self.document_info = {
                'Format': 'SPDX',
                'SPDX Version': data.get('spdxVersion', ''),
                'Data License': data.get('dataLicense', ''),
                'Document Name': data.get('name', ''),
                'Document Namespace': data.get('documentNamespace', ''),
                'Creator': ', '.join(data.get('creationInfo', {}).get('creators', [])),
                'Created': data.get('creationInfo', {}).get('created', '')
            }
            
            # Extract packages
            packages = data.get('packages', [])
            for pkg in packages:
                package_data = {
                    'Package Name': pkg.get('name', ''),
                    'Version': pkg.get('versionInfo', ''),
                    'Type': 'library',  # SPDX doesn't have explicit type
                    'PURL': '',  # SPDX uses different identification
                    'SPDXID': pkg.get('SPDXID', ''),
                    'Download Location': pkg.get('downloadLocation', ''),
                    'License': pkg.get('licenseConcluded', ''),
                    'License Declared': pkg.get('licenseDeclared', ''),
                    'Copyright': pkg.get('copyrightText', ''),
                    'Description': pkg.get('description', ''),
                    'Checksum': self._extract_checksum(pkg.get('checksums', [])),
                    'External References': self._extract_external_refs(pkg.get('externalRefs', [])),
                    'Supplier': pkg.get('supplier', ''),
                    'Author': pkg.get('originator', ''),
                    'Files Analyzed': pkg.get('filesAnalyzed', True),
                    'Verification Code': pkg.get('packageVerificationCode', {}).get('packageVerificationCodeValue', ''),
                    'Comment': pkg.get('comment', '')
                }
                
                # Try to extract PURL from external refs
                for ref in pkg.get('externalRefs', []):
                    if ref.get('referenceType') == 'purl':
                        package_data['PURL'] = ref.get('referenceLocator', '')
                        break
                
                self.packages.append(package_data)
                
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON format: {e}")
    
    def _parse_tag_value(self, content: str):
        """Parse tag-value format SPDX file"""
        lines = content.split('\n')
        current_package = None
        in_package = False
        
        # Initialize document info
        self.document_info = {'Format': 'SPDX'}
        
        for line in lines:
            line = line.strip()
            if not line or line.startswith('#'):
                continue
                
            # Parse tag-value pairs
            match = re.match(r'^(\w+):\s*(.*)$', line)
            if not match:
                continue
                
            tag, value = match.groups()
            
            # Document information tags
            if tag == 'SPDXVersion':
                self.document_info['SPDX Version'] = value
            elif tag == 'DataLicense':
                self.document_info['Data License'] = value
            elif tag == 'DocumentName':
                self.document_info['Document Name'] = value
            elif tag == 'DocumentNamespace':
                self.document_info['Document Namespace'] = value
            elif tag == 'Creator':
                if 'Creator' not in self.document_info:
                    self.document_info['Creator'] = value
                else:
                    self.document_info['Creator'] += f", {value}"
            elif tag == 'Created':
                self.document_info['Created'] = value
            
            # Package tags
            elif tag == 'PackageName':
                if current_package:
                    self.packages.append(current_package)
                current_package = {
                    'Package Name': value,
                    'Version': '',
                    'Type': 'library',
                    'PURL': '',
                    'SPDXID': '',
                    'Download Location': '',
                    'License': '',
                    'License Declared': '',
                    'Copyright': '',
                    'Description': '',
                    'Checksum': '',
                    'External References': '',
                    'Supplier': '',
                    'Author': '',
                    'Files Analyzed': True,
                    'Verification Code': '',
                    'Comment': ''
                }
                in_package = True
            elif in_package and current_package:
                if tag == 'SPDXID':
                    current_package['SPDXID'] = value
                elif tag == 'PackageVersion':
                    current_package['Version'] = value
                elif tag == 'PackageDownloadLocation':
                    current_package['Download Location'] = value
                elif tag == 'FilesAnalyzed':
                    current_package['Files Analyzed'] = value.lower() == 'true'
                elif tag == 'PackageVerificationCode':
                    current_package['Verification Code'] = value
                elif tag == 'PackageChecksum':
                    current_package['Checksum'] = value
                elif tag == 'PackageLicenseConcluded':
                    current_package['License'] = value
                elif tag == 'PackageLicenseDeclared':
                    current_package['License Declared'] = value
                elif tag == 'PackageCopyrightText':
                    current_package['Copyright'] = value
                elif tag == 'PackageDescription':
                    current_package['Description'] = value
                elif tag == 'PackageComment':
                    current_package['Comment'] = value
                elif tag == 'ExternalRef':
                    if current_package['External References']:
                        current_package['External References'] += f"; {value}"
                    else:
                        current_package['External References'] = value
                    # Check for PURL
                    if 'purl' in value:
                        parts = value.split()
                        if len(parts) >= 2 and parts[0] == 'PACKAGE-MANAGER':
                            current_package['PURL'] = parts[1]
                elif tag == 'PackageSupplier':
                    current_package['Supplier'] = value
                elif tag == 'PackageOriginator':
                    current_package['Author'] = value
        
        # Don't forget the last package
        if current_package:
            self.packages.append(current_package)
    
    def _extract_checksum(self, checksums: List[Dict]) -> str:
        """Extract checksum information from JSON format"""
        if not checksums:
            return ''
        checksum_strs = []
        for cs in checksums:
            algo = cs.get('algorithm', '')
            value = cs.get('checksumValue', '')
            checksum_strs.append(f"{algo}: {value}")
        return '; '.join(checksum_strs)
    
    def _extract_external_refs(self, refs: List[Dict]) -> str:
        """Extract external references from JSON format"""
        if not refs:
            return ''
        ref_strs = []
        for ref in refs:
            ref_type = ref.get('referenceType', '')
            locator = ref.get('referenceLocator', '')
            ref_strs.append(f"{ref_type}: {locator}")
        return '; '.join(ref_strs)


class CycloneDXParser(SBOMParser):
    """Parse CycloneDX files in JSON and XML formats"""
    
    def parse(self) -> Tuple[List[Dict], Dict]:
        """Determine file format and parse accordingly"""
        with open(self.filepath, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Check if JSON format
        if content.strip().startswith('{'):
            self._parse_json(content)
        else:
            # Assume XML format
            self._parse_xml(content)
            
        return self.packages, self.document_info
    
    def _parse_json(self, content: str):
        """Parse JSON format CycloneDX file"""
        try:
            data = json.loads(content)
            
            # Extract document information
            metadata = data.get('metadata', {})
            self.document_info = {
                'Format': 'CycloneDX',
                'BOM Format': data.get('bomFormat', 'CycloneDX'),
                'Spec Version': data.get('specVersion', ''),
                'Serial Number': data.get('serialNumber', ''),
                'Version': data.get('version', '1'),
                'Created': metadata.get('timestamp', ''),
                'Component Name': metadata.get('component', {}).get('name', ''),
                'Component Version': metadata.get('component', {}).get('version', ''),
                'Tool': ', '.join([tool.get('name', '') for tool in metadata.get('tools', [])])
            }
            
            # Extract components
            components = data.get('components', [])
            for comp in components:
                package_data = {
                    'Package Name': comp.get('name', ''),
                    'Version': comp.get('version', ''),
                    'Type': comp.get('type', ''),
                    'PURL': comp.get('purl', ''),
                    'SPDXID': '',  # CycloneDX doesn't use SPDXID
                    'Download Location': '',
                    'License': self._extract_licenses(comp.get('licenses', [])),
                    'License Declared': '',  # CycloneDX doesn't distinguish
                    'Copyright': comp.get('copyright', ''),
                    'Description': comp.get('description', ''),
                    'Checksum': self._extract_hashes(comp.get('hashes', [])),
                    'External References': self._extract_external_references(comp.get('externalReferences', [])),
                    'Supplier': comp.get('supplier', {}).get('name', ''),
                    'Author': comp.get('author', ''),
                    'Files Analyzed': False,  # CycloneDX doesn't have this concept
                    'Verification Code': '',  # CycloneDX doesn't use verification codes
                    'Comment': ''
                }
                
                self.packages.append(package_data)
                
        except json.JSONDecodeError as e:
            raise ValueError(f"Invalid JSON format: {e}")
    
    def _parse_xml(self, content: str):
        """Parse XML format CycloneDX file"""
        try:
            root = ET.fromstring(content)
            
            # Handle namespaces
            ns = {'cdx': 'http://cyclonedx.org/schema/bom/1.5'}
            if root.tag.startswith('{'):
                ns['cdx'] = root.tag.split('}')[0][1:]
            
            # Extract document information
            metadata = root.find('.//cdx:metadata', ns)
            self.document_info = {
                'Format': 'CycloneDX',
                'BOM Format': 'CycloneDX',
                'Spec Version': root.get('version', ''),
                'Serial Number': root.get('serialNumber', ''),
                'Version': root.get('version', '1'),
                'Created': '',
                'Component Name': '',
                'Component Version': '',
                'Tool': ''
            }
            
            if metadata is not None:
                timestamp = metadata.find('cdx:timestamp', ns)
                if timestamp is not None:
                    self.document_info['Created'] = timestamp.text or ''
                
                component = metadata.find('cdx:component', ns)
                if component is not None:
                    name = component.find('cdx:name', ns)
                    version = component.find('cdx:version', ns)
                    if name is not None:
                        self.document_info['Component Name'] = name.text or ''
                    if version is not None:
                        self.document_info['Component Version'] = version.text or ''
                
                tools = metadata.findall('.//cdx:tool', ns)
                tool_names = []
                for tool in tools:
                    tool_name = tool.find('cdx:name', ns)
                    if tool_name is not None and tool_name.text:
                        tool_names.append(tool_name.text)
                self.document_info['Tool'] = ', '.join(tool_names)
            
            # Extract components
            components = root.findall('.//cdx:component', ns)
            for comp in components:
                package_data = {
                    'Package Name': self._get_xml_text(comp, 'cdx:name', ns),
                    'Version': self._get_xml_text(comp, 'cdx:version', ns),
                    'Type': comp.get('type', ''),
                    'PURL': self._get_xml_text(comp, 'cdx:purl', ns),
                    'SPDXID': '',
                    'Download Location': '',
                    'License': self._extract_xml_licenses(comp, ns),
                    'License Declared': '',
                    'Copyright': self._get_xml_text(comp, 'cdx:copyright', ns),
                    'Description': self._get_xml_text(comp, 'cdx:description', ns),
                    'Checksum': self._extract_xml_hashes(comp, ns),
                    'External References': self._extract_xml_external_refs(comp, ns),
                    'Supplier': self._extract_xml_supplier(comp, ns),
                    'Author': self._get_xml_text(comp, 'cdx:author', ns),
                    'Files Analyzed': False,
                    'Verification Code': '',
                    'Comment': ''
                }
                
                self.packages.append(package_data)
                
        except ET.ParseError as e:
            raise ValueError(f"Invalid XML format: {e}")
    
    def _get_xml_text(self, element: ET.Element, path: str, ns: Dict[str, str]) -> str:
        """Safely get text from XML element"""
        found = element.find(path, ns)
        return found.text if found is not None and found.text else ''
    
    def _extract_licenses(self, licenses: List[Dict]) -> str:
        """Extract license information from CycloneDX JSON"""
        license_strs = []
        for lic in licenses:
            if 'license' in lic:
                if 'id' in lic['license']:
                    license_strs.append(lic['license']['id'])
                elif 'name' in lic['license']:
                    license_strs.append(lic['license']['name'])
        return '; '.join(license_strs)
    
    def _extract_xml_licenses(self, comp: ET.Element, ns: Dict[str, str]) -> str:
        """Extract license information from CycloneDX XML"""
        licenses = comp.findall('.//cdx:license', ns)
        license_strs = []
        for lic in licenses:
            lic_id = lic.find('cdx:id', ns)
            lic_name = lic.find('cdx:name', ns)
            if lic_id is not None and lic_id.text:
                license_strs.append(lic_id.text)
            elif lic_name is not None and lic_name.text:
                license_strs.append(lic_name.text)
        return '; '.join(license_strs)
    
    def _extract_hashes(self, hashes: List[Dict]) -> str:
        """Extract hash information from CycloneDX JSON"""
        hash_strs = []
        for h in hashes:
            algo = h.get('alg', '')
            content = h.get('content', '')
            hash_strs.append(f"{algo}: {content}")
        return '; '.join(hash_strs)
    
    def _extract_xml_hashes(self, comp: ET.Element, ns: Dict[str, str]) -> str:
        """Extract hash information from CycloneDX XML"""
        hashes = comp.findall('.//cdx:hash', ns)
        hash_strs = []
        for h in hashes:
            algo = h.get('alg', '')
            content = h.text or ''
            if algo and content:
                hash_strs.append(f"{algo}: {content}")
        return '; '.join(hash_strs)
    
    def _extract_external_references(self, refs: List[Dict]) -> str:
        """Extract external references from CycloneDX JSON"""
        ref_strs = []
        for ref in refs:
            ref_type = ref.get('type', '')
            url = ref.get('url', '')
            ref_strs.append(f"{ref_type}: {url}")
        return '; '.join(ref_strs)
    
    def _extract_xml_external_refs(self, comp: ET.Element, ns: Dict[str, str]) -> str:
        """Extract external references from CycloneDX XML"""
        refs = comp.findall('.//cdx:externalReference', ns)
        ref_strs = []
        for ref in refs:
            ref_type = ref.get('type', '')
            url_elem = ref.find('cdx:url', ns)
            url = url_elem.text if url_elem is not None else ''
            if ref_type and url:
                ref_strs.append(f"{ref_type}: {url}")
        return '; '.join(ref_strs)
    
    def _extract_xml_supplier(self, comp: ET.Element, ns: Dict[str, str]) -> str:
        """Extract supplier information from CycloneDX XML"""
        supplier = comp.find('cdx:supplier', ns)
        if supplier is not None:
            name = supplier.find('cdx:name', ns)
            if name is not None and name.text:
                return name.text
        return ''


def detect_format(filepath: Path) -> str:
    """Detect the SBOM format from file content"""
    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read(1000)  # Read first 1000 chars
    
    # Check for SPDX indicators
    if 'SPDXVersion' in content or '"spdxVersion"' in content:
        return 'SPDX'
    
    # Check for CycloneDX indicators
    if 'cyclonedx' in content.lower() or 'bomFormat' in content or '<bom' in content:
        return 'CycloneDX'
    
    # Default based on file extension
    if filepath.suffix.lower() == '.spdx':
        return 'SPDX'
    
    return 'Unknown'


def create_excel_report(packages: List[Dict], document_info: Dict, output_path: Path):
    """Create Excel file with SBOM data"""
    
    # Create a Pandas Excel writer
    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        
        # Create packages sheet
        if packages:
            df_packages = pd.DataFrame(packages)
            
            # Reorder columns for better readability
            column_order = [
                'Package Name', 'Version', 'Type', 'License', 'PURL', 
                'Author', 'Supplier', 'Copyright', 'Description', 'Checksum',
                'External References', 'Download Location', 'SPDXID', 
                'License Declared', 'Files Analyzed', 'Verification Code', 'Comment'
            ]
            
            # Only include columns that exist
            column_order = [col for col in column_order if col in df_packages.columns]
            df_packages = df_packages[column_order]
            
            # Remove empty columns
            df_packages = df_packages.loc[:, (df_packages != '').any(axis=0)]
            
            df_packages.to_excel(writer, sheet_name='Components', index=False)
            
            # Auto-adjust column widths
            worksheet = writer.sheets['Components']
            for idx, col in enumerate(df_packages.columns):
                max_length = max(
                    df_packages[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
                # Limit column width to 50 characters for readability
                worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
        
        # Create document info sheet
        df_info = pd.DataFrame(list(document_info.items()), columns=['Property', 'Value'])
        df_info.to_excel(writer, sheet_name='Document Info', index=False)
        
        # Auto-adjust column widths for document info
        worksheet = writer.sheets['Document Info']
        for idx, col in enumerate(df_info.columns):
            max_length = max(
                df_info[col].astype(str).map(len).max(),
                len(col)
            ) + 2
            worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
        
        # Create summary statistics sheet
        if packages:
            # Calculate statistics
            df_packages_full = pd.DataFrame(packages)
            
            summary_data = {
                'SBOM Format': document_info.get('Format', 'Unknown'),
                'Total Components': len(packages),
                'Unique Licenses': len(set(p['License'] for p in packages if p['License'])),
                'Components with Version': sum(1 for p in packages if p['Version']),
                'Components with PURL': sum(1 for p in packages if p['PURL']),
                'Components with Checksums': sum(1 for p in packages if p['Checksum']),
                'Components with External Refs': sum(1 for p in packages if p['External References']),
                'Component Types': ', '.join(set(p['Type'] for p in packages if p['Type']))
            }
            
            # Add license breakdown
            license_counts = {}
            for p in packages:
                if p['License']:
                    # Handle multiple licenses separated by semicolon
                    licenses = [l.strip() for l in p['License'].split(';')]
                    for lic in licenses:
                        if lic:
                            license_counts[lic] = license_counts.get(lic, 0) + 1
            
            df_summary = pd.DataFrame(list(summary_data.items()), columns=['Metric', 'Value'])
            df_summary.to_excel(writer, sheet_name='Summary', index=False)
            
            # Auto-adjust column widths for summary
            worksheet = writer.sheets['Summary']
            for idx, col in enumerate(df_summary.columns):
                max_length = max(
                    df_summary[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
                worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)
            
            # Add license breakdown if there are licenses
            if license_counts:
                df_licenses = pd.DataFrame(
                    list(license_counts.items()), 
                    columns=['License', 'Count']
                ).sort_values('Count', ascending=False)
                
                df_licenses.to_excel(writer, sheet_name='License Summary', index=False)
                
                # Auto-adjust column widths
                worksheet = writer.sheets['License Summary']
                for idx, col in enumerate(df_licenses.columns):
                    max_length = max(
                        df_licenses[col].astype(str).map(len).max(),
                        len(col)
                    ) + 2
                    worksheet.column_dimensions[chr(65 + idx)].width = min(max_length, 50)


def main():
    parser = argparse.ArgumentParser(
        description='Convert SBOM files (SPDX and CycloneDX) to Excel format',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog='''
Supported formats:
  - SPDX 2.2/2.3 (tag-value and JSON)
  - CycloneDX 1.3/1.4/1.5 (JSON and XML)

Examples:
  %(prog)s input.spdx
  %(prog)s cyclonedx.json -o output.xlsx
  %(prog)s sbom.xml -o sbom_report.xlsx
  %(prog)s any_sbom_file.json -v
        '''
    )
    
    parser.add_argument('input', help='Input SBOM file (SPDX or CycloneDX format)')
    parser.add_argument('-o', '--output', help='Output Excel file (default: input_file.xlsx)')
    parser.add_argument('-v', '--verbose', action='store_true', help='Enable verbose output')
    parser.add_argument('-f', '--format', choices=['auto', 'spdx', 'cyclonedx'], 
                       default='auto', help='SBOM format (default: auto-detect)')
    
    args = parser.parse_args()
    
    # Determine output filename
    input_path = Path(args.input)
    if not input_path.exists():
        print(f"Error: Input file '{args.input}' not found")
        sys.exit(1)
    
    if args.output:
        output_path = Path(args.output)
    else:
        output_path = input_path.with_suffix('.xlsx')
    
    try:
        # Detect format
        if args.format == 'auto':
            detected_format = detect_format(input_path)
            if detected_format == 'Unknown':
                print("Error: Could not auto-detect SBOM format. Please specify with -f option.")
                sys.exit(1)
        else:
            detected_format = args.format.upper()
        
        if args.verbose:
            print(f"Detected format: {detected_format}")
            print(f"Parsing {detected_format} file: {input_path}")
        
        # Parse SBOM file
        if detected_format == 'SPDX':
            parser = SPDXParser(input_path)
        elif detected_format in ['CYCLONEDX', 'CycloneDX']:
            parser = CycloneDXParser(input_path)
        else:
            print(f"Error: Unsupported format '{detected_format}'")
            sys.exit(1)
        
        packages, document_info = parser.parse()
        
        if args.verbose:
            print(f"Found {len(packages)} components")
            print(f"Document: {document_info.get('Document Name', document_info.get('Component Name', 'Unknown'))}")
        
        # Create Excel report
        create_excel_report(packages, document_info, output_path)
        
        print(f"Successfully created Excel report: {output_path}")
        
        if args.verbose:
            print("\nReport contents:")
            print(f"  - Document Info sheet: Basic SBOM document information")
            print(f"  - Components sheet: Detailed information for {len(packages)} components")
            print(f"  - Summary sheet: Statistical summary of the SBOM")
            print(f"  - License Summary sheet: License usage breakdown")
        
    except Exception as e:
        print(f"Error: {e}")
        if args.verbose:
            import traceback
            traceback.print_exc()
        sys.exit(1)


if __name__ == '__main__':
    main()