#!/usr/bin/env python3
"""
================================================================================
Benchmark Analysis Tool - Version 2.9
================================================================================
Script to process multiple Excel files and generate comparison with version tracking.

Version History:
- v1.0: Initial release with CSV file inputs
- v2.0: Major redesign - switched to Excel file input, optional folders, 3 ratio columns
- v2.1: Multiple Excel files support, API version detection, cross-version comparisons
- v2.2: Reorganized column order - backend columns first, then ratio columns
- v2.3: Simplified column names when only one input file (no version suffix)
- v2.4: Backup original backend pages into output workbook as separate sheets
- v2.5: Separate comparison pages per version + dedicated cross-version page
- v2.6: Extract summary columns from existing comparison page and append to new pages
- v2.7: Report issues AND filter benchmarks to only include those existing in all backends
- v2.8: Detect Skia version (m[0-9]+) and intelligently determine baseline/compare versions
- v2.9: Cache backend name extraction, simplify baseline handling, improve performance

Features:
- Accepts 0-2 folder paths for trace analysis (optional)
- Accepts MULTIPLE Excel files with multiple sheets (one per backend)
- Detects API version (api[0-9]+) AND Skia version (m[0-9]+) from filename
- Intelligently determines baseline and compare versions
- Generates comparison pages with baseline as reference
- Auto-increments output filename version number with baseline info
- Caches backend name extraction for better performance

Usage: python script.py [<folder_path1>] [<folder_path2>] <excel_file1.xlsx> [<excel_file2.xlsx> ...]

Author: Benchmark Analysis Tool
Version: 2.9
Date: 2026-06-15
================================================================================
"""

import sys
import os
import json
import re
import pandas as pd
from pathlib import Path
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment
from openpyxl.worksheet.table import Table, TableStyleInfo
from collections import defaultdict
import warnings
warnings.filterwarnings('ignore')

# Version constant
MAJOR_VERSION = 2
MINOR_VERSION = 9
VERSION = f"{MAJOR_VERSION}.{MINOR_VERSION}"

# Global variable for API version prefixes
API_PREFIXES = ['api', 'r']
SKIA_PREFIXES = ['m']

# Build regex pattern dynamically from API_PREFIXES
API_PATTERN = '|'.join(re.escape(prefix) for prefix in API_PREFIXES)
SKIA_PATTERN = '|'.join(re.escape(prefix) for prefix in SKIA_PREFIXES)

API_REGEX = re.compile(f'({API_PATTERN})(\\d+)', re.IGNORECASE)
SKIA_REGEX = re.compile(f'({SKIA_PATTERN})(\\d+)', re.IGNORECASE)

def extract_versions_from_filename(filename):
    """Extract API version (api[0-9]+ or r[0-9]+) and Skia version (m[0-9]+) from filename."""
    api_version = None
    skia_version = None
    
    # Extract API version - uses global API_REGEX
    api_match = API_REGEX.search(filename)
    if api_match:
        prefix = api_match.group(1).lower()
        number = api_match.group(2)
        api_version = f"{prefix}{number}"
    
    # Extract Skia version - uses global SKIA_REGEX
    skia_match = SKIA_REGEX.search(filename)
    if skia_match:
        prefix = skia_match.group(1).lower()
        number = skia_match.group(2)
        skia_version = f"{prefix}{number}"
    
    return api_version, skia_version

def determine_baseline_and_compare(version_groups):
    """
    Determine which version is baseline and which is compare.
    Rules:
    - If there's only one version, use it as baseline and extract its components
    - If input files have same Skia version, use Skia version as baseline, API version as compare
    - If input files have same API version, use API version as baseline, Skia version as compare
    - If neither, use the first version as baseline
    """
    versions = list(version_groups.keys())
    
    # Handle single version case
    if len(versions) <= 1:
        full_version = versions[0]
        # Extract API and Skia from the single version
        api_match = API_REGEX.search(full_version)
        skia_match = SKIA_REGEX.search(full_version)
        
        if api_match and skia_match:
            # Both API and Skia present - use Skia as baseline, API as compare
            api_prefix = api_match.group(1).lower()
            api_number = api_match.group(2)
            skia_prefix = skia_match.group(1).lower()
            skia_number = skia_match.group(2)
            baseline_version = f"{skia_prefix}{skia_number}"
            compare_versions = [f"{api_prefix}{api_number}"]
            comparison_type = "single_with_both"
            print(f"\n📌 Single version detected: {full_version}")
            print(f"   Using Skia version as baseline: {baseline_version}")
            print(f"   API version as compare: {compare_versions}")
        elif api_match:
            # Only API present
            prefix = api_match.group(1).lower()
            number = api_match.group(2)
            baseline_version = f"{prefix}{number}"
            compare_versions = [baseline_version]
            comparison_type = "single_api_only"
            print(f"\n📌 Single version detected: {full_version}")
            print(f"   Using API version as baseline: {baseline_version}")
        elif skia_match:
            # Only Skia present
            prefix = skia_match.group(1).lower()
            number = skia_match.group(2)
            baseline_version = f"{prefix}{number}"
            compare_versions = [baseline_version]
            comparison_type = "single_skia_only"
            print(f"\n📌 Single version detected: {full_version}")
            print(f"   Using Skia version as baseline: {baseline_version}")
        else:
            # No version detected
            baseline_version = full_version
            compare_versions = [baseline_version]
            comparison_type = "single_no_version"
            print(f"\n📌 Single version detected: {full_version}")
            print(f"   Using full version as baseline: {baseline_version}")
        
        return baseline_version, compare_versions, comparison_type
    
    # Extract components from version strings for multiple versions
    version_components = {}
    for version in versions:
        api_match = API_REGEX.search(version)
        skia_match = SKIA_REGEX.search(version)
        version_components[version] = {
            'api': f"{api_match.group(1).lower()}{api_match.group(2)}" if api_match else None,
            'skia': f"{skia_match.group(1).lower()}{skia_match.group(2)}" if skia_match else None,
            'full': version
        }
    
    # Check if all versions have same Skia version
    skia_versions = set()
    for v in version_components.values():
        if v['skia']:
            skia_versions.add(v['skia'])
    
    # Check if all versions have same API version
    api_versions = set()
    for v in version_components.values():
        if v['api']:
            api_versions.add(v['api'])
    
    # Case: All files have the SAME Skia version (regardless of API)
    if len(skia_versions) == 1 and len(versions) > 1:
        # Same Skia version across all files - use Skia as baseline
        baseline_skia = list(skia_versions)[0]
        baseline_version = baseline_skia
        
        # Compare versions are the API versions (extract from full version)
        compare_versions = []
        for version in versions:
            api_match = API_REGEX.search(version)
            if api_match:
                compare_versions.append(f"{api_match.group(1).lower()}{api_match.group(2)}")
            else:
                compare_versions.append(version)
        # Remove duplicates
        seen = set()
        compare_versions = [x for x in compare_versions if not (x in seen or seen.add(x))]
        
        # Determine if API versions are also the same
        if len(api_versions) == 1:
            comparison_type = "same_skia_same_api"
            print(f"\n📌 Detected: Same Skia version ({baseline_skia}) AND same API version across files")
        else:
            comparison_type = "same_skia_different_api"
            print(f"\n📌 Detected: Same Skia version ({baseline_skia}) across files")
        
        print(f"   Using Skia version as baseline, compare versions: {compare_versions}")
        return baseline_version, compare_versions, comparison_type
    
    # Case: All files have the SAME API version (different Skia)
    if len(api_versions) == 1 and len(versions) > 1:
        # Same API version across all files - use API as baseline
        baseline_api = list(api_versions)[0]
        baseline_version = baseline_api
        
        # Compare versions are the Skia versions (extract from full version)
        compare_versions = []
        for version in versions:
            skia_match = SKIA_REGEX.search(version)
            if skia_match:
                compare_versions.append(f"{skia_match.group(1).lower()}{skia_match.group(2)}")
            else:
                compare_versions.append(version)
        # Remove duplicates
        seen = set()
        compare_versions = [x for x in compare_versions if not (x in seen or seen.add(x))]
        
        comparison_type = "same_api_different_skia"
        print(f"\n📌 Detected: Same API version ({baseline_api}) across files")
        print(f"   Using API version as baseline, compare versions: {compare_versions}")
        return baseline_version, compare_versions, comparison_type
    
    # Default: use first version as baseline
    baseline_version = versions[0]
    compare_versions = versions[1:]
    comparison_type = "mixed_versions"
    print(f"\n📌 No clear version pattern detected")
    print(f"   Using '{baseline_version}' as baseline")
    return baseline_version, compare_versions, comparison_type

def get_next_version_number(base_filename):
    """Get the next available version number for output filename."""
    pattern = re.compile(rf'{re.escape(base_filename)}_v(\d+)\.xlsx$')
    existing_versions = []
    
    # Check for existing files with version numbers
    for file in Path('.').glob(f'{base_filename}_v*.xlsx'):
        match = pattern.match(file.name)
        if match:
            existing_versions.append(int(match.group(1)))
    
    if existing_versions:
        next_version = max(existing_versions) + 1
    else:
        next_version = 1
    
    return next_version

def generate_output_filename(baseline_version):
    """Generate output filename with baseline version."""
    # Create base name: [baseline_version]_crossplatform_comparison
    base_name = f"{baseline_version}_crossplatform_comparison"
    
    # Get next version number
    version_num = get_next_version_number(base_name)
    
    # Create filename
    output_file = f"{base_name}_v{version_num}.xlsx"
    
    return output_file, version_num

def validate_arguments():
    """Validate command line arguments."""
    if len(sys.argv) < 2:
        print("Error: At least one argument required (Excel file)")
        print(f"Usage: {sys.argv[0]} <excel_file1.xlsx> [<excel_file2.xlsx> ...]")
        print("Note: At least one Excel file parameter.")
        sys.exit(1)
    
    # Parse arguments: first 0-2 could be folders, rest are Excel files
    excel_files = []
    
    # Check each argument
    for arg in sys.argv[1:]:
        if arg.lower().endswith('.xlsx') or arg.lower().endswith('.xls'):
            excel_files.append(arg)
        else:
            print(f"Error: No other parameters allowed. Extra parameter: '{arg}'")
            sys.exit(1)
    
    if len(excel_files) == 0:
        print("Error: At least one Excel file (.xlsx or .xls) must be provided")
        print(f"Usage: {sys.argv[0]} [<folder_path1>] [<folder_path2>] <excel_file1.xlsx> [<excel_file2.xlsx> ...]")
        sys.exit(1)
    
    # Check Excel files exist
    for excel_file in excel_files:
        if not os.path.isfile(excel_file):
            print(f"Error: Excel file '{excel_file}' does not exist")
            sys.exit(1)
    
    return excel_files

def detect_backend_type(folder_path):
    """Detect backend type from folder name."""
    folder_name = Path(folder_path).name.lower()
    if 'gr' in folder_name:
        return 'graphite'
    else:
        return 'ganesh'

def extract_summary_columns_from_comparison(excel_file):
    """Extract summary columns from existing comparison sheet if it exists."""
    try:
        # Load the workbook
        wb = load_workbook(excel_file, data_only=True)
        
        # Check if comparison sheet exists
        if 'comparison' not in wb.sheetnames:
            print(f"  No 'comparison' sheet found in {Path(excel_file).name}")
            return None
        
        # Read the comparison sheet
        df = pd.read_excel(excel_file, sheet_name='comparison')
        
        # Find columns that are summary columns (contain 'summary' in name or are not standard columns)
        standard_columns = ['ID', 'Bench']
        summary_columns = {}
        
        for col in df.columns:
            if col not in standard_columns:
                # Check if it looks like a summary column
                if 'summary' in col.lower() or 'draw' in col.lower() or 'trace' in col.lower():
                    summary_columns[col] = df[col].tolist()
                    print(f"  - Found summary column: '{col}'")
        
        if summary_columns:
            print(f"  Extracted {len(summary_columns)} summary columns from existing comparison page")
            return summary_columns
        else:
            print(f"  No summary columns found in existing comparison page")
            return None
            
    except Exception as e:
        print(f"  Error reading comparison sheet from {Path(excel_file).name}: {e}")
        return None

def extract_summary_columns_for_version(excel_file, version_tag):
    """Extract summary columns from comparison page for a specific version."""
    try:
        # Load the workbook
        wb = load_workbook(excel_file, data_only=True)
        
        # Check if comparison sheet exists
        if 'comparison' not in wb.sheetnames:
            return None
        
        # Read the comparison sheet
        df = pd.read_excel(excel_file, sheet_name='comparison')
        
        # Find columns that are summary columns
        standard_columns = ['ID', 'Bench']
        summary_columns = {}
        
        for col in df.columns:
            if col not in standard_columns:
                # Check if it looks like a summary column
                if 'summary' in col.lower() or 'draw' in col.lower() or 'trace' in col.lower():
                    summary_columns[col] = df[col].tolist()
                    print(f"    - Found summary column for version {version_tag}: '{col}'")
        
        if summary_columns:
            print(f"    Extracted {len(summary_columns)} summary columns from comparison page for version {version_tag}")
            return summary_columns
        else:
            return None
            
    except Exception as e:
        print(f"    Error reading comparison sheet from {Path(excel_file).name}: {e}")
        return None

def check_columns_coverage(version_groups, all_backends):
    """Check and report columns (backends) that don't exist in all versions."""
    print("\n" + "="*60)
    print("📊 Column/Backend Coverage Analysis")
    print("="*60)
    
    # For each backend, track which versions have it
    backend_coverage = defaultdict(list)
    
    for version_name, version_data in version_groups.items():
        # Get all backend names (without version suffix) for this version
        version_backends = set()
        for col_name in version_data['dataframes'].keys():
            # Get backend name from mapping
            backend_name = version_data['backend_mapping'].get(col_name, col_name)
            version_backends.add(backend_name)
        
        for backend in version_backends:
            backend_coverage[backend].append(version_name)
    
    # Find backends missing from some versions
    all_versions = set(version_groups.keys())
    missing_backends = {}
    
    for backend, present_versions in backend_coverage.items():
        present_set = set(present_versions)
        missing = all_versions - present_set
        if missing:
            missing_backends[backend] = {
                'present_in': sorted(present_versions),
                'missing_in': sorted(missing)
            }
    
    if missing_backends:
        print("\n⚠️  Backends missing from some versions:")
        for backend, info in sorted(missing_backends.items()):
            print(f"\n  Backend '{backend}':")
            print(f"    - Present in: {', '.join(info['present_in'])}")
            print(f"    - Missing in: {', '.join(info['missing_in'])}")
    else:
        print("\n✅ All backends are present in all versions!")
    
    # Also find backends that are only in a single version
    single_version_backends = {b: v for b, v in backend_coverage.items() if len(v) == 1}
    if single_version_backends:
        print(f"\n📌 Backends unique to a single version ({len(single_version_backends)}):")
        for backend, versions in sorted(single_version_backends.items()):
            print(f"    - '{backend}' (only in {versions[0]})")
    
    print("="*60)
    
    return missing_backends, single_version_backends

def check_duplicate_benchmarks(version_groups):
    """Check and report duplicate benchmark names within each sheet."""
    print("\n" + "="*60)
    print("🔍 Duplicate Benchmark Detection")
    print("="*60)
    
    duplicates_found = False
    duplicate_report = {}
    
    for version_name, version_data in version_groups.items():
        print(f"\n  Checking version: {version_name}")
        version_duplicates = {}
        
        for sheet_name, df in version_data['sheets'].items():
            # Check for duplicates in the bench column
            bench_counts = df['bench'].value_counts()
            duplicates = bench_counts[bench_counts > 1]
            
            if len(duplicates) > 0:
                duplicates_found = True
                version_duplicates[sheet_name] = duplicates.to_dict()
                print(f"    ⚠️  Sheet '{sheet_name}' has {len(duplicates)} duplicate benchmark(s):")
                for bench_name, count in duplicates.items():
                    print(f"        - '{bench_name}' appears {count} times")
            else:
                print(f"    ✓ Sheet '{sheet_name}' has no duplicate benchmarks")
        
        if version_duplicates:
            duplicate_report[version_name] = version_duplicates
    
    if not duplicates_found:
        print("\n✅ No duplicate benchmarks found in any sheet!")
    
    print("="*60)
    
    return duplicate_report

def check_missing_benchmarks(version_groups):
    """Check and report benchmark cases that don't exist in all pages."""
    print("\n" + "="*60)
    print("📋 Benchmark Coverage Analysis")
    print("="*60)
    
    # Get all unique benchmarks across all versions
    all_benchmarks = set()
    for version_data in version_groups.values():
        for df in version_data['dataframes'].values():
            all_benchmarks.update(df['bench'].tolist())
    
    # For each version, track which benchmarks are present
    version_benchmarks = {}
    for version_name, version_data in version_groups.items():
        benchmarks = set()
        for df in version_data['dataframes'].values():
            benchmarks.update(df['bench'].tolist())
        version_benchmarks[version_name] = benchmarks
    
    # Find benchmarks missing from each version
    missing_report = {}
    for version_name, benchmarks in version_benchmarks.items():
        missing = all_benchmarks - benchmarks
        if missing:
            missing_report[version_name] = missing
    
    if missing_report:
        print("\n⚠️  Missing Benchmarks Detected:")
        for version_name, missing in sorted(missing_report.items()):
            print(f"\n  Version '{version_name}' is missing {len(missing)} benchmark(s):")
            for bench in sorted(missing):
                # Find which versions have this benchmark
                present_in = []
                for v, b in version_benchmarks.items():
                    if bench in b:
                        present_in.append(v)
                if present_in:
                    print(f"    - {bench} (present in: {', '.join(present_in)})")
                else:
                    print(f"    - {bench} (present in: none)")
    else:
        print("\n✅ All benchmarks are present in all versions!")
    
    # Also check for benchmarks that are unique to a single version
    benchmark_versions = defaultdict(list)
    for version_name, benchmarks in version_benchmarks.items():
        for bench in benchmarks:
            benchmark_versions[bench].append(version_name)
    
    unique_benchmarks = {bench: versions for bench, versions in benchmark_versions.items() if len(versions) == 1}
    if unique_benchmarks:
        print(f"\n📌 Benchmarks unique to a single version ({len(unique_benchmarks)}):")
        for bench, versions in sorted(unique_benchmarks.items()):
            print(f"    - {bench} (only in {versions[0]})")
    
    print("="*60)
    
    return missing_report, unique_benchmarks, version_benchmarks

def read_excel_sheets(excel_file, version_tag):
    """Read all sheets from Excel file, excluding 'comparison' sheet."""
    try:
        # Dictionary to store dataframes with backend names
        dataframes = {}
        # Dictionary to store original sheet data for backup
        original_sheets = {}
        # Dictionary to store backend name mapping
        backend_mapping = {}
        
        # Load the workbook
        wb = load_workbook(excel_file, data_only=True)
        
        # Iterate through all sheets
        for sheet_name in wb.sheetnames:
            # Skip the 'comparison' sheet if it exists
            if sheet_name.lower() == 'comparison':
                print(f"  - Skipping existing 'comparison' sheet in {Path(excel_file).name}")
                continue
            
            # Read sheet into DataFrame
            df = pd.read_excel(excel_file, sheet_name=sheet_name)
            
            # Validate required columns
            if 'bench' not in df.columns or 'mean' not in df.columns:
                print(f"  Warning: Sheet '{sheet_name}' in {Path(excel_file).name} missing 'bench' or 'mean' column - skipping")
                continue
            
            # Store original sheet data for backup
            original_sheets[sheet_name] = df
            
            # Extract backend name by removing version suffixes (any part with digits)
            parts = sheet_name.split('_')
            backend_parts = []
            for part in parts:
                if not any(char.isdigit() for char in part):
                    backend_parts.append(part)
            backend_name = '_'.join(backend_parts)
            
            # Create column name with version tag
            column_name = f"{sheet_name}_{version_tag}" if version_tag else sheet_name
            dataframes[column_name] = df
            backend_mapping[column_name] = backend_name
            print(f"  - Loaded sheet '{sheet_name}' -> column '{column_name}' (backend: {backend_name}): {len(df)} rows")
        
        return dataframes, original_sheets, backend_mapping
        
    except Exception as e:
        print(f"Error reading Excel file '{excel_file}': {e}")
        return None, None, None

def import_existing_comparison_page(excel_file, version_tag, version_data):
    """Import existing comparison page from Excel file if it exists."""
    try:
        # Load the workbook with data_only=False to access formulas
        wb = load_workbook(excel_file, data_only=False)
        
        # Check if comparison sheet exists
        if 'comparison' not in wb.sheetnames:
            print(f"    No existing 'comparison' sheet found in {Path(excel_file).name}")
            return None
        
        # Get the comparison sheet
        sheet = wb['comparison']
        
        # Read the sheet with formulas
        # First, get all data as values
        data = []
        headers = []
        
        # Get headers from first row
        for col_idx, cell in enumerate(sheet[1], 1):
            if cell.value is not None:
                headers.append(cell.value)
            else:
                headers.append(f"Unnamed: {col_idx}")
        
        # Read data rows
        for row in sheet.iter_rows(min_row=2, values_only=False):
            row_data = []
            for cell in row:
                if cell.value is not None:
                    # Check if it's a formula
                    if cell.data_type == 'f':  # 'f' means formula
                        # Store as FORMULA: prefix with the formula string
                        row_data.append(f"FORMULA:{cell.value}")
                    else:
                        row_data.append(cell.value)
                else:
                    row_data.append(None)
            data.append(row_data)
        
        # Create DataFrame
        df = pd.DataFrame(data, columns=headers)
        
        # Validate required columns
        if 'Bench' not in df.columns:
            print(f"    Warning: Existing comparison sheet missing 'Bench' column")
            return None
        
        # Clean up columns: Remove empty columns, unnamed columns, and columns that are entirely NaN
        columns_to_keep = []
        for col in df.columns:
            # Skip columns that are completely empty (all NaN or None)
            if df[col].isna().all() or df[col].isnull().all():
                print(f"    Skipping empty column: '{col}'")
                continue
            
            # Skip unnamed columns (typically from Excel where columns outside table area)
            if col.startswith('Unnamed:') or col == '' or col is None:
                print(f"    Skipping unnamed column: '{col}'")
                continue
            
            columns_to_keep.append(col)
        
        # Filter dataframe to only keep valid columns
        if columns_to_keep:
            df = df[columns_to_keep].copy()
            print(f"    Kept {len(columns_to_keep)} columns: {', '.join(columns_to_keep)}")
        else:
            print(f"    Warning: No valid columns found after cleaning")
            return None
        
        # Detect formula columns
        formula_columns = []
        if len(df) > 0:
            for col in df.columns:
                first_val = df[col].iloc[0]
                if isinstance(first_val, str) and first_val.startswith('FORMULA:'):
                    formula_columns.append(col)
        
        if formula_columns:
            print(f"    Detected formula columns: {', '.join(formula_columns)}")
        
        # Check if we need to filter benchmarks to those common across all backends
        dataframes = version_data['dataframes']
        backend_benchmarks = []
        for df_backend in dataframes.values():
            backend_benchmarks.append(set(df_backend['bench'].tolist()))
        
        if backend_benchmarks:
            common_benchmarks = set.intersection(*backend_benchmarks)
            # Filter the comparison dataframe to only include common benchmarks
            filtered_df = df[df['Bench'].isin(common_benchmarks)].copy()
            
            # Sort by Bench name for consistency
            filtered_df = filtered_df.sort_values('Bench').reset_index(drop=True)
            
            # Ensure ID column exists and is sequential
            if 'ID' in filtered_df.columns:
                # If ID exists, keep it but renumber
                filtered_df['ID'] = range(1, len(filtered_df) + 1)
            else:
                # Insert ID as first column
                filtered_df.insert(0, 'ID', range(1, len(filtered_df) + 1))
            
            print(f"    ✓ Imported existing comparison page: {len(filtered_df)} benchmarks (filtered to common backends)")
            print(f"    Columns preserved (including all config columns): {', '.join(filtered_df.columns.tolist())}")
            
            # Detect and report config columns (columns that look like ratio/diff)
            config_columns = [col for col in filtered_df.columns if 'ratio' in col.lower() or 'diff' in col.lower() or 'vs' in col.lower()]
            if config_columns:
                print(f"    Detected config columns: {', '.join(config_columns)}")
            
            return filtered_df
        else:
            # If no filtering needed, just return the cleaned dataframe
            # Ensure ID column exists and is sequential
            if 'ID' in df.columns:
                df['ID'] = range(1, len(df) + 1)
            else:
                df.insert(0, 'ID', range(1, len(df) + 1))
            
            print(f"    ✓ Imported existing comparison page: {len(df)} benchmarks")
            print(f"    Columns preserved (including all config columns): {', '.join(df.columns.tolist())}")
            
            # Detect and report config columns
            config_columns = [col for col in df.columns if 'ratio' in col.lower() or 'diff' in col.lower() or 'vs' in col.lower()]
            if config_columns:
                print(f"    Detected config columns: {', '.join(config_columns)}")
            
            return df
            
    except Exception as e:
        print(f"    Error importing comparison sheet from {Path(excel_file).name}: {e}")
        return None

def check_existing_comparison_page(excel_file):
    """Check if a comparison page exists in the Excel file."""
    try:
        wb = load_workbook(excel_file, data_only=True)
        exists = 'comparison' in wb.sheetnames
        wb.close()
        return exists
    except Exception as e:
        return False
        
def read_multiple_excel_files(excel_files):
    """Read all Excel files and combine their data."""
    all_dataframes = {}
    all_backends = set()
    version_info = {}
    all_original_sheets = {}
    version_groups = defaultdict(dict)  # Group data by version tag
    summary_columns_by_version = {}  # Store summary columns per version
    has_comparison_page = {}  # Track which versions have existing comparison pages
    
    print(f"\n📖 Reading Excel files...")
    
    # Process each Excel file
    for excel_file in excel_files:
        filename = Path(excel_file).stem
        api_version, skia_version = extract_versions_from_filename(filename)
        
        # Create version tag (prioritize showing both if available)
        if api_version and skia_version:
            version_tag = f"{api_version}_{skia_version}"
            print(f"\n📖 Reading Excel file: {Path(excel_file).name} (API: {api_version}, Skia: {skia_version})")
        elif api_version:
            version_tag = api_version
            print(f"\n📖 Reading Excel file: {Path(excel_file).name} (API: {api_version})")
        elif skia_version:
            version_tag = skia_version
            print(f"\n📖 Reading Excel file: {Path(excel_file).name} (Skia: {skia_version})")
        else:
            version_tag = "default"
            print(f"\n📖 Reading Excel file: {Path(excel_file).name} (no version detected)")
        
        # Check if comparison page exists
        has_comparison = check_existing_comparison_page(excel_file)
        has_comparison_page[version_tag] = has_comparison
        if has_comparison:
            print(f"  ✓ Found existing 'comparison' page in this file")
        else:
            print(f"  ℹ️ No existing 'comparison' page found")
        
        # Extract summary columns from this file's comparison page
        print(f"  Checking for existing summary columns...")
        summary_columns = extract_summary_columns_for_version(excel_file, version_tag)
        if summary_columns:
            summary_columns_by_version[version_tag] = summary_columns
        
        # Read the sheets
        dataframes, original_sheets, backend_mapping = read_excel_sheets(excel_file, version_tag)
        
        if dataframes and original_sheets and backend_mapping:
            all_dataframes.update(dataframes)
            
            # Store original sheets with version info
            for sheet_name, df in original_sheets.items():
                backup_sheet_name = f"{sheet_name}_{version_tag}"
                all_original_sheets[backup_sheet_name] = df
            
            # Group data by version
            version_groups[version_tag]['dataframes'] = dataframes
            version_groups[version_tag]['file'] = excel_file
            version_groups[version_tag]['sheets'] = original_sheets
            version_groups[version_tag]['backend_mapping'] = backend_mapping
            version_groups[version_tag]['api'] = api_version
            version_groups[version_tag]['skia'] = skia_version
            version_groups[version_tag]['has_comparison_page'] = has_comparison
            
            # Store summary columns for this version
            if summary_columns:
                version_groups[version_tag]['summary_columns'] = summary_columns
            
            version_info[version_tag] = {
                'file': excel_file,
                'columns': list(dataframes.keys()),
                'api': api_version,
                'skia': skia_version,
                'backend_mapping': backend_mapping,
                'has_summary': summary_columns is not None,
                'has_comparison_page': has_comparison
            }
            
            # Collect unique backend names from the mapping
            for backend_name in backend_mapping.values():
                all_backends.add(backend_name)
    
    if len(all_dataframes) == 0:
        print("Error: No valid data loaded from any Excel file")
        sys.exit(1)
    
    # Determine baseline and compare versions
    baseline_version, compare_versions, comparison_type = determine_baseline_and_compare(version_groups)
    
    # Check for missing/extra backends across versions
    missing_backends, single_version_backends = check_columns_coverage(version_groups, all_backends)
    
    # Check for duplicate benchmarks
    duplicate_report = check_duplicate_benchmarks(version_groups)
    
    # Check for missing benchmarks
    missing_report, unique_benchmarks, version_benchmarks = check_missing_benchmarks(version_groups)
    
    # Determine which summary columns to use (prefer baseline version's summary if available)
    primary_summary_columns = None
    if baseline_version in summary_columns_by_version:
        primary_summary_columns = summary_columns_by_version[baseline_version]
        print(f"\n📋 Using summary columns from baseline version '{baseline_version}'")
    elif summary_columns_by_version:
        # Use the first available summary columns
        first_version = list(summary_columns_by_version.keys())[0]
        primary_summary_columns = summary_columns_by_version[first_version]
        print(f"\n📋 Using summary columns from version '{first_version}' (baseline has no summary)")
    
    print(f"\n📊 Summary: Loaded {len(all_dataframes)} backend columns from {len(excel_files)} files")
    print(f"   Backends found: {', '.join(sorted(all_backends))}")
    print(f"   Versions found: {', '.join(version_groups.keys())}")
    print(f"   Baseline version: {baseline_version}")
    print(f"   Compare versions: {', '.join(compare_versions) if compare_versions else 'None'}")
    print(f"   Original sheets to backup: {len(all_original_sheets)}")
    
    return (all_dataframes, all_backends, version_info, version_groups, 
            all_original_sheets, primary_summary_columns, missing_report, 
            unique_benchmarks, duplicate_report, missing_backends, 
            single_version_backends, version_benchmarks, baseline_version, 
            compare_versions, comparison_type, summary_columns_by_version, has_comparison_page)

def create_version_comparison_page(version_data, version_name,
                                    summary_columns, missing_benchmarks_for_version, 
                                    version_benchmarks, baseline_version, is_baseline=False):
    """Create comparison page for a specific version."""
    dataframes = version_data['dataframes']
    backend_mapping = version_data['backend_mapping']
    
    # Use version-specific summary columns if available, otherwise use the provided ones
    version_summary_columns = version_data.get('summary_columns', summary_columns)
    
    # Check if we have an existing comparison page to import
    has_existing = version_data.get('has_comparison_page', False)
    excel_file = version_data.get('file')
    
    if has_existing and excel_file:
        print(f"    Found existing comparison page in {Path(excel_file).name} - importing it")
        imported_df = import_existing_comparison_page(excel_file, version_name, version_data)
        if imported_df is not None and not imported_df.empty:
            print(f"    ✓ Successfully imported existing comparison page with {len(imported_df)} benchmarks")
            print(f"    ✓ Preserved all existing columns including configs")
            return imported_df
        else:
            print(f"    ⚠️ Failed to import existing comparison page, generating new one")
    
    # If no existing page or import failed, generate using current logic
    print(f"    Generating new comparison page for version: {version_name}")
    
    # Find benchmarks that exist in ALL backends for this version
    backend_benchmarks = []
    for df in dataframes.values():
        backend_benchmarks.append(set(df['bench'].tolist()))
    
    if backend_benchmarks:
        # Find intersection of all benchmarks across all backends
        common_benchmarks = set.intersection(*backend_benchmarks)
        benches = sorted(list(common_benchmarks))
        print(f"    Filtered benchmarks: {len(common_benchmarks)} common out of {len(set.union(*backend_benchmarks))} total")
    else:
        benches = []
    
    if not benches:
        print(f"    WARNING: No common benchmarks found across all backends for version {version_name}")
        return pd.DataFrame()
    
    # Create ordered ID column
    ordered_ids = list(range(1, len(benches) + 1))
    
    # Prepare comparison data
    comparison_data = {
        'ID': ordered_ids,
        'Bench': benches
    }
    
    # Add mean columns for each backend (only for common benchmarks)
    backend_columns = {}  # Store column names for formula reference
    for col_name, df in sorted(dataframes.items()):
        mean_dict = dict(zip(df['bench'], df['mean']))
        
        # Get backend name from mapping (already extracted)
        display_name = backend_mapping.get(col_name, col_name)
        
        comparison_data[display_name] = [mean_dict.get(bench, float('nan')) for bench in benches]
        backend_columns[display_name] = display_name
    
    print(f"    Backend columns found: {', '.join(backend_columns.keys())}")
    
    # Generate ratio_configs dynamically:
    # 1. Compare any backend (except glesdmsaa) with glesdmsaa
    # 2. Compare any backend starting with 'gr' with grdawn_vk
    # 3. Compare grvk with vkdmsaa if both exist
    ratio_configs = []
    
    # Get all backends except glesdmsaa
    other_backends = [b for b in backend_columns.keys() if b != 'glesdmsaa']
    
    # 1. Compare each backend with glesdmsaa
    for backend in other_backends:
        ratio_configs.append((f"{backend} vs glesdmsaa", backend, 'glesdmsaa'))
    
    # 2. Compare any backend starting with 'gr' with grdawn_vk (if grdawn_vk exists)
    if 'grdawn_vk' in backend_columns:
        gr_backends = [b for b in backend_columns.keys() if b.startswith('gr') and b != 'grdawn_vk']
        for backend in gr_backends:
            ratio_configs.append((f"{backend} vs grdawn_vk", backend, 'grdawn_vk'))
    
    # 3. Also add grvk vs vkdmsaa if both exist (skip if already added by above rules)
    if 'grvk' in backend_columns and 'vkdmsaa' in backend_columns:
        # Check if this config already exists
        already_added = False
        for config_name, _, _ in ratio_configs:
            if config_name == 'grvk vs vkdmsaa':
                already_added = True
                break
        if not already_added:
            ratio_configs.append(('grvk vs vkdmsaa', 'grvk', 'vkdmsaa'))
    
    print(f"    Generated {len(ratio_configs)} ratio configurations:")
    for config in ratio_configs:
        print(f"      - {config[0]} = {config[1]}/{config[2]}")
    
    ratio_columns = []
    for config_name, num_backend, den_backend in ratio_configs:
        if num_backend in backend_columns and den_backend in backend_columns:
            # Ratio column (division)
            ratio_col_name = f"{config_name} (ratio)"
            comparison_data[ratio_col_name] = [f"FORMULA:{num_backend}/{den_backend}"] * len(benches)
            ratio_columns.append(ratio_col_name)
            print(f"      ✓ Added ratio column: '{ratio_col_name}' = {num_backend}/{den_backend}")
            
            # Diff column (subtraction)
            diff_col_name = f"{config_name} (diff)"
            comparison_data[diff_col_name] = [f"FORMULA:{num_backend}-{den_backend}"] * len(benches)
            ratio_columns.append(diff_col_name)
            print(f"      ✓ Added diff column: '{diff_col_name}' = {num_backend}-{den_backend}")
        else:
            print(f"      ✗ WARNING: Backends not found for {config_name}")
            if num_backend not in backend_columns:
                print(f"        Missing numerator: {num_backend}")
            if den_backend not in backend_columns:
                print(f"        Missing denominator: {den_backend}")
    
    print(f"    Total columns added: {len(ratio_columns)} (ratio + diff)")
    
    # Add summary columns from this version's comparison page if provided
    if version_summary_columns:
        for col_name, col_values in version_summary_columns.items():
            if len(col_values) == len(benches):
                comparison_data[col_name] = col_values
                print(f"      ✓ Added summary column: '{col_name}'")
            else:
                print(f"      ✗ WARNING: Summary column '{col_name}' has {len(col_values)} values, expected {len(benches)}")
    
    # Create DataFrame
    df = pd.DataFrame(comparison_data)
    
    # Add average row at the end
    if len(df) > 0:
        # Create a new row with 'AVERAGE' in the Bench column
        avg_row = {'ID': '', 'Bench': 'AVERAGE'}
        
        # Calculate average for each column (excluding non-numeric columns)
        for col in df.columns:
            if col not in ['ID', 'Bench']:
                # Try to convert to numeric, if fails, leave as empty string
                try:
                    # Check if column contains formula placeholders
                    if df[col].dtype == 'object' and len(df) > 0:
                        first_val = df[col].iloc[0]
                        if isinstance(first_val, str) and first_val.startswith('FORMULA:'):
                            # For formula columns, skip average (formulas can't be averaged as strings)
                            avg_row[col] = ''
                            continue
                    
                    # Convert column to numeric, coercing errors to NaN
                    numeric_values = pd.to_numeric(df[col], errors='coerce')
                    if not numeric_values.isna().all():
                        # Calculate mean, ignoring NaN
                        avg_value = numeric_values.mean()
                        avg_row[col] = avg_value
                    else:
                        avg_row[col] = ''
                except:
                    avg_row[col] = ''
        
        # Append the average row to the DataFrame
        df = pd.concat([df, pd.DataFrame([avg_row])], ignore_index=True)
    
    return df

def create_cross_version_page(version_groups, all_backends, summary_columns, missing_report, 
                               version_benchmarks, baseline_version, compare_versions, comparison_type):
    """Create cross-version comparison page with baseline as reference."""
    if len(version_groups) <= 1:
        return None
    
    # Find benchmarks that exist in ALL backends of ALL versions
    all_version_benchmarks = []
    
    for version_name, version_data in version_groups.items():
        # Get benchmarks for this version that exist in all backends of this version
        backend_benchmarks = []
        for df in version_data['dataframes'].values():
            backend_benchmarks.append(set(df['bench'].tolist()))
        
        if backend_benchmarks:
            version_common = set.intersection(*backend_benchmarks)
            all_version_benchmarks.append(version_common)
            print(f"    Version {version_name}: {len(version_common)} common benchmarks out of {len(set.union(*backend_benchmarks))} total")
    
    if all_version_benchmarks:
        # Find intersection across all versions
        common_benchmarks = set.intersection(*all_version_benchmarks)
        benches = sorted(list(common_benchmarks))
        print(f"    Cross-version common benchmarks: {len(common_benchmarks)} across all versions")
    else:
        benches = []
    
    if not benches:
        print(f"    WARNING: No common benchmarks found across all versions")
        return pd.DataFrame()
    
    # Create mapping between full version tag and compare version
    version_mapping = {}
    for full_tag in version_groups.keys():
        if comparison_type in ["same_skia_different_api", "same_skia_same_api"]:
            # Extract API part as compare version
            api_match = API_REGEX.search(full_tag)
            if api_match:
                compare_ver = f"{api_match.group(1).lower()}{api_match.group(2)}"
                version_mapping[full_tag] = compare_ver
            else:
                version_mapping[full_tag] = full_tag
        elif comparison_type == "same_api_different_skia":
            # Extract Skia part as compare version
            skia_match = SKIA_REGEX.search(full_tag)
            if skia_match:
                compare_ver = f"{skia_match.group(1).lower()}{skia_match.group(2)}"
                version_mapping[full_tag] = compare_ver
            else:
                version_mapping[full_tag] = full_tag
        else:
            # Mixed case - use full tag as is
            version_mapping[full_tag] = full_tag
    
    print(f"    Version mapping: {version_mapping}")
    
    # Prepare comparison data
    comparison_data = {
        'ID': list(range(1, len(benches) + 1)),
        'Bench': benches
    }
    
    # Track column names for formula references
    column_names = {}
    
    # Add all mean columns using compare version as column name
    for full_tag, version_data in version_groups.items():
        compare_ver = version_mapping[full_tag]
        backend_mapping = version_data.get('backend_mapping', {})
        
        for col_name, df in sorted(version_data['dataframes'].items()):
            backend = backend_mapping.get(col_name, col_name)
            mean_dict = dict(zip(df['bench'], df['mean']))
            # Use compare version as part of column name
            display_name = f"{backend}_{compare_ver}"
            comparison_data[display_name] = [mean_dict.get(bench, float('nan')) for bench in benches]
            column_names[display_name] = display_name
    
    # Add ratio AND diff columns comparing ALL possible pairs of compare versions
    # Only compare between compare_versions, NOT including baseline
    backends = ['grdawn_vk', 'glesdmsaa', 'vkdmsaa', 'grvk']
    
    print(f"    Generating ratio and diff columns for all compare version pairs...")
    
    # Generate pairs only from compare_versions (exclude baseline)
    for i, version1 in enumerate(compare_versions):
        for version2 in compare_versions[i+1:]:
            for backend in backends:
                col1_name = f"{backend}_{version1}"
                col2_name = f"{backend}_{version2}"
                
                if col1_name in column_names and col2_name in column_names:
                    # Generate ratio: version1 vs version2
                    ratio_col_name = f"{backend}_{version1}_vs_{version2}(ratio)"
                    comparison_data[ratio_col_name] = [f"FORMULA:{col1_name}/{col2_name}"] * len(benches)
                    print(f"      Added ratio: {ratio_col_name} = {col1_name}/{col2_name}")
                    
                    # Generate diff: version1 vs version2
                    diff_col_name = f"{backend}_{version1}_vs_{version2}(diff)"
                    comparison_data[diff_col_name] = [f"FORMULA:{col1_name}-{col2_name}"] * len(benches)
                    print(f"      Added diff: {diff_col_name} = {col1_name}-{col2_name}")
                else:
                    if col1_name not in column_names:
                        print(f"      Warning: Missing column {col1_name}")
                    if col2_name not in column_names:
                        print(f"      Warning: Missing column {col2_name}")
    
    # Add summary columns from existing comparison page if provided
    if summary_columns:
        for col_name, col_values in summary_columns.items():
            if len(col_values) == len(benches):
                comparison_data[col_name] = col_values
    
    # Create DataFrame
    df = pd.DataFrame(comparison_data)
    
    # Add average row at the end (same as version comparison)
    if len(df) > 0:
        # Create a new row with 'AVERAGE' in the Bench column
        avg_row = {'ID': '', 'Bench': 'AVERAGE'}
        
        # Calculate average for each column (excluding non-numeric columns)
        for col in df.columns:
            if col not in ['ID', 'Bench']:
                # Try to convert to numeric, if fails, leave as empty string
                try:
                    # Check if column contains formula placeholders
                    if df[col].dtype == 'object' and len(df) > 0:
                        first_val = df[col].iloc[0]
                        if isinstance(first_val, str) and first_val.startswith('FORMULA:'):
                            # For formula columns, skip average (formulas can't be averaged as strings)
                            avg_row[col] = ''
                            continue
                    
                    # Convert column to numeric, coercing errors to NaN
                    numeric_values = pd.to_numeric(df[col], errors='coerce')
                    if not numeric_values.isna().all():
                        # Calculate mean, ignoring NaN
                        avg_value = numeric_values.mean()
                        avg_row[col] = avg_value
                    else:
                        avg_row[col] = ''
                except:
                    avg_row[col] = ''
        
        # Append the average row to the DataFrame
        df = pd.concat([df, pd.DataFrame([avg_row])], ignore_index=True)
    
    return df

def write_dataframe_with_formulas(writer, sheet_name, df, add_average_row=True):
    """Write dataframe to Excel with proper Excel formulas, average row outside table."""
    if df.empty:
        print(f"    WARNING: DataFrame for '{sheet_name}' is empty, skipping")
        return
    
    # Check if we have an average row
    has_average = False
    avg_data = None
    
    if add_average_row and len(df) > 0 and 'Bench' in df.columns:
        if df.iloc[-1]['Bench'] == 'AVERAGE':
            has_average = True
            # Extract the average row data
            avg_data = df.iloc[-1].to_dict()
            # Remove the average row from the dataframe
            df = df.iloc[:-1]
    
    # First write the dataframe values without formulas
    df_for_write = df.copy()
    
    # Store formula information before clearing
    formula_info = {}
    for col in df_for_write.columns:
        if len(df_for_write) > 0:
            # Check if this column contains formulas
            col_formulas = []
            for idx, val in enumerate(df_for_write[col]):
                if isinstance(val, str) and val.startswith('FORMULA:'):
                    col_formulas.append((idx, val[8:]))  # Remove 'FORMULA:' prefix
            if col_formulas:
                formula_info[col] = col_formulas
                # Replace formula placeholders with None for initial write
                for idx, _ in col_formulas:
                    df_for_write[col].iloc[idx] = None
    
    # Write the dataframe
    df_for_write.to_excel(writer, sheet_name=sheet_name, index=False)
    
    # Now add the formulas directly to the Excel sheet
    workbook = writer.book
    sheet = workbook[sheet_name]
    
    # Build column letter mapping
    col_letters = {}
    for idx, col_name in enumerate(df.columns, 1):
        col_letters[col_name] = get_column_letter(idx)
    
    # The last data row (excluding average row)
    last_data_row = len(df) + 1  # +1 for header
    
    # Add formulas from formula_info
    for col_name, formulas in formula_info.items():
        if col_name in col_letters:
            col_letter = col_letters[col_name]
            col_idx = list(df.columns).index(col_name) + 1
            
            # Check if this column's formulas reference other columns in our sheet
            for row_idx_in_df, formula_expr in formulas:
                excel_row = row_idx_in_df + 2  # +2 because: 0-based index, +1 for header, +1 for 1-based Excel
                
                # Check if formula references columns that exist in our sheet
                # Parse the formula to find column references
                import re
                col_refs = re.findall(r'([A-Z]+)\d+', formula_expr)
                all_cols_exist = True
                for col_ref in col_refs:
                    # Check if this column reference corresponds to a column in our dataframe
                    # This is a simplified check - we'd need to map column letters to actual column names
                    pass
                
                # If the formula is a simple calculation, we can try to adjust it
                if '/' in formula_expr and not formula_expr.startswith('='):
                    # Simple division formula
                    parts = formula_expr.split('/')
                    if len(parts) == 2:
                        num_col_name = parts[0].strip()
                        den_col_name = parts[1].strip()
                        
                        if num_col_name in col_letters and den_col_name in col_letters:
                            num_col_letter = col_letters[num_col_name]
                            den_col_letter = col_letters[den_col_name]
                            formula = f"={num_col_letter}{excel_row}/{den_col_letter}{excel_row}"
                            cell = sheet.cell(row=excel_row, column=col_idx)
                            cell.value = formula
                            cell.number_format = "0.000"
                elif '-' in formula_expr and not formula_expr.startswith('='):
                    # Simple subtraction formula
                    parts = formula_expr.split('-')
                    if len(parts) == 2:
                        num_col_name = parts[0].strip()
                        den_col_name = parts[1].strip()
                        
                        if num_col_name in col_letters and den_col_name in col_letters:
                            num_col_letter = col_letters[num_col_name]
                            den_col_letter = col_letters[den_col_name]
                            formula = f"={num_col_letter}{excel_row}-{den_col_letter}{excel_row}"
                            cell = sheet.cell(row=excel_row, column=col_idx)
                            cell.value = formula
                            cell.number_format = "0.000"
                elif formula_expr.startswith('='):
                    # This is an Excel formula that was imported
                    # Try to preserve it as-is, but we may need to adjust row references
                    # For now, use it directly
                    cell = sheet.cell(row=excel_row, column=col_idx)
                    cell.value = formula_expr
                    if 'ratio' in col_name.lower() or 'diff' in col_name.lower():
                        cell.number_format = "0.000"
                else:
                    # More complex formula - try to preserve it
                    cell = sheet.cell(row=excel_row, column=col_idx)
                    cell.value = f"={formula_expr}"
                    if 'ratio' in col_name.lower() or 'diff' in col_name.lower():
                        cell.number_format = "0.000"
    
    # Also check for formula columns that might not have been caught in formula_info
    for col_idx, col_name in enumerate(df.columns, 1):
        if len(df) > 0:
            first_val = df[col_name].iloc[0]
            if isinstance(first_val, str) and first_val.startswith('FORMULA:') and col_name not in formula_info:
                # This column has formulas but wasn't caught in the initial detection
                formula_expr = first_val[8:]
                col_letter = get_column_letter(col_idx)
                
                for row_idx in range(2, last_data_row + 1):
                    if '/' in formula_expr:
                        parts = formula_expr.split('/')
                        if len(parts) == 2:
                            num_col_name = parts[0]
                            den_col_name = parts[1]
                            if num_col_name in col_letters and den_col_name in col_letters:
                                num_col_letter = col_letters[num_col_name]
                                den_col_letter = col_letters[den_col_name]
                                formula = f"={num_col_letter}{row_idx}/{den_col_letter}{row_idx}"
                                cell = sheet.cell(row=row_idx, column=col_idx)
                                cell.value = formula
                                cell.number_format = "0.000"
                    elif '-' in formula_expr:
                        parts = formula_expr.split('-')
                        if len(parts) == 2:
                            num_col_name = parts[0]
                            den_col_name = parts[1]
                            if num_col_name in col_letters and den_col_name in col_letters:
                                num_col_letter = col_letters[num_col_name]
                                den_col_letter = col_letters[den_col_name]
                                formula = f"={num_col_letter}{row_idx}-{den_col_letter}{row_idx}"
                                cell = sheet.cell(row=row_idx, column=col_idx)
                                cell.value = formula
                                cell.number_format = "0.000"
    
    # Add average row below the table (if it exists)
    if has_average and avg_data:
        avg_row_excel = last_data_row + 2  # Leave one empty row after table
        
        # Add AVERAGE label
        for idx, col_name in enumerate(df.columns, 1):
            if col_name == 'Bench':
                cell = sheet.cell(row=avg_row_excel, column=idx)
                cell.value = "AVERAGE"
                cell.font = Font(bold=True)
                cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif col_name not in ['ID']:
                # Check if this column should have an average
                should_have_average = False
                
                # Check if it's a formula column
                if col_name in formula_info or (len(df) > 0 and isinstance(df[col_name].iloc[0], str) and df[col_name].iloc[0].startswith('FORMULA:')):
                    should_have_average = True
                else:
                    # Check if it's a numeric column by looking at the actual data
                    for check_row in range(2, min(last_data_row + 1, 10)):  # Check first few rows
                        cell_value = sheet.cell(row=check_row, column=idx).value
                        if cell_value is not None and isinstance(cell_value, (int, float)):
                            should_have_average = True
                            break
                        elif cell_value is not None and isinstance(cell_value, str):
                            try:
                                float(cell_value)
                                should_have_average = True
                                break
                            except:
                                pass
                
                if should_have_average:
                    # Add AVERAGE formula for numeric/formula columns
                    col_letter = get_column_letter(idx)
                    avg_formula = f"=AVERAGE({col_letter}2:{col_letter}{last_data_row + 1})"
                    cell = sheet.cell(row=avg_row_excel, column=idx)
                    cell.value = avg_formula
                    cell.number_format = "0.000"
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                else:
                    # Non-numeric column - leave empty
                    cell = sheet.cell(row=avg_row_excel, column=idx)
                    cell.value = ""
                    cell.font = Font(bold=True)
                    cell.fill = PatternFill(start_color="FFF2CC", end_color="FFF2CC", fill_type="solid")
                    cell.alignment = Alignment(horizontal='center', vertical='center')
        
        print(f"    ✓ Added average row at row {avg_row_excel} (outside table)")
    
    return df  # Return the modified dataframe (without average row)

def apply_table_formatting_to_sheet(sheet, df):
    """Apply Excel table formatting to a sheet."""
    
    # Determine the range of the table (exclude average row)
    start_row = 1
    start_col = 1
    end_row = len(df) + 1  # +1 for header
    end_col = len(df.columns)
    
    # Create table range reference
    table_range = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(end_col)}{end_row}"
    
    # Create table
    table_name = f"Table_{sheet.title.replace(' ', '_')[:20]}"
    
    # Remove existing table if it exists
    for existing_table in list(sheet.tables.keys()):
        if existing_table.startswith("Table_"):
            del sheet.tables[existing_table]
    
    table = Table(displayName=table_name, ref=table_range)
    
    style = TableStyleInfo(
        name="TableStyleMedium9",
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,
        showColumnStripes=False
    )
    table.tableStyleInfo = style
    
    sheet.add_table(table)
    
    # Auto-adjust column widths
    for column in sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 60)
        sheet.column_dimensions[column_letter].width = adjusted_width
    
    # Style header row
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    for cell in sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Freeze panes
    sheet.freeze_panes = sheet['B2']
    
    print(f"    ✓ Applied table formatting to range: {table_range}")
    
    return table

def backup_original_sheets(writer, all_original_sheets):
    """Backup original backend sheets into the output workbook."""
    workbook = writer.book
    
    print(f"\n📋 Backing up original backend pages...")
    
    for sheet_name, df in all_original_sheets.items():
        # Clean sheet name (Excel has 31 char limit)
        clean_sheet_name = sheet_name[:31]
        
        # Check if sheet already exists and rename if needed
        final_sheet_name = clean_sheet_name
        counter = 1
        while final_sheet_name in workbook.sheetnames:
            final_sheet_name = f"{clean_sheet_name[:27]}_{counter}"
            counter += 1
        
        # Write dataframe to sheet
        df.to_excel(writer, sheet_name=final_sheet_name, index=False)
        
        # Get the sheet and format it
        sheet = workbook[final_sheet_name]
        
        # Auto-adjust column widths
        for column in sheet.columns:
            max_length = 0
            column_letter = column[0].column_letter
            for cell in column:
                try:
                    if len(str(cell.value)) > max_length:
                        max_length = len(str(cell.value))
                except:
                    pass
            adjusted_width = min(max_length + 2, 50)
            sheet.column_dimensions[column_letter].width = adjusted_width
        
        # Style header
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        for cell in sheet[1]:
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = Alignment(horizontal='center', vertical='center')
        
        # Freeze header row
        sheet.freeze_panes = sheet['A2']
        
        print(f"  ✓ Backed up '{sheet_name}' -> sheet '{final_sheet_name}' ({len(df)} rows)")

def print_summary(draw_types_maps, version_groups, 
                  output_file, version_num, all_original_sheets, summary_columns, 
                  missing_report, unique_benchmarks, duplicate_report, 
                  missing_backends, single_version_backends, baseline_version, 
                  compare_versions, comparison_type):
    """Print a summary of the analysis."""
    print("\n" + "="*60)
    print("✅ Analysis complete!")
    
    print(f"\n📊 Versions processed: {len(version_groups)}")
    for version in version_groups.keys():
        num_backends = len(version_groups[version]['dataframes'])
        print(f"  - {version}: {num_backends} backends")
    
    print(f"\n🎯 Baseline version: {baseline_version}")
    if compare_versions:
        print(f"   Compare versions: {', '.join(compare_versions)}")
    
    # Print comparison type
    comparison_type_names = {
        "same_skia_different_api": "Same Skia, Different API",
        "same_api_different_skia": "Same API, Different Skia",
        "same_skia_same_api": "Same Skia and Same API",
        "mixed_versions": "Mixed Versions",
        "single": "Single Version"
    }
    print(f"\n📋 Comparison type: {comparison_type_names.get(comparison_type, comparison_type)}")
    
    # Print column/backend coverage summary
    if missing_backends:
        print(f"\n⚠️  Backends missing from some versions: {len(missing_backends)}")
        for backend, info in missing_backends.items():
            print(f"    - '{backend}' missing in: {', '.join(info['missing_in'])}")
    
    if single_version_backends:
        print(f"\n📌 Backends unique to single version: {len(single_version_backends)}")
    
    # Print duplicate benchmark summary
    if duplicate_report:
        total_duplicates = sum(len(sheets) for sheets in duplicate_report.values())
        print(f"\n⚠️  Duplicate Benchmarks Found: {total_duplicates} sheets with duplicates")
    
    # Print missing benchmark summary
    if missing_report:
        total_missing = sum(len(m) for m in missing_report.values())
        print(f"\n⚠️  Missing Benchmarks: {total_missing} total missing entries across versions")
    
    if unique_benchmarks:
        print(f"\n📌 Benchmarks unique to single version: {len(unique_benchmarks)}")
    
    print(f"\n📋 Original sheets backed up: {len(all_original_sheets)}")
    
    if summary_columns:
        print(f"\n📋 Summary columns extracted from existing comparison: {len(summary_columns)}")
    
    print(f"\n📁 Output file: {output_file}")
    print(f"   Version number: v{version_num}")
    
    print("\n📑 Sheets in output workbook:")
    for version in compare_versions:
        print(f"  - {version}_comparison (compare version page)")
    
    # Only show cross-version sheet if it was created
    if len(version_groups) > 1 and comparison_type != "same_skia_same_api":
        print("  - cross_version_comparison (baseline + compare versions side-by-side)")
    elif comparison_type == "same_skia_same_api":
        print("  - ℹ️ No cross-version comparison (all versions have same Skia and API)")
    elif len(version_groups) <= 1:
        print("  - ℹ️ No cross-version comparison (only one version found)")
    
    print("  - backend_version (original data backups - unfiltered)")
    
    print("\n💡 Tips for using the Excel file:")
    print("  1. Compare version pages show benchmarks that exist in ALL backends of that version")
    if len(version_groups) > 1 and comparison_type != "same_skia_same_api":
        print("  2. Cross-version page shows baseline vs compare versions side-by-side")
    print("  3. Original backup sheets contain complete unfiltered data")
    print("  4. Use drop-down arrows in headers to sort/filter data")
    print("  5. First row and column are frozen for easy scrolling")

def main():
    """Main function to orchestrate the script."""
    print("="*60)
    print(f"Benchmark Analysis Tool - Version {VERSION}")
    print("="*60)
    
    # Validate arguments
    excel_files = validate_arguments()
    
    print(f"\n📄 Excel files ({len(excel_files)}):")
    for excel_file in excel_files:
        print(f"  - {excel_file}")
    
    # Read all Excel files
    (all_dataframes, all_backends, version_info, version_groups, 
     all_original_sheets, primary_summary_columns, missing_report, 
     unique_benchmarks, duplicate_report, missing_backends, 
     single_version_backends, version_benchmarks, baseline_version, 
     compare_versions, comparison_type, summary_columns_by_version, 
     has_comparison_page) = read_multiple_excel_files(excel_files)
    
    # Get unique benches for JSON analysis
    all_benches = set()
    for df in all_dataframes.values():
        all_benches.update(df['bench'].tolist())
    print(f"\n📋 Found {len(all_benches)} unique benchmarks")
    
    # Analyze ftrace files (if folders provided)
    draw_types_maps = []
    
    # Generate output filename with baseline version
    output_file, version_num = generate_output_filename(baseline_version)
    
    # Write to Excel
    print(f"\n💾 Generating Excel workbook: {output_file} (version {version_num})")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        
        # Create comparison pages for each compare version
        for compare_version in compare_versions:
            # Find the full version that contains this compare version
            full_version = None
            for v in version_groups.keys():
                if compare_version in v:
                    full_version = v
                    break
            
            if full_version and full_version in version_groups:
                print(f"\n  Creating comparison page for version: {compare_version}")
                version_data = version_groups[full_version]
                missing_for_version = missing_report.get(full_version, set())
                
                # Use version-specific summary columns if available
                version_summary = version_data.get('summary_columns', primary_summary_columns)
                
                version_df = create_version_comparison_page(
                    version_data, full_version, 
                    version_summary, missing_for_version, version_benchmarks,
                    baseline_version, is_baseline=False
                )
                
                if not version_df.empty:
                    sheet_name = f"{compare_version}_comparison"[:31]
                    write_dataframe_with_formulas(writer, sheet_name, version_df)
                    
                    workbook = writer.book
                    sheet = workbook[sheet_name]
                    apply_table_formatting_to_sheet(sheet, version_df)
                    print(f"    ✓ Created '{sheet_name}' with {len(version_df)} benchmarks, {len(version_df.columns)} columns")
                else:
                    print(f"    ✗ Skipping '{compare_version}_comparison' - no common benchmarks found")
        
        # Create cross-version comparison page if multiple versions AND different versions exist
        if len(version_groups) > 1 and comparison_type != "same_skia_same_api":
            print(f"\n  Creating cross-version comparison page")
            cross_version_df = create_cross_version_page(
                version_groups, all_backends, primary_summary_columns, missing_report, 
                version_benchmarks, baseline_version, compare_versions, comparison_type
            )
            if cross_version_df is not None and not cross_version_df.empty:
                write_dataframe_with_formulas(writer, 'cross_version_comparison', cross_version_df)
                
                workbook = writer.book
                sheet = workbook['cross_version_comparison']
                apply_table_formatting_to_sheet(sheet, cross_version_df)
                print(f"    ✓ Created 'cross_version_comparison' with {len(cross_version_df)} benchmarks, {len(cross_version_df.columns)} columns")
            else:
                print(f"    ✗ Skipping 'cross_version_comparison' - no common benchmarks across all versions")
        else:
            if comparison_type == "same_skia_same_api":
                print(f"\n  ℹ️ Skipping cross-version comparison - all versions have same Skia and API")
            elif len(version_groups) <= 1:
                print(f"\n  ℹ️ Skipping cross-version comparison - only one version found")
        
        # Backup original backend sheets
        backup_original_sheets(writer, all_original_sheets)
    
    # Print summary - now with comparison_type
    print_summary(draw_types_maps, version_groups, 
                  output_file, version_num, all_original_sheets, primary_summary_columns, 
                  missing_report, unique_benchmarks, duplicate_report, 
                  missing_backends, single_version_backends, baseline_version, 
                  compare_versions, comparison_type)
    
    print("\n" + "="*60)
    print(f"Benchmark Analysis Tool v{VERSION} - Execution Complete")
    print("="*60)

if __name__ == "__main__":
    main()