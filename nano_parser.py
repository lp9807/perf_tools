#!/usr/bin/env python3
"""
Script to process CSV benchmark files and generate an Excel workbook with comparison.
Usage: python script.py <folder_path1> [folder_path2] <csv_file1> [csv_file2 ...]
Maximum 2 folder paths can be provided as the first parameters.
If a folder parameter is not a valid directory, trace parser logic will be skipped for that folder.
"""

import sys
import os
import json
import re
import pandas as pd
from pathlib import Path
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side
from openpyxl.comments import Comment
from openpyxl.worksheet.table import Table, TableStyleInfo
import warnings
warnings.filterwarnings('ignore')

def validate_arguments():
    """Validate command line arguments."""
    if len(sys.argv) < 3:
        print("Error: At least 3 arguments required (at least one folder and one CSV file)")
        print(f"Usage: {sys.argv[0]} <folder_path1> [folder_path2] <csv_file1> [csv_file2 ...]")
        print("Note: Maximum 2 folder paths can be provided as the first parameters")
        sys.exit(1)
    
    # Parse arguments: first 1-2 could be folders, rest are CSV files
    folder_paths = []
    csv_files = []
    
    # Check each argument to determine if it's a folder or CSV
    for arg in sys.argv[1:]:
        if arg.lower().endswith('.csv'):
            csv_files.append(arg)
        elif len(folder_paths) < 2:  # Max 2 folders
            folder_paths.append(arg)
        else:
            # If we already have 2 folders and this isn't a CSV, it's an error
            print(f"Error: Maximum 2 folder parameters allowed. Extra parameter: '{arg}'")
            sys.exit(1)
    
    if len(folder_paths) == 0:
        print("Error: At least one folder path must be provided")
        sys.exit(1)
    
    if len(csv_files) == 0:
        print("Error: At least one CSV file must be provided")
        sys.exit(1)
    
    # Check if folders exist (warn but don't exit)
    folder_exists = []
    for folder_path in folder_paths:
        exists = os.path.isdir(folder_path)
        folder_exists.append(exists)
        if not exists:
            print(f"Warning: Folder path '{folder_path}' does not exist - trace analysis will be skipped for this folder")
    
    # Check CSV files exist
    for csv_file in csv_files:
        if not os.path.isfile(csv_file):
            print(f"Error: CSV file '{csv_file}' does not exist")
            sys.exit(1)
    
    return folder_paths, csv_files, folder_exists

def detect_backend_type(folder_path):
    """Detect backend type from folder name."""
    folder_name = Path(folder_path).name.lower()
    if 'gr' in folder_name:
        return 'graphite'
    else:
        return 'ganesh'

def read_csv_files(csv_files):
    """Read CSV files and validate they contain 'bench' and 'mean' columns."""
    dataframes = {}
    for csv_file in csv_files:
        try:
            df = pd.read_csv(csv_file)
            if 'bench' not in df.columns or 'mean' not in df.columns:
                print(f"Error: CSV file '{csv_file}' must contain 'bench' and 'mean' columns")
                sys.exit(1)
            dataframes[csv_file] = df
        except Exception as e:
            print(f"Error reading CSV file '{csv_file}': {e}")
            sys.exit(1)
    return dataframes

def analyze_ftrace_files_ganesh(folder_path, benches, folder_name):
    """Analyze ftrace JSON files for Ganesh backend."""
    folder = Path(folder_path)
    draw_types_map = {}
    
    # Pattern to match SurfaceDrawContext::draw* functions
    draw_pattern = re.compile(r'SurfaceDrawContext::draw([a-zA-Z]+)')
    # Pattern to match GrDrawingManager::flush specifically
    flush_pattern = re.compile(r'GrDrawingManager::flush')
    
    for bench in benches:
        json_file = folder / f"{bench}.json"
        draw_counts = {}
        flush_count = 0
        
        if json_file.exists():
            try:
                with open(json_file, 'r') as f:
                    data = json.load(f)
                
                # Handle different ftrace JSON structures
                def process_trace(trace_data):
                    nonlocal flush_count
                    if isinstance(trace_data, dict):
                        # Check for function name in common fields
                        func_name = trace_data.get('func') or trace_data.get('function') or trace_data.get('name')
                        if func_name:
                            # Count draw types
                            draw_match = draw_pattern.search(func_name)
                            if draw_match:
                                draw_type = draw_match.group(1)
                                draw_counts[draw_type] = draw_counts.get(draw_type, 0) + 1
                            
                            # Count GrDrawingManager::flush calls
                            if flush_pattern.search(func_name):
                                flush_count += 1
                        
                        # Recursively process nested structures
                        for value in trace_data.values():
                            if isinstance(value, (dict, list)):
                                process_trace(value)
                    elif isinstance(trace_data, list):
                        for item in trace_data:
                            process_trace(item)
                
                process_trace(data)
                
                if draw_counts and flush_count > 0:
                    # Calculate division result for each draw type: total_count / flush_count
                    sorted_counts = sorted(draw_counts.items(), key=lambda x: x[1], reverse=True)
                    result_parts = []
                    for draw_type, count in sorted_counts:
                        division_result = count / flush_count
                        # Format with 2 decimal places for the division result
                        result_parts.append(f"{draw_type}:{division_result:.2f}")
                    draw_types_map[bench] = ', '.join(result_parts)
                elif draw_counts:
                    draw_types_map[bench] = f"No flush calls found. Draw types: {', '.join([f'{dt}:{cnt}' for dt, cnt in draw_counts.items()])}"
                else:
                    draw_types_map[bench] = "No draw functions found"
                    
            except Exception as e:
                draw_types_map[bench] = f"Error parsing JSON: {str(e)[:50]}"
        else:
            draw_types_map[bench] = f"JSON file not found: {json_file}"
    
    return draw_types_map

def analyze_ftrace_files_graphite(folder_path, benches):
    """Analyze ftrace JSON files for Graphite backend (TODO feature)."""
    folder = Path(folder_path)
    draw_types_map = {}
    
    for bench in benches:
        json_file = folder / f"{bench}.json"
        
        if json_file.exists():
            # TODO: Implement Graphite-specific trace analysis
            draw_types_map[bench] = "TODO: Graphite trace analysis not yet implemented"
        else:
            draw_types_map[bench] = f"JSON file not found: {json_file}"
    
    return draw_types_map

def analyze_ftrace_files(folder_path, benches, folder_exists, folder_name):
    """Main dispatcher for ftrace analysis based on backend type.
    
    Args:
        folder_path: Path to folder containing JSON files
        benches: List of benchmark names
        folder_exists: Boolean indicating if folder exists
        folder_name: Name for display purposes
    
    Returns:
        Dictionary mapping bench to draw_types summary, or None if analysis skipped
    """
    if not folder_exists:
        print(f"\n⚠️  Skipping trace file analysis for '{folder_name}': Folder does not exist")
        return None
    
    backend_type = detect_backend_type(folder_path)
    print(f"  Backend type for '{folder_name}': {backend_type.upper()}")
    
    if backend_type == 'graphite':
        return analyze_ftrace_files_graphite(folder_path, benches)
    else:  # ganesh
        return analyze_ftrace_files_ganesh(folder_path, benches, folder_name)

def create_comparison_page(dataframes, draw_types_maps, folder_paths):
    """Create the comparison dataframe with summary columns for each folder."""
    # Get unique benches from all CSVs
    all_benches = set()
    for df in dataframes.values():
        all_benches.update(df['bench'].tolist())
    
    benches = sorted(list(all_benches))
    
    # Create ordered ID column
    ordered_ids = list(range(1, len(benches) + 1))
    
    # Prepare comparison data
    comparison_data = {
        'ID': ordered_ids,
        'Bench': benches
    }
    
    # Add mean columns for each CSV (named after filename/backend)
    for csv_file, df in dataframes.items():
        backend_name = Path(csv_file).stem  # Remove .csv extension
        mean_dict = dict(zip(df['bench'], df['mean']))
        comparison_data[backend_name] = [mean_dict.get(bench, float('nan')) for bench in benches]
    
    # Add summary columns for each folder
    for idx, folder_path in enumerate(folder_paths):
        folder_name = Path(folder_path).name
        if folder_name == '.' or folder_name == '/':
            folder_name = f"folder_{idx + 1}"
        
        # Get the draw_types map for this folder (None if analysis was skipped)
        draw_types_map = draw_types_maps[idx] if idx < len(draw_types_maps) else None
        
        # Column name: "summary of [folder_name]"
        summary_col_name = f"summary of {folder_name}"
        
        if draw_types_map is not None:
            comparison_data[summary_col_name] = [draw_types_map.get(bench, "Bench not found in trace files") for bench in benches]
        else:
            comparison_data[summary_col_name] = ["Trace analysis skipped (folder not found)" for _ in benches]
    
    # Create dataframe
    comparison_df = pd.DataFrame(comparison_data)
    
    return comparison_df

def add_excel_ratio_formulas(writer, comparison_df, dataframes):
    """Add Excel formulas for all ratio columns."""
    workbook = writer.book
    sheet = workbook['comparison']
    
    # Define ratio configurations
    ratio_configs = [
        {
            'name': 'grdawn_vk vs glesdmsaa',
            'numerator_patterns': ['grdawn_vk', 'grdawn-vk', 'grdawnvk'],
            'denominator_patterns': ['glesdmsaa', 'gles-dmsaa']
        },
        {
            'name': 'vkdmsaa vs glesdmsaa',
            'numerator_patterns': ['vkdmsaa', 'vk-dmsaa', 'vkdmsaa'],
            'denominator_patterns': ['glesdmsaa', 'gles-dmsaa']
        },
        {
            'name': 'grvk vs grdawn_vk',
            'numerator_patterns': ['grvk', 'gr-vk', 'grvk'],
            'denominator_patterns': ['grdawn_vk', 'grdawn-vk', 'grdawnvk']
        }
    ]
    
    # Get all headers
    headers = [cell.value for cell in sheet[1]]
    
    # Find column indices for each backend
    column_map = {}
    for col_idx, header in enumerate(headers, 1):
        if header and isinstance(header, str):
            header_lower = header.lower()
            column_map[header_lower] = col_idx
    
    # Process each ratio configuration
    for config in ratio_configs:
        ratio_name = config['name']
        
        # Find numerator column
        numerator_col = None
        numerator_name = None
        for pattern in config['numerator_patterns']:
            for col_name, col_idx in column_map.items():
                if pattern in col_name:
                    numerator_col = col_idx
                    numerator_name = [h for h in headers if column_map.get(h.lower()) == col_idx][0]
                    break
            if numerator_col:
                break
        
        # Find denominator column
        denominator_col = None
        denominator_name = None
        for pattern in config['denominator_patterns']:
            for col_name, col_idx in column_map.items():
                if pattern in col_name:
                    denominator_col = col_idx
                    denominator_name = [h for h in headers if column_map.get(h.lower()) == col_idx][0]
                    break
            if denominator_col:
                break
        
        if not numerator_col or not denominator_col:
            print(f"\n⚠️  Warning: Could not find columns for ratio '{ratio_name}'")
            if not numerator_col:
                print(f"  - Numerator not found (patterns: {config['numerator_patterns']})")
            if not denominator_col:
                print(f"  - Denominator not found (patterns: {config['denominator_patterns']})")
            continue
        
        # Check if ratio column already exists
        if ratio_name in headers:
            ratio_col_idx = headers.index(ratio_name) + 1
        else:
            # Add column before summary columns if possible
            summary_cols = [col for col in headers if col and col.startswith('summary of')]
            if summary_cols:
                first_summary_idx = headers.index(summary_cols[0]) + 1
                sheet.insert_cols(first_summary_idx)
                ratio_col_idx = first_summary_idx
                sheet.cell(row=1, column=ratio_col_idx, value=ratio_name)
                # Update headers list
                headers.insert(first_summary_idx - 1, ratio_name)
            else:
                # Add at the end
                ratio_col_idx = len(headers) + 1
                sheet.cell(row=1, column=ratio_col_idx, value=ratio_name)
        
        # Add Excel formula for each row
        print(f"\nAdding Excel formula for ratio column: {ratio_name}")
        print(f"  = {numerator_name} / {denominator_name}")
        
        formula_count = 0
        for row_idx in range(2, len(comparison_df) + 2):
            # Create Excel formula
            formula = f"={get_column_letter(numerator_col)}{row_idx}/{get_column_letter(denominator_col)}{row_idx}"
            
            # Set the formula in the cell
            cell = sheet.cell(row=row_idx, column=ratio_col_idx)
            cell.value = formula
            cell.number_format = "0.000"  # Format with 3 decimal places
            
            # Add a comment explaining the formula
            cell.comment = Comment(f"Formula: {numerator_name} / {denominator_name}", "Script")
            formula_count += 1
        
        # Auto-adjust column width
        column_letter = get_column_letter(ratio_col_idx)
        sheet.column_dimensions[column_letter].width = 20
        
        print(f"  - Added formula to {formula_count} rows")
        print(f"  - Column: {column_letter}")
        print(f"  - Example formula: {formula}")

def apply_table_formatting(writer, comparison_df):
    """Apply Excel table formatting to the comparison sheet for sorting/filtering."""
    workbook = writer.book
    sheet = workbook['comparison']
    
    # Determine the range of the table
    start_row = 1
    start_col = 1
    end_row = len(comparison_df) + 1  # +1 for header
    
    # Find the last column with data
    max_col = 1
    for row in sheet.iter_rows(max_row=1):
        for cell in row:
            if cell.value:
                max_col = max(max_col, cell.column)
    
    # Create table range reference
    table_range = f"{get_column_letter(start_col)}{start_row}:{get_column_letter(max_col)}{end_row}"
    
    # Create table
    table_name = "ComparisonTable"
    
    # Remove existing table if it exists
    if table_name in sheet.tables:
        del sheet.tables[table_name]
    
    table = Table(displayName=table_name, ref=table_range)
    
    # Define table style
    style = TableStyleInfo(
        name="TableStyleMedium9",  # Clean, professional style
        showFirstColumn=False,
        showLastColumn=False,
        showRowStripes=True,  # Alternating row colors
        showColumnStripes=False
    )
    table.tableStyleInfo = style
    
    # Add table to sheet
    sheet.add_table(table)
    
    print(f"\n📊 Applied Excel table formatting to range: {table_range}")
    print(f"   - Table name: {table_name}")
    print("   - Features: Sorting, Filtering, Alternating row colors")
    
    return table

def format_excel_workbook(writer, comparison_df, dataframes, folder_paths):
    """Format the Excel workbook with styling."""
    workbook = writer.book
    
    # Format comparison sheet
    comparison_sheet = workbook['comparison']
    
    # Apply table formatting (this enables sorting/filtering)
    apply_table_formatting(writer, comparison_df)
    
    # Auto-adjust column widths (do this after table creation)
    for column in comparison_sheet.columns:
        max_length = 0
        column_letter = column[0].column_letter
        for cell in column:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = min(max_length + 2, 60)
        comparison_sheet.column_dimensions[column_letter].width = adjusted_width
    
    # Freeze the header row and first column for easy scrolling
    comparison_sheet.freeze_panes = comparison_sheet['B2']  # Freeze row 1 and column A
    
    print(f"   - Frozen panes: Row 1 (headers) and Column A (ID/Bench)")
    
    # Format individual CSV sheets
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    
    for csv_file in dataframes.keys():
        sheet_name = Path(csv_file).stem[:31]  # Excel sheet name max 31 chars
        if sheet_name in workbook.sheetnames:
            sheet = workbook[sheet_name]
            
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
            for cell in sheet[1]:
                cell.font = header_font
                cell.fill = header_fill
                cell.alignment = Alignment(horizontal='center', vertical='center')
            
            # Freeze header row for CSV sheets too
            sheet.freeze_panes = sheet['A2']

def print_summary(folder_paths, folder_exists_list, draw_types_maps, dataframes, comparison_df, ratio_columns):
    """Print a summary of the analysis."""
    print("\n" + "="*60)
    print("✅ Analysis complete!")
    
    print("\n📁 Folders analyzed:")
    for idx, folder_path in enumerate(folder_paths):
        folder_name = Path(folder_path).name
        exists = folder_exists_list[idx] if idx < len(folder_exists_list) else False
        status = "✅ Analyzed" if exists else "⚠️  Skipped (not found)"
        print(f"  {idx + 1}. {folder_name}: {status}")
        
        # Show statistics for this folder if analysis was done
        if exists and draw_types_maps and idx < len(draw_types_maps) and draw_types_maps[idx] is not None:
            benches_with_info = sum(1 for v in draw_types_maps[idx].values() 
                                   if not v.startswith("JSON file not found") 
                                   and not v.startswith("Error")
                                   and v != "Bench not found in trace files")
            total_benches = len(draw_types_maps[idx])
            print(f"     Benchmarks with trace info: {benches_with_info}/{total_benches}")
    
    print(f"\n📊 CSV files processed: {len(dataframes)}")
    for csv_file in dataframes.keys():
        print(f"  - {Path(csv_file).name}: {len(dataframes[csv_file])} rows")
    
    print(f"\n📈 Comparison page: {len(comparison_df)} benchmarks")
    print(f"   Base columns: ID, Bench, backend columns ({len(dataframes)}), summary columns ({len(folder_paths)})")
    print(f"   Ratio columns added: {len(ratio_columns)} ({', '.join(ratio_columns)})")
    
    # Note about table features
    print("\n📊 Excel Table Features Available:")
    print("  - Click drop-down arrows in headers to sort/filter")
    print("  - Alternating row colors for easy reading")
    print("  - Header row frozen for scrolling")
    print("  - Resizable and customizable table style")
    
    # Note about ratio formulas
    print("\n📐 Ratio Columns (Excel Formulas):")
    for ratio_col in ratio_columns:
        print(f"  - {ratio_col}: Calculated dynamically when Excel file is opened")

def main():
    """Main function to orchestrate the script."""
    print("Starting benchmark analysis...")
    print("="*60)
    
    # Validate arguments and get folder paths and CSV files
    folder_paths, csv_files, folder_exists_list = validate_arguments()
    
    print(f"\n📁 Folder paths ({len(folder_paths)}):")
    for folder_path in folder_paths:
        print(f"  - {folder_path}")
    
    print(f"\n📄 CSV files ({len(csv_files)}):")
    for csv_file in csv_files:
        print(f"  - {csv_file}")
    
    # Read CSV files
    print("\n📖 Reading CSV files...")
    dataframes = read_csv_files(csv_files)
    
    # Get unique benches for JSON analysis
    all_benches = set()
    for df in dataframes.values():
        all_benches.update(df['bench'].tolist())
    print(f"\n📋 Found {len(all_benches)} unique benchmarks")
    
    # Analyze ftrace JSON files for each folder
    print("\n🔍 Analyzing ftrace JSON files...")
    draw_types_maps = []
    for idx, folder_path in enumerate(folder_paths):
        folder_name = Path(folder_path).name
        print(f"\n  Processing folder {idx + 1}: {folder_name}")
        folder_exists = folder_exists_list[idx] if idx < len(folder_exists_list) else False
        draw_types_map = analyze_ftrace_files(folder_path, all_benches, folder_exists, folder_name)
        draw_types_maps.append(draw_types_map)
    
    # Create comparison dataframe with summary columns
    print("\n📊 Creating comparison page...")
    comparison_df = create_comparison_page(dataframes, draw_types_maps, folder_paths)
    
    # Generate output filename
    output_file = "combined.xlsx"
    
    # Write to Excel with multiple sheets
    print(f"\n💾 Generating Excel workbook: {output_file}")
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:
        # Write comparison sheet as first sheet
        comparison_df.to_excel(writer, sheet_name='comparison', index=False)
        
        # Write individual CSV sheets
        for csv_file, df in dataframes.items():
            sheet_name = Path(csv_file).stem[:31]  # Excel sheet name max 31 chars
            df.to_excel(writer, sheet_name=sheet_name, index=False)
        
        # Add Excel formulas for ratio columns
        add_excel_ratio_formulas(writer, comparison_df, dataframes)
        
        # Apply formatting (including table formatting)
        format_excel_workbook(writer, comparison_df, dataframes, folder_paths)
    
    # Define ratio columns for summary
    ratio_columns = ['grdawn_vk vs glesdmsaa', 'vkdmsaa vs glesdmsaa', 'grvk vs grdawn_vk']
    
    # Print summary
    print_summary(folder_paths, folder_exists_list, draw_types_maps, dataframes, comparison_df, ratio_columns)
    
    # Show sample of the comparison table structure (only for non-formula columns)
    print("\n📋 Sample of comparison table structure (first 5 rows, values from CSV data only):")
    display_cols = ['ID', 'Bench']
    
    # Add first two backend columns for sample (these contain actual values)
    backend_cols = [col for col in comparison_df.columns 
                   if col not in ['ID', 'Bench'] 
                   and not col.startswith('summary of')][:2]
    display_cols.extend(backend_cols)
    
    # Add first summary column (contains trace analysis results)
    summary_cols = [col for col in comparison_df.columns if col.startswith('summary of')]
    if summary_cols:
        display_cols.append(summary_cols[0])
    
    # Display the dataframe sample (without ratio columns which are Excel-only formulas)
    print(comparison_df[display_cols].head().to_string())
    print("\n   Note: Ratio columns (grdawn_vk vs glesdmsaa, vkdmsaa vs glesdmsaa, grvk vs grdawn_vk)")
    print("         are Excel formulas and not shown in this preview. They will appear when you")
    print("         open the Excel file and will automatically calculate based on the backend columns.")
    
    print("\n📁 Output file details:")
    print(f"  - File: {output_file}")
    print(f"  - Total sheets: {len(dataframes) + 1}")
    print(f"  - First sheet: comparison ({len(comparison_df)} benchmarks)")
    print(f"  - Summary columns: {len(summary_cols)}")
    for col in summary_cols:
        print(f"    * {col}")
    print(f"  - Ratio columns: {len(ratio_columns)} (Excel formulas)")
    for col in ratio_columns:
        print(f"    * {col}")
    
    print("\n💡 Tips for using the Excel file:")
    print("  1. Use the drop-down arrows in any header to sort or filter data")
    print("  2. Ratio columns automatically calculate: grdawn_vk/glesdmsaa, vkdmsaa/glesdmsaa, grvk/grdawn_vk")
    print("  3. The first row and column are frozen for easy scrolling")
    print("  4. Table formatting updates automatically when you add/remove data")
    print("  5. All formulas recalculate when source data changes")
    print("  6. Summary columns show trace analysis results (draw types per flush)")
    
    print("\n" + "="*60)

if __name__ == "__main__":
    main()