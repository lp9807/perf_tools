#!/usr/bin/env python3
import sys
import os
import json
import csv
import re

# Your specified string-to-bucket classification map
CATEGORY_MAP = {
    '::snap(': 'PipelineCreation',
    '::addRecording': 'RSCreation',
    '::SubmitToGpu': 'Submission',
    '::checkForFinishedWork': 'Execution'
}

def parse_filename_attributes(file_name):
    """
    Parses attributes from filenames matching: [name]-[backend]-m[skia version]-api[SDK version](-v[iteration]).[ext]
    Example 1: 'render_test-vulkan-m124-api34.json'
    Example 2: 'render_test-vulkan-m124-api34-v2.json'
    """
    # Updated regex pattern: 
    # (-v(?P<iteration>.+?))? makes the entire iteration block optional
    pattern = re.compile(
        r'^(?P<name>.+?)-(?P<backend>.+?)-m(?P<skia_ver>.+?)-api(?P<sdk_ver>\d+)(?:-v(?P<iteration>.+?))?\.[^.]+$'
    )
    match = pattern.match(file_name)
    
    if match:
        return (
            match.group('name'),
            match.group('backend'),
            match.group('skia_ver'),
            match.group('sdk_ver'),
            match.group('iteration') if match.group('iteration') else ""
        )
    else:
        # Fallback fields if a file in the folder does not match the naming convention strictly
        return (file_name, "Unknown", "Unknown", "Unknown", "")

def parse_single_trace_file(file_path):
    file_name = os.path.basename(file_path)
    
    # Track the extracted filename attributes (now including iteration)
    name_attr, backend_attr, skia_attr, sdk_attr, iter_attr = parse_filename_attributes(file_name)
    
    # State tracking metrics isolated for this specific file run
    open_sections = {}    # For B/E pairs: (pid, tid, name) -> list of start timestamps
    last_instant_ts = {}  # For I sequences: (pid, tid, name) -> previous instant timestamp
    
    # Dynamic metric aggregation maps for each category bucket
    category_data = {cat_name: [] for cat_name in CATEGORY_MAP.values()}

    try:
        with open(file_path, 'r') as f:
            data = json.load(f)
    except json.JSONDecodeError as e:
        print(f"[ERROR] {file_name}: Invalid JSON format ({e})")
        return None
    except Exception as e:
        print(f"[ERROR] {file_name}: Could not read file ({e})")
        return None

    if isinstance(data, dict) and "traceEvents" in data:
        events = data["traceEvents"]
    elif isinstance(data, list):
        events = data
    else:
        print(f"[ERROR] {file_name}: Structural root format not supported.")
        return None

    for event in events:
        if not isinstance(event, dict):
            continue
            
        phase = event.get("ph")
        name = event.get("name", "")
        ts = event.get("ts")
        pid = event.get("pid")
        tid = event.get("tid")
        dur = event.get("dur")
        
        if phase is None or ts is None:
            continue

        # Determine which visual bucket this event belongs to based on partial substring map matching
        assigned_category = None
        for substring, cat_name in CATEGORY_MAP.items():
            if substring in name:
                assigned_category = cat_name
                break  # First matching rule wins
                
        if not assigned_category:
            continue

        # --- Case 1: Complete Event ("X") ---
        if phase == "X":
            if dur is not None:
                category_data[assigned_category].append(float(dur) / 1000.0)

        # --- Case 2: Begin Event ("B") ---
        elif phase == "B":
            thread_key = (pid, tid, name)
            if thread_key not in open_sections:
                open_sections[thread_key] = []
            open_sections[thread_key].append(ts)

        # --- Case 3: End Event ("E") ---
        elif phase == "E":
            thread_key = (pid, tid, name)
            if thread_key in open_sections and open_sections[thread_key]:
                start_ts = open_sections[thread_key].pop()
                category_data[assigned_category].append((ts - start_ts) / 1000.0)
                if not open_sections[thread_key]:
                    del open_sections[thread_key]

        # --- Case 4: Instant Event ("I") ---
        elif phase == "I":
            thread_key = (pid, tid, name)
            if thread_key in last_instant_ts:
                prev_ts = last_instant_ts[thread_key]
                category_data[assigned_category].append((ts - prev_ts) / 1000.0)
                
            last_instant_ts[thread_key] = ts

    # Compile the final flat single row array values for this file
    row_cells = [file_name, name_attr, backend_attr, skia_attr, sdk_attr, iter_attr]
    
    # Dynamically build max calculation columns matching our sorted category list
    for cat_name in sorted(CATEGORY_MAP.values()):
        durations = category_data[cat_name]
        if durations:
            max_duration = max(durations)
            row_cells.append(f"{max_duration:.4f}")
        else:
            row_cells.append("0.0000")
            
    return row_cells


def scan_directory_to_max_csv(folder_path):
    if not os.path.isdir(folder_path):
        print(f"Error: Target path '{folder_path}' is not a valid directory.")
        sys.exit(1)
        
    files = sorted(os.listdir(folder_path))
    
    # Exclude 'trace_report.csv' from self-parsing
    trace_files = [
        f for f in files 
        if os.path.isfile(os.path.join(folder_path, f)) and f != 'trace_report.csv'
    ]
    
    if not trace_files:
        print("No source trace files found to process inside the target folder.")
        return

    output_csv_path = os.path.join(folder_path, 'trace_report.csv')

    print(f"Scanning folder   : {os.path.abspath(folder_path)}")
    print(f"Writing report to : {os.path.abspath(output_csv_path)}...")

    with open(output_csv_path, mode='w', newline='', encoding='utf-8') as csv_file:
        writer = csv.writer(csv_file)
        
        # New Column Layout adding 'Iteration'
        headers = ["File Name", "Name", "Backend", "Skia Version", "SDK Version", "Iteration"]
        for cat_name in sorted(CATEGORY_MAP.values()):
            headers.append(f"{cat_name} Max Time (ms)")
        writer.writerow(headers)
        
        rows_written = 0
        for file_name in trace_files:
            full_path = os.path.join(folder_path, file_name)
            row_data = parse_single_trace_file(full_path)
            
            if row_data:
                writer.writerow(row_data)
                rows_written += 1

    print(f"Success! Processed {len(trace_files)} files. Generated '{output_csv_path}' with {rows_written} rows.")


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python3 parse_trace_final.py <folder_path>")
        print("Example: python3 parse_trace_final.py ./perf_logs/")
        sys.exit(1)
        
    target_folder = sys.argv[1]
    
    scan_directory_to_max_csv(target_folder)
