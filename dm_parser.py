#!/usr/bin/env python3
import sys
import os
import json
import csv
import re

# Map applied if filename backend CONTAINS 'gr'
CATEGORY_MAP_GR = {
    '::snap(': 'PipelineCreation',
    '::addRecording': 'RSCreation',
    '::submitToGpu': 'Submission',
    '::checkForFinishedWork': 'Execution'
}

# Map applied if filename backend DOES NOT contain 'gr'
CATEGORY_MAP_NON_GR = {
    '::flush': 'PipelineCreation',
    '::submitCommandBuffer': 'Submission',
    'Gpu::ReadPixels': 'Execution'
}

# Unique bucket union list to build uniform CSV headers
ALL_CATEGORIES = sorted(list(set(list(CATEGORY_MAP_GR.values()) + list(CATEGORY_MAP_NON_GR.values()))))

# Threshold gap (in microseconds) to merge sequential top-level slices.
SEQUENTIAL_MERGE_THRESHOLD_US = 1000.0

def parse_filename_attributes(file_name):
    """
    Parses attributes from filenames matching: [name]-[backend]-m[skia version]-api[SDK version](-v[iteration]).[ext]
    """
    pattern = re.compile(
            r'^(?P<name>[^-.]+)'                 # 1. Name: Required (stops at '-' or '.')
            r'(?:-(?P<backend>[^-.]+))?'         # 2. Backend: Optional
            r'(?:-m(?P<skia_ver>[^-.]+))?'       # 3. Skia: Optional
            r'(?:-api(?P<sdk_ver>\d+))?'         # 4. SDK: Optional
            r'(?:-v(?P<iteration>[^.]+))?'       # 5. Iteration: Optional
            r'(?:\.[^.]+)?$'                     # 6. Extension: Optional
    )
    match = pattern.match(file_name)
    
    if match:
        return (
            match.group('name'),
            match.group('backend') if match.group('backend') else "",
            match.group('skia_ver') if match.group('skia_ver') else "",
            match.group('sdk_ver') if match.group('sdk_ver') else "",
            match.group('iteration') if match.group('iteration') else ""
        )
    else:
        return (file_name, "", "", "", "")

def parse_single_trace_file(file_path):
    file_name = os.path.basename(file_path)
    name_attr, backend_attr, skia_attr, sdk_attr, iter_attr = parse_filename_attributes(file_name)
    
    # Conditional mapping selection based on the backend string
    if 'gr' in backend_attr.lower():
        active_category_map = CATEGORY_MAP_GR
    else:
        active_category_map = CATEGORY_MAP_NON_GR

    # State tracking metrics isolated for this specific file run
    open_sections = {}       
    last_instant_ts = {}     
    
    # Track nesting depths and top-level block boundaries per thread context
    thread_depths = {}
    thread_top_slices = {}
    
    # Initialize metric arrays for all known union categories
    category_data = {cat_name: [] for cat_name in ALL_CATEGORIES}

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

        ts_float = float(ts)
        thread_id = (pid, tid)
        
        if thread_id not in thread_depths:
            thread_depths[thread_id] = 0
            thread_top_slices[thread_id] = []

        # Map substring categories using the dynamically chosen dictionary path
        assigned_category = None
        for substring, cat_name in active_category_map.items():
            if substring in name:
                assigned_category = cat_name
                break  

        # --- Case 1: Complete Event ("X") ---
        if phase == "X":
            current_depth = thread_depths[thread_id]
            if dur is not None:
                duration_ms = float(dur) / 1000.0
                
                if assigned_category:
                    category_data[assigned_category].append(duration_ms)
                
                if current_depth == 0:
                    thread_top_slices[thread_id].append((ts_float, ts_float + float(dur)))

        # --- Case 2: Begin Event ("B") ---
        elif phase == "B":
            current_depth = thread_depths[thread_id]
            
            thread_key = (pid, tid, name)
            if thread_key not in open_sections:
                open_sections[thread_key] = []
            open_sections[thread_key].append((ts_float, current_depth))
            
            thread_depths[thread_id] += 1

        # --- Case 3: End Event ("E") ---
        elif phase == "E":
            if thread_depths[thread_id] > 0:
                thread_depths[thread_id] -= 1
                
            thread_key = (pid, tid, name)
            if thread_key in open_sections and open_sections[thread_key]:
                start_ts, depth_at_start = open_sections[thread_key].pop()
                duration_ms = (ts_float - start_ts) / 1000.0
                
                if assigned_category:
                    category_data[assigned_category].append(duration_ms)
                    
                if depth_at_start == 0:
                    thread_top_slices[thread_id].append((start_ts, ts_float))
                    
                if not open_sections[thread_key]:
                    del open_sections[thread_key]

        # --- Case 4: Instant Event ("I") ---
        elif phase == "I":
            if assigned_category:
                thread_key = (pid, tid, name)
                if thread_key in last_instant_ts:
                    prev_ts = last_instant_ts[thread_key]
                    category_data[assigned_category].append((ts_float - prev_ts) / 1000.0)
                last_instant_ts[thread_key] = ts_float

    # Calculate overall combined sequential duration metrics across all threads
    max_overall_duration_ms = 0.0

    for thread_id, slices in thread_top_slices.items():
        if not slices:
            continue
            
        slices.sort(key=lambda x: x[0])
        
        merged_slices = []
        current_start, current_end = slices[0]
        
        for next_start, next_end in slices[1:]:
            if next_start <= current_end + SEQUENTIAL_MERGE_THRESHOLD_US:
                current_end = max(current_end, next_end)
            else:
                merged_slices.append((current_start, current_end))
                current_start, current_end = next_start, next_end
        merged_slices.append((current_start, current_end))
        
        for start, end in merged_slices:
            span_ms = (end - start) / 1000.0
            if span_ms > max_overall_duration_ms:
                max_overall_duration_ms = span_ms

    # Build row cells array layout
    row_cells = [
        file_name, 
        name_attr, 
        backend_attr, 
        skia_attr, 
        sdk_attr, 
        iter_attr, 
        f"{max_overall_duration_ms:.4f}"
    ]
    
    # Write out maximums to uniform columns. 
    # If a category doesn't apply to a backend format, it cleanly lists as 0.0000
    for cat_name in ALL_CATEGORIES:
        durations = category_data[cat_name]
        if durations:
            row_cells.append(f"{max(durations):.4f}")
        else:
            row_cells.append("0.0000")
            
    return row_cells

def scan_directory_to_max_csv(folder_path):
    if not os.path.isdir(folder_path):
        print(f"Error: Target path '{folder_path}' is not a valid directory.")
        sys.exit(1)
        
    files = sorted(os.listdir(folder_path))
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
        
        # Build Table Headers with globally mapped categories
        headers = [
            "File Name", "Name", "Backend", "Skia Version", "SDK Version", "Iteration", 
            "Overall Duration (ms)"
        ]
        for cat_name in ALL_CATEGORIES:
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
        sys.exit(1)
        
    target_folder = sys.argv[1]
    
    scan_directory_to_max_csv(target_folder)
