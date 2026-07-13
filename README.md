# BENCHMARK ANALYSIS TOOL - TRACE PARSER

## SYNOPSIS
```
nano_parser.py [FOLDER_PATHS...] EXCEL_FILE [EXCEL_FILE...]
```

## DESCRIPTION
Process benchmark data from Excel files and generate comprehensive comparison reports with version tracking and trace analysis.

## USAGE

### Basic Usage
```bash
# Single file
python nano_parser.py data.xlsx

# With trace folders
python nano_parser.py traces/ data.xlsx

# Multiple versions
python nano_parser.py data_api123.xlsx data_api124.xlsx
```

## ARGUMENTS

### Excel Files (Required)
One or more Excel files containing benchmark data. Each sheet = one backend (e.g., `glesdmsaa`, `grdawn_vk`). Each sheet must have `bench` and `mean` columns.

**Version detection in filenames:**
- `api[0-9]+` → API version (e.g., `api123`)
- `m[0-9]+` → Skia version (e.g., `m456`)

### Folder Paths (Optional, 0-2)
Folders containing ftrace JSON files for draw type analysis. Folders with `gr` in name = Graphite backend, otherwise Ganesh.

## OUTPUT

### Filename Format
- **Single file**: `[basename]_benchmark_comparison_v[N].xlsx`
- **Multiple files**: `benchmark_comparison_v[N].xlsx`

### Sheets
```
[version]_comparison        # Per-version comparison
cross_version_comparison    # Cross-version comparison (if multiple versions)
[backend]_[version]         # Original data backups
```

### Comparison Page Columns
- `ID`, `Bench`
- Backend mean values
- Ratio columns: `[backend1] vs [backend2] (ratio)`
- Diff columns: `[backend1] vs [backend2] (diff)`
- `AVERAGE` row (excluded from sorting)

## VERSION DETECTION
- **Same Skia version** → Use Skia as baseline (e.g., `m123`)
- **Same API version** → Use API as baseline (e.g., `api123`)  
- **Mixed** → Use first version as baseline

## TRACE ANALYSIS SUMMARY

### Ganesh (`traces_ganesh/`)
- **Submissions**: Between `sk_gpu_test::TestContext::flushAndWaitOnSync` events
- **Output**: Draw type frequency per submission
- **Example**: `sub1: 5[Rect:3,Text:2], sub2: 3[Rect:2,Circle:1]`

### Graphite (`traces_graphite/`)
- **Submissions**: Between `submitRecordingAndWaitOnSync` and `Recorder::snap`
- **Renderer counts**: Parsed from event arguments
- **Output**: Renderer usage with flush detection
- **Example**: `sub1: 5[RendererA:3,RendererB:2|f:RendererC:1]`

## EXAMPLES
```bash
# Single file
python nano_parser.py data.xlsx

# With Ganesh traces
python nano_parser.py traces_ganesh/ data.xlsx

# Multiple versions with Graphite
python nano_parser.py traces_graphite/ data_api123.xlsx data_api124.xlsx

# Two trace folders + multiple versions
python nano_parser.py traces_ganesh/ traces_graphite/ data_api123.xlsx data_api124.xlsx
```

## NOTES
- Output filenames auto-increment to prevent overwriting
- Excel sheet names = backend names
- Version patterns: `api[0-9]+` and `m[0-9]+` are auto-detected
- Average row is excluded from table sorting

## REQUIREMENTS
- Python 3.6+
- openpyxl, pandas

# BENCHMARK ANALYSIS TOOL - CROSS VERSION COMPARISON

## USAGE
```bash
python nano_cross_compare.py <file1.xlsx> [file2.xlsx ...]
```

## INPUT FORMAT

### Sheets must contain:
| Column | Required |
|--------|----------|
| `bench` | ✓ |
| `mean` | ✓ |

### Sheet naming:
```
[backend]_[version]  (e.g., ganesh_api123, graphite_api123_m10)
```

**Backends:** `ganesh`, `graphite`, `glesdmsaa`, `grdawn_vk`, `vkdmsaa`, `grvk`

### Version detection from filename:
| Pattern | Example |
|---------|---------|
| `api[0-9]+` or `r[0-9]+` | `api123`, `r456` |
| `m[0-9]+` | `m10`, `m87` |

## OUTPUT

**File:** `[baseline]_crossplatform_comparison_v[N].xlsx`

**Sheets:**
- `[version]_comparison` - Per-version comparison
- `cross_version_comparison` - Cross-version (if multiple versions)
- `[backend]_[version]` - Original data backups

## EXAMPLES

```bash
# Single file
python nano_cross_compare.py api123.xlsx

# Multiple files
python nano_cross_compare.py api123.xlsx api456.xlsx

# With version info
python nano_cross_compare.py api123_m10.xlsx api456_m10.xlsx
```

## VERSION DETERMINATION

| Case | Baseline | Compare |
|------|----------|---------|
| Same Skia | Skia version | API versions |
| Same API | API version | Skia versions |
| Single file | Skia version | API version |

## REQUIREMENTS

```bash
pip install pandas openpyxl
```

## VERSION

2.9

