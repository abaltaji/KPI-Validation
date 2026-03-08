# Version Comparison Feature - User Guide

## Overview

The KPI Validation system now includes an advanced **Version Comparison** feature that automatically detects and highlights changes between different versions of your KPI data. This helps you track modifications, identify new programs, and monitor removed items across model revisions.

## Features

### Automatic Change Detection
- **Field-level granularity**: Each KPI parameter is compared individually
- **Three change types**: Added (new programs), Modified (value changes), Removed (deleted programs)
- **Intelligent comparison**: Handles numerical values with proper type conversion

### Visual Highlighting (Excel)
- 🟢 **Green** (#C6EFCE): New programs or fields added in current version
- 🟡 **Yellow** (#FFF2CC): Values that have been modified from previous version
- 🔴 **Red** (#FFC7CE): Programs or fields removed from current version

### Comprehensive Output
Each change record shows:
- **Program**: Which space/program is affected
- **Field**: Which KPI parameter changed
- **Previous Value**: The value in the previous version
- **Current Value**: The value in the current version
- **Status**: Added/Modified/Removed

## Quick Start

### Compare Two Speckle Models
```python
from 03_extraction import extract_capsule_areas
from 04_reporting import generate_excel

# Extract data from current and previous models
current_rows = extract_capsule_areas(current_model)
previous_rows = extract_capsule_areas(previous_model)

# Generate Excel with Version Comparison sheet
output_path = generate_excel(current_rows, previous_rows=previous_rows)
```

### Compare Against Saved Baseline
```python
from 07_version_comparison import load_previous_data_from_file
from 04_reporting import generate_excel
from 03_extraction import extract_capsule_areas

# Load previous export as baseline
previous_rows = load_previous_data_from_file("baseline_2025.xlsx")

# Compare against new data
current_rows = extract_capsule_areas(current_model)
output_path = generate_excel(current_rows, previous_rows=previous_rows)
```

### Google Sheets Comparison
```python
from 04_reporting import update_google_sheet
from 03_extraction import extract_capsule_areas

current_rows = extract_capsule_areas(current_model)
previous_rows = extract_capsule_areas(previous_model)

# Update Google Sheet with Version Comparison tab
url = update_google_sheet(
    current_rows,
    sheet_id="your-sheet-id",
    service_account_json="credentials",
    previous_rows=previous_rows
)
```

## Parameters Compared

The Version Comparison tracks changes in all KPI parameters:

| Parameter | Field Name | Description |
|-----------|-----------|-------------|
| Area | `PRG_PAR_Area` | Total area in m² |
| Use Ratio | `PRG_PAR_UseRatio` | Space utilization ratio |
| Resource Consumption Ratio | `PRG_PAR_ResourceConsRatio` | Resource efficiency metric |
| Geometry Weight | `PRG_PAR_GeometryWeight` | Geometric weighting factor |
| Mean Distance to Exit | `PRG_PAR_MeanDistToExit` | Average evacuation distance (m) |
| Ideal Distance to Exit | `PRG_PAR_IdealDistToExit` | Target evacuation distance (m) |

## Excel Output Structure

When `previous_rows` is provided, your Excel file contains:

### Sheet 1: Program_KPI Parameters
- Current version data with all programs
- KPI values for each program
- Summary total row (bold)

### Sheet 2: Summary
- Total area calculation
- Area per Tower breakdown
- Area per Level breakdown
- Area per Program breakdown
- Tower × Level matrix

### Sheet 3: Version Comparison ⭐ (NEW - only if changes exist)
- Lists all detected changes
- Color-coded by change type
- Sorted by Program and Field name
- Shows Previous and Current values side-by-side

## Example Comparison Output

```
Program      | Field           | Previous Value | Current Value | Status
-------------|-----------------|----------------|---------------|----------
Lobby        | PRG_PAR_Area    | 250.5          | 275.3         | Modified
Cafeteria    | PRG_PAR_UseRatio | 0.75           | 0.75          | Modified
NewOffice    | Program         | N/A            | New Program   | Added
Storage      | Program         | Removed Program| N/A           | Removed
```

(Cells highlighted: Green=Added, Yellow=Modified, Red=Removed)

## Backward Compatibility

✅ **100% backward compatible**
- The `previous_rows` parameter is fully optional
- Existing code works unchanged - simply don't provide previous_rows
- Feature is completely opt-in

## API Reference

### `generate_excel(rows, previous_rows=None) -> str`

Generates Excel report with optional version comparison.

**Parameters:**
- `rows` (list[dict]): Current version data
- `previous_rows` (list[dict], optional): Previous version data for comparison

**Returns:** Path to generated XLSX file

**Example:**
```python
# Without comparison
output = generate_excel(current_rows)

# With comparison
output = generate_excel(current_rows, previous_rows=previous_rows)
```

### `update_google_sheet(rows, sheet_id, service_account_json, previous_rows=None) -> str`

Updates Google Sheet with optional version comparison.

**Parameters:**
- `rows` (list[dict]): Current version data
- `sheet_id` (str): Google Sheet ID
- `service_account_json` (str): Google API credentials
- `previous_rows` (list[dict], optional): Previous version data

**Returns:** URL to Google Sheet

### `_compare_versions(current_rows, previous_rows) -> list[dict]`

Core comparison logic - compares two versions and returns differences.

**Returns:** List of comparison records with fields:
- `program`: Program name
- `field`: Changed field name
- `previous_value`: Old value
- `current_value`: New value
- `status`: "Added", "Modified", or "Removed"
- `is_changed`: Boolean flag

## Helper Functions

### `generate_comparison_report(current_model, previous_model)`

High-level function to generate comparison between two Speckle models.

```python
from 07_version_comparison import generate_comparison_report

output_path = generate_comparison_report(current_model, previous_model)
```

### `load_previous_data_from_file(file_path)`

Load previous version from saved Excel export for comparison.

```python
from 07_version_comparison import load_previous_data_from_file

previous_rows = load_previous_data_from_file("baseline_2025.xlsx")
current_rows = extract_capsule_areas(current_model)
output_path = generate_excel(current_rows, previous_rows=previous_rows)
```

## Troubleshooting

### No Version Comparison sheet appears
- **Check**: Is `previous_rows` parameter provided and non-empty?
- **Check**: Does the previous data contain programs that exist in current data?
- **Note**: Comparison sheet only appears if differences are detected

### All programs showing as "Added"
- **Cause**: Program names may differ (case-sensitive)
- **Solution**: Verify program names match exactly between versions

### Google Sheets formatting not visible
- **Note**: Basic formatting transfers to Google Sheets
- **Recommendation**: Use Excel export for full color highlighting

## Best Practices

1. **Keep baseline exports** saved for regular comparisons
2. **Use consistent program naming** across versions
3. **Export before major changes** to establish a baseline
4. **Review changes regularly** to catch data issues early
5. **Document significant changes** in project notes

## Integration Example

```python
# Complete workflow example
from 03_extraction import extract_capsule_areas
from 04_reporting import generate_excel, update_google_sheet
from 07_version_comparison import load_previous_data_from_file
from 01_inputs import FunctionInputs, OutputFormat

def process_with_comparison(current_model, previous_model, function_inputs):
    """Process model data with version comparison."""
    
    # Extract from both versions
    current_rows = extract_capsule_areas(current_model)
    previous_rows = extract_capsule_areas(previous_model)
    
    # Generate based on output format preference
    if function_inputs.output_format == OutputFormat.GOOGLE_SHEET:
        url = update_google_sheet(
            current_rows,
            function_inputs.google_sheet_id,
            function_inputs.google_service_account_json.get_secret_value(),
            previous_rows=previous_rows  # Pass comparison data
        )
        return f"Updated Google Sheet: {url}"
    else:
        output_path = generate_excel(
            current_rows,
            previous_rows=previous_rows  # Pass comparison data
        )
        return f"Generated Excel: {output_path}"
```

## Future Enhancements

Possible improvements for future versions:
- Tolerance/threshold for "significant" changes
- Percentage change calculations
- Multi-version trend analysis
- Change notifications/alerts
- Timeline view for historical tracking
