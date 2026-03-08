# Data Validation Feature

## Overview

The Data Validation sheet provides automated quality checks for extracted KPI data, ensuring that all properties meet specific integrity requirements. This sheet is automatically generated as the 4th tab in every Excel report.

## Validation Categories

### 1. Data Completeness
- **Programs with data**: Verifies that at least one program with valid data exists
- **Programs with area data**: Checks that all programs have area values greater than 0

### 2. Numerical Thresholds
- **Area > 0 m²**: All programs must have positive area values
- **Use Ratio (0-1)**: Use ratios must be between 0 and 1 (inclusive)
- **Resource Ratio (0-1)**: Resource consumption ratios must be between 0 and 1
- **Geometry Weight >= 0**: Geometry weights must be non-negative
- **Mean Distance to Exit >= 0**: Distance values must be non-negative
- **Ideal Distance to Exit >= 0**: Distance values must be non-negative

### 3. Data Consistency
- **All KPI parameters present**: Verifies that all required KPI parameters are provided for each program

## Validation Results Format

Each validation check shows:

| Column | Purpose |
|--------|---------|
| Category | Type of validation (Data Completeness, Numerical Thresholds, etc.) |
| Check | Description of what is being validated |
| Status | Pass (green) or Fail (red) |
| Details | Specific information about the check result |

## Example Results

```
Category: Data Completeness
Check: Programs with data
Status: Pass
Details: Found 4 programs with valid data

Category: Numerical Thresholds
Check: Area > 0 m²
Status: Fail
Details: Invalid: Office B
```

## Color Coding

- **Green background**: Validation check PASSED
- **Red background**: Validation check FAILED
- **Blue headers**: Column headers and legend title

## Legend

The Data Validation sheet includes a legend in columns G-H that explains the color coding:

- **Pass** (green): Validation check passed - data meets requirements
- **Fail** (red): Validation check failed - data needs review and correction

The legend is positioned to the right of the validation results table for easy reference.

## How It Works

### Automatic Execution

The data validation sheet is created automatically whenever you:
1. Call `generate_excel(rows)` directly
2. Run the debug script (`06_debug.py`)
3. Use the version comparison feature

### Manual Validation

You can also test validation on any dataset:

```python
from reporting_module import _validate_data

results = _validate_data(rows)
for result in results:
    print(f"{result['check']}: {result['status']}")
    print(f"  {result['details']}")
```

## Common Issues and Solutions

### Issue: "Area > 0 m²" shows FAIL
**Cause**: One or more programs have area value of 0 or negative
**Solution**: Check the Program_KPI Parameters sheet and verify area values

### Issue: "Use Ratio (0-1)" shows FAIL
**Cause**: A program has a use ratio outside the 0-1 range
**Solution**: Verify the ratio values are between 0 and 1

### Issue: "All KPI parameters present" shows FAIL
**Cause**: Some programs are missing one or more KPI parameters
**Solution**: Ensure all programs have complete data for:
- PRG_PAR_Area
- PRG_PAR_UseRatio
- PRG_PAR_ResourceConsRatio
- PRG_PAR_GeometryWeight
- PRG_PAR_MeanDistToExit
- PRG_PAR_IdealDistToExit

## Integration with Version Comparison

When using the version comparison feature with `generate_excel(rows, previous_rows=previous_rows)`:

1. **Program_KPI Parameters** sheet: Current data with light blue program separators
2. **Summary** sheet: Aggregated statistics
3. **Version Comparison** sheet: All changes between versions (if previous version provided)
4. **Data Validation** sheet: Quality checks on current data

## Customizing Validation Rules

To modify or add validation rules, edit the `_validate_data()` function in `04_reporting.py`:

```python
def _validate_data(rows: list[dict]) -> list[dict]:
    """Add or modify validation checks here"""
    validation_results = []
    
    # Example: Add custom minimum area threshold
    min_area = 100  # Minimum area in m²
    invalid = [r["program"] for r in rows if r.get("area", 0) < min_area]
    
    validation_results.append({
        "category": "Custom Rules",
        "check": f"Minimum area {min_area} m²",
        "status": "Pass" if len(invalid) == 0 else "Fail",
        "details": f"All areas >= {min_area}" if len(invalid) == 0 else f"Invalid: {', '.join(invalid[:3])}"
    })
    
    return validation_results
```

## Performance Considerations

The validation process is highly optimized:
- **Data Completeness checks**: O(n) where n = number of programs
- **Threshold checks**: O(n*m) where m = number of validation rules (typically 8)
- **Sheet creation**: O(n) for formatting

For typical datasets (100-10,000 programs), validation completes in milliseconds.

## See Also

- [VERSION_COMPARISON_GUIDE.md](VERSION_COMPARISON_GUIDE.md) - Version comparison feature
- [04_reporting.py](04_reporting.py) - Main reporting module
- [test_data_validation.py](test_data_validation.py) - Example validation test
