"""Module for generating reports (Excel, Google Sheets)."""

import json
import os
import traceback
from json import JSONDecodeError
from collections import defaultdict

from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter
import gspread
from google.oauth2.service_account import Credentials


def generate_excel(rows: list[dict]) -> str:
    """
    Create an Excel report containing raw data and summary statistics.
    
    Process:
    1. Creates a new Excel Workbook.
    2. Sheet 1 ('Capsule Areas'): Writes the raw list of extracted rows.
    3. Sheet 2 ('Summary'): Calculates and writes pivot tables/aggregations (e.g., Area per Tower).
    """
    workbook = Workbook()

    # =========================
    # SHEET 1: RAW DATA
    # =========================
    raw_sheet = workbook.active
    raw_sheet.title = "Program_KPI Parameters"

    headers = [
        "program",
        "PRG_PAR_Area",
        "PRG_PAR_UseRatio",
        "PRG_PAR_GeometryWeight",
        "PRG_PAR_DependenciesDistance",
        "PRG_PAR_IdealDependenciesDistance",
    ]
    raw_sheet.append(headers)

    # Styling Headers (Bold, Blue Background, Centered)
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    
    for cell in raw_sheet[1]:
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center")

    # Aggregate data by program
    aggregated_data = defaultdict(float)
    for r in rows:
        program = r.get("program")
        # Filter: Only include elements that actually have a program value (not "Unspecified")
        if not program or program == "Unspecified":
            continue

        area = r.get("area", 0.0)
        aggregated_data[program] += area

    # Calculate Grand Totals
    grand_total_area = 0.0
    grand_total_use_ratio = 0.0
    grand_total_geometry_weight = 0.0

    for program, total_area in sorted(aggregated_data.items()):
        grand_total_area += total_area
        # Placeholder additions (currently 0.0, but ready for future data)
        grand_total_use_ratio += 0.0
        grand_total_geometry_weight += 0.0

        raw_sheet.append([
            program,        # space_name
            total_area,     # PRG_PAR_Area
            0.0,            # PRG_PAR_UseRatio (Placeholder)
            0.0,            # PRG_PAR_GeometryWeight (Placeholder)
            0.0,            # PRG_PAR_DependenciesDistance (Placeholder)
            0.0             # PRG_PAR_IdealDependenciesDistance (Placeholder)
        ])

    # Append Total Row
    total_row = ["Total", grand_total_area, grand_total_use_ratio, grand_total_geometry_weight, "", ""]
    raw_sheet.append(total_row)

    # Style the Total Row (Bold)
    last_row = raw_sheet.max_row
    for col in range(1, 7):
        raw_sheet.cell(row=last_row, column=col).font = Font(bold=True)

    # Auto-adjust column widths
    for col_idx, column_cells in enumerate(raw_sheet.columns, start=1):
        max_length = 0
        column = get_column_letter(col_idx)
        for cell in column_cells:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        raw_sheet.column_dimensions[column].width = adjusted_width

    # =========================
    # DATA AGGREGATION
    # =========================
    # Calculate totals and group data by categories for the summary sheet
    total_area = sum(r["area"] for r in rows)

    area_per_tower: dict[str, float] = defaultdict(float)
    area_per_level: dict[str, float] = defaultdict(float)
    area_per_program: dict[str, float] = defaultdict(float)
    matrix: dict[str, dict[str, float]] = defaultdict(lambda: defaultdict(float))

    for r in rows:
        tower = r["tower"] or "Unknown"
        level = r["level"] or "Unspecified"
        program = r["program"] or "Unspecified"

        area_per_tower[tower] += r["area"]
        area_per_level[level] += r["area"]
        area_per_program[program] += r["area"]
        matrix[level][tower] += r["area"]

    # =========================
    # SHEET 2: SUMMARY
    # =========================
    summary_sheet = workbook.create_sheet("Summary")
    summary_sheet.append(["Metric", "Value"])
    summary_sheet.append(["Total Area (m2)", total_area])
    summary_sheet.append([])

    summary_sheet.append(["Area per Tower"])
    summary_sheet.append(["Tower", "Area (m2)"])
    for tower, area in sorted(area_per_tower.items()):
        summary_sheet.append([tower, area])
    summary_sheet.append([])

    summary_sheet.append(["Area per Level"])
    summary_sheet.append(["Level", "Area (m2)"])
    for level, area in sorted(area_per_level.items()):
        summary_sheet.append([level, area])
    summary_sheet.append([])

    summary_sheet.append(["Area per Program"])
    summary_sheet.append(["Program", "Area (m2)"])
    for program, area in sorted(area_per_program.items()):
        summary_sheet.append([program, area])
    summary_sheet.append([])

    # Create a Matrix: Levels (rows) x Towers (columns)
    towers_sorted = sorted(area_per_tower.keys())
    levels_sorted = sorted(matrix.keys())

    summary_sheet.append(["Tower x Level Matrix (m2)"])
    summary_sheet.append(["Level"] + towers_sorted)

    for level in levels_sorted:
        summary_sheet.append([level] + [matrix[level].get(t, 0.0) for t in towers_sorted])

    output_path = os.path.join(os.getcwd(), "capsule_areas.xlsx")
    workbook.save(output_path)

    return output_path


def update_google_sheet(
    rows: list[dict], sheet_id: str, service_account_json: str
) -> str:
    """
    Authenticate with Google and update a specific Spreadsheet with the report data.
    
    Process:
    1. Validates inputs (Sheet ID and JSON credentials).
    2. Authenticates using the Service Account credentials.
    3. Connects to the Google Sheet by ID.
    4. Clears and updates the 'Capsule Areas' tab with raw data.
    5. Clears and updates the 'Summary' tab with calculated statistics.
    """
    if not sheet_id:
        raise ValueError("Google Sheet ID is required for this format.")
    if not service_account_json:
        raise ValueError("Service Account JSON is required for this format.")

    try:
        # Clean input: remove whitespace and potential wrapping quotes
        json_str = service_account_json.strip()
        if json_str.startswith("'") and json_str.endswith("'"):
            json_str = json_str[1:-1]

        creds_dict = json.loads(json_str)
        
        if not isinstance(creds_dict, dict):
            raise ValueError("The provided JSON is not a dictionary. Please paste the full content of the JSON file.")
        if "private_key" not in creds_dict or "client_email" not in creds_dict:
            raise ValueError("The JSON key is missing required fields ('private_key' or 'client_email').")
            
        # Define scopes to allow reading/writing to Sheets and Drive
        scopes = [
            "https://www.googleapis.com/auth/spreadsheets",
            "https://www.googleapis.com/auth/drive",
        ]
        # Create credentials object
        creds = Credentials.from_service_account_info(creds_dict, scopes=scopes)
        # Authorize the gspread client
        client = gspread.authorize(creds)
        sh = client.open_by_key(sheet_id)
    except JSONDecodeError as e:
        raise ValueError(
            f"Invalid JSON format in Service Account Key. "
            f"It seems the key is truncated or has a syntax error. "
            f"Error details: {e}"
        )
    except Exception as e:
        print("Detailed error traceback:")
        traceback.print_exc()
        raise ValueError(f"Google Sheets Error ({type(e).__name__}): {e}")

    # 1. Raw Data
    # Select or create the worksheet for raw data
    try:
        ws_raw = sh.worksheet("Program_KPI Parameters")
        ws_raw.clear()
    except gspread.WorksheetNotFound:
        ws_raw = sh.add_worksheet(title="Program_KPI Parameters", rows=1000, cols=20)

    headers = [
        "program",
        "PRG_PAR_Area",
        "PRG_PAR_UseRatio",
        "PRG_PAR_GeometryWeight",
        "PRG_PAR_DependenciesDistance",
        "PRG_PAR_IdealDependenciesDistance",
    ]

    # Aggregate data by program
    aggregated_data = defaultdict(float)
    for r in rows:
        program = r.get("program")
        # Filter: Only include elements that actually have a program value (not "Unspecified")
        if not program or program == "Unspecified":
            continue

        area = r.get("area", 0.0)
        aggregated_data[program] += area

    # Calculate Grand Totals
    grand_total_area = 0.0
    grand_total_use_ratio = 0.0
    grand_total_geometry_weight = 0.0
    sorted_data = sorted(aggregated_data.items())

    for _, total_area in sorted_data:
        grand_total_area += total_area
        # Placeholder additions
        grand_total_use_ratio += 0.0
        grand_total_geometry_weight += 0.0

    data_values = [headers] + [
        [program, total_area, 0.0, 0.0, 0.0, 0.0]
        for program, total_area in sorted_data
    ]
    
    # Append Total Row
    data_values.append(["Total", grand_total_area, grand_total_use_ratio, grand_total_geometry_weight, "", ""])
    
    ws_raw.update(values=data_values)

    # Format Total Row in Google Sheets (Bold)
    total_row_idx = len(data_values)
    # Apply bold formatting to the last row (A:F)
    ws_raw.format(f"A{total_row_idx}:F{total_row_idx}", {"textFormat": {"bold": True}})

    # 2. Summary
    # Calculate aggregates (reusing logic from Excel generation)
    total_area = sum(r["area"] for r in rows)
    
    area_per_tower = defaultdict(float)
    area_per_level = defaultdict(float)
    area_per_program = defaultdict(float)
    matrix = defaultdict(lambda: defaultdict(float))

    for r in rows:
        tower = r["tower"] or "Unknown"
        level = r["level"] or "Unspecified"
        program = r["program"] or "Unspecified"

        area_per_tower[tower] += r["area"]
        area_per_level[level] += r["area"]
        area_per_program[program] += r["area"]
        matrix[level][tower] += r["area"]

    # Select or create the worksheet for summary data
    try:
        ws_summary = sh.worksheet("Summary")
        ws_summary.clear()
    except gspread.WorksheetNotFound:
        ws_summary = sh.add_worksheet(title="Summary", rows=100, cols=10)

    summary_data = [
        ["Metric", "Value"],
        ["Total Area (m2)", total_area],
        [],
        ["Area per Tower"],
        ["Tower", "Area (m2)"],
    ]
    for tower, area in sorted(area_per_tower.items()):
        summary_data.append([tower, area])
    
    summary_data.append([])
    summary_data.append(["Area per Level"])
    summary_data.append(["Level", "Area (m2)"])
    for level, area in sorted(area_per_level.items()):
        summary_data.append([level, area])

    summary_data.append([])
    summary_data.append(["Area per Program"])
    summary_data.append(["Program", "Area (m2)"])
    for program, area in sorted(area_per_program.items()):
        summary_data.append([program, area])

    # Matrix Tower x Level aggregation
    summary_data.append([])
    summary_data.append(["Tower x Level Matrix (m2)"])
    
    towers_sorted = sorted(area_per_tower.keys())
    levels_sorted = sorted(matrix.keys())
    
    # Header row for matrix
    summary_data.append(["Level"] + towers_sorted)
    
    # Data rows for matrix
    for level in levels_sorted:
        row = [level] + [matrix[level].get(t, 0.0) for t in towers_sorted]
        summary_data.append(row)

    # Use named argument 'values' for compatibility with newer gspread versions
    ws_summary.update(values=summary_data)

    return sh.url