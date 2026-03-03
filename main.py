"""Module for Speckle Automate KPI validation.

Authentication module for Speckle - provides a reusable get_client() function.
"""

import os
from collections import defaultdict
from typing import Any

from dotenv import load_dotenv
# INSTRUCTION: Uncommented openpyxl. You MUST have openpyxl installed 
# in your environment/requirements.txt or the Workbook() call will fail.
from openpyxl import Workbook 
from pydantic import Field, SecretStr
from speckle_automate import (
    AutomateBase,
    AutomationContext,
    execute_automate_function,
)
from specklepy.api.client import SpeckleClient

from flatten import flatten_base


def get_client() -> SpeckleClient:
    """
    Authenticate and return a SpeckleClient instance.
    
    Requires SPECKLE_TOKEN in environment or .env file.
    Optionally set SPECKLE_SERVER (defaults to app.speckle.systems).
    """
    load_dotenv()

    token = os.environ.get("SPECKLE_TOKEN")
    server_host = os.environ.get("SPECKLE_SERVER", "app.speckle.systems")

    if not token:
        raise ValueError("Set SPECKLE_TOKEN in your .env file and re-run.")

    client = SpeckleClient(host=server_host)
    client.authenticate_with_token(token)

    return client


def upload_file_to_speckle(
    client: SpeckleClient, project_id: str, file_path: str, file_name: str
) -> str:
    """Upload a file to Speckle using the REST API."""
    import requests
    
    token = os.environ.get("SPECKLE_TOKEN")
    if not token:
        raise ValueError("SPECKLE_TOKEN not found in environment")
    
    url = f"{client.url}/api/file/create"
    
    with open(file_path, "rb") as f:
        files = {"files": (file_name, f)}
        params = {"streamId": project_id}
        headers = {"Authorization": f"Bearer {token}"}
        
        response = requests.post(url, files=files, params=params, headers=headers)
        response.raise_for_status()
        
        result = response.json()
        file_id = result.get("fileIds", [None])[0]
        
        if not file_id:
            raise ValueError("No file ID returned from upload")
        
        return file_id


def post_comment_with_file(
    client: SpeckleClient, model_id: str, project_id: str, file_id: str, file_name: str
) -> None:
    """Post a comment on the model with the uploaded file."""
    comment_text = f"📊 KPI Validation Report: {file_name}"
    client.comment.create(
        stream_id=project_id,
        object_id=model_id,
        text=comment_text,
        resources=[{"resourceType": "file", "resourceId": file_id}],
    )


class FunctionInputs(AutomateBase):
    """These are function author-defined values."""
    whisper_message: SecretStr = Field(title="This is a secret message")
    forbidden_speckle_type: str = Field(
        title="Forbidden speckle type",
        description=(
            "If a object has the following speckle_type,"
            " it will be marked with an error."
        ),
    )


def _get_attr(obj: Any, *names: str, default=None):
    """Safely get first available attribute/key from object or dict."""
    for name in names:
        if isinstance(obj, dict) and name in obj:
            return obj[name]
        if hasattr(obj, name):
            return getattr(obj, name)
    return default


def extract_capsule_areas(model: Any) -> list[dict]:
    """Extract area rows from a Speckle model."""
    rows: list[dict] = []
    for item in flatten_base(model):
        # INSTRUCTION: For Meshes, ensure they have a custom attribute called "area".
        # Standard Speckle Meshes do not calculate area automatically in Automate.
        area = _get_attr(item, "area", "Area", default=None)
        if area is None:
            continue

        try:
            area_value = float(area)
        except (TypeError, ValueError):
            continue

        rows.append(
            {
                "tower": _get_attr(item, "tower", "Tower", default="Unknown"),
                "level": _get_attr(item, "level", "Level", default="Unspecified"),
                "capsule": _get_attr(
                    item, "capsule", "capsule_no", "CapsuleNo", default=""
                ),
                "program": _get_attr(item, "program", "Program", default="Unspecified"),
                "location": _get_attr(item, "location", "Location", default=""),
                "area": area_value,
            }
        )
    return rows


# INSTRUCTION: Changed the input to accept the direct Base object, not a weird context dict.
def handler(model: Any) -> dict:
    """Main logic function to extract data and build the Excel file."""
    if model is None:
        raise ValueError("No referencedObject found. The received model is empty.")

    rows = extract_capsule_areas(model)
    if not rows:
        return {"message": "No 2D area data found in the model meshes.", "rows": 0}

    workbook = Workbook()

    # =========================
    # SHEET 1: RAW DATA
    # =========================
    raw_sheet = workbook.active
    raw_sheet.title = "Capsule Areas"

    headers = ["Tower", "Level", "Capsule No.", "Program", "Location", "2D Area (m2)"]
    raw_sheet.append(headers)

    for r in rows:
        raw_sheet.append(
            [r["tower"], r["level"], r["capsule"], r["program"], r["location"], r["area"]]
        )

    # =========================
    # DATA AGGREGATION
    # =========================
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

    # Tower x Level matrix
    towers_sorted = sorted(area_per_tower.keys())
    levels_sorted = sorted(matrix.keys())

    summary_sheet.append(["Tower x Level Matrix (m2)"])
    summary_sheet.append(["Level"] + towers_sorted)

    for level in levels_sorted:
        summary_sheet.append([level] + [matrix[level].get(t, 0.0) for t in towers_sorted])

    output_path = "capsule_areas.xlsx"
    workbook.save(output_path)

    return {
        "message": "Excel generated successfully.",
        "output": output_path,
        "rows": len(rows),
    }


def automate_function(
    automation_context: AutomationContext, function_inputs: FunctionInputs
) -> None:
    """Speckle Automate entry point."""
    
    # INSTRUCTION: Safely receive the actual 3D model base object from the Context
    version_root_object = automation_context.receive_version()
    
    # Pass the object to our data handler
    result = handler(version_root_object)
    
    # If no data was found, mark the run as successful but notify the user
    if result.get("rows", 0) == 0:
        automation_context.mark_run_success(result["message"])
        return
    
    # Upload to Speckle
    try:
        client = get_client()
        project_id = automation_context.speckle_client.workspace_id or automation_context.project_id
        model_id = automation_context.model_id
        
        file_path = result["output"]
        file_name = os.path.basename(file_path)
        
        # Upload file
        file_id = upload_file_to_speckle(client, project_id, file_path, file_name)
        
        # Post comment with file attachment
        post_comment_with_file(client, model_id, project_id, file_id, file_name)
        
        # INSTRUCTION: We must explicitly tell Automate that the process finished correctly
        automation_context.mark_run_success(f"✓ File uploaded to Speckle: {file_name}")
        
    except Exception as e:
        # INSTRUCTION: If upload fails, tell Automate the run failed
        automation_context.mark_run_failed(f"⚠ Could not upload to Speckle: {e}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        # Running with Speckle Automate arguments
        execute_automate_function(automate_function, FunctionInputs)
    else:
        # Test authentication and Excel export with real model data
        from specklepy.transports.server import ServerTransport
        from specklepy.api import operations
        
        print("=== Testing Speckle Authentication ===")
        try:
            client = get_client()
            user = client.active_user.get()
            print(f"✓ Logged in as {user.name} on {client.url}")
            
            # Fetch model data and test Excel export
            print("\n=== Testing Excel Export with Model Data ===")
            PROJECT_ID = "08c875bbe4"
            MODEL_ID = "7631638073"
            
            model = client.model.get(MODEL_ID, PROJECT_ID)
            print(f"✓ Model: {model.name}")
            
            versions = client.version.get_versions(MODEL_ID, PROJECT_ID, limit=1)
            latest_version = versions.items[0]
            print(f"  Latest version: {latest_version.id}")
            
            transport = ServerTransport(client=client, stream_id=PROJECT_ID)
            model_data = operations.receive(latest_version.referenced_object, transport)
            
            # INSTRUCTION: Feed the actual object, not a nested dict
            result = handler(model_data)
            
            if result.get("rows", 0) > 0:
                print(f"\n✓ Excel Export Success!")
                print(f"  Output: {result.get('output')}")
                print(f"  Rows processed: {result.get('rows')}")
                
                print(f"\n=== Testing File Upload to Speckle ===")
                file_path = result["output"]
                file_name = os.path.basename(file_path)
                
                file_id = upload_file_to_speckle(client, PROJECT_ID, file_path, file_name)
                print(f"✓ File uploaded!")
                print(f"  File ID: {file_id}")
                
                post_comment_with_file(client, MODEL_ID, PROJECT_ID, file_id, file_name)
                print(f"✓ Comment posted to model with file attachment!")
            else:
                print("\n⚠ Script ran, but no mesh area data was found.")
            
        except ValueError as e:
            print(f"✗ Authentication failed: {e}")
            print("  Make sure SPECKLE_TOKEN is set in your .env file")
        except Exception as e:
            print(f"✗ Error: {e}")
            import traceback
            traceback.print_exc()