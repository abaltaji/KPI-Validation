"""Module for Speckle Automate KPI validation.

Authentication module for Speckle - provides a reusable get_client() function.
"""

import os
import importlib

from dotenv import load_dotenv
from speckle_automate import (
    AutomationContext,
    execute_automate_function,
)
from specklepy.api.client import SpeckleClient

# Import modules using importlib because they start with numbers
inputs_module = importlib.import_module("01_inputs")
FunctionInputs = inputs_module.FunctionInputs
OutputFormat = inputs_module.OutputFormat

extraction_module = importlib.import_module("03_extraction")
extract_capsule_areas = extraction_module.extract_capsule_areas

reporting_module = importlib.import_module("04_reporting")
generate_excel = reporting_module.generate_excel
update_google_sheet = reporting_module.update_google_sheet


def get_client() -> SpeckleClient:
    """
    Create and authenticate a SpeckleClient instance.

    Process:
    1. Loads environment variables (looking for .env file).
    2. Retrieves SPECKLE_TOKEN and SPECKLE_SERVER from the environment.
    3. Initializes the client and authenticates using the token.
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
    client: SpeckleClient, project_id: str, file_path: str, file_name: str, token: str
) -> str:
    """
    Upload a local file to a specific Speckle project via the REST API.

    This function bypasses the standard Python SDK transport to directly post
    a file to the server's blob storage, returning the file ID for later use.
    """
    import requests
    
    # INSTRUCTION: Use the 'token' argument passed to the function.
    # Do NOT use os.environ.get here as it fails in the cloud.
    url = f"{client.url}/api/file/create"
    
    with open(file_path, "rb") as f:
        files = {"files": (file_name, f)}
        params = {"streamId": project_id}
        headers = {"Authorization": f"Bearer {token}"} # Fixed to use the argument
        
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
    """
    Create a comment on a specific model (object) attaching a file.

    Used to notify users in the Speckle Viewer about the generated report.
    """
    comment_text = f"📊 KPI Validation Report: {file_name}"
    client.comment.create(
        stream_id=project_id,
        object_id=model_id,
        text=comment_text,
        resources=[{"resourceType": "file", "resourceId": file_id}],
    )


def automate_function(
    automation_context: AutomationContext, function_inputs: FunctionInputs
):
    """
    Main entry point for the Speckle Automate function.

    Process:
    1. Receives the model version data from the automation context.
    2. Extracts capsule area data using the extraction module.
    3. Generates a report (Excel or Google Sheet) based on user inputs.
    4. Marks the automation run as succeeded or failed based on the outcome.
    """
    
    # 1. Receive the model version data
    version_root_object = automation_context.receive_version()
    
    if not version_root_object:
        automation_context.mark_run_failed("No model data received.")
        return
    
    # 2. Extract Data
    rows = extract_capsule_areas(version_root_object)
    
    if not rows:
        automation_context.mark_run_success("No 2D area data found in the model meshes.")
        return

    # 3. Process based on Output Format
    try:
        if function_inputs.output_format == OutputFormat.GOOGLE_SHEET:
            sheet_url = update_google_sheet(
                rows,
                function_inputs.google_sheet_id,
                function_inputs.google_service_account_json.get_secret_value(),
            )
            automation_context.mark_run_success(f"✓ Google Sheet updated: {sheet_url}")

        else:
            # Default to Excel
            file_path = generate_excel(rows)
            automation_context.store_file_result(file_path)
            automation_context.mark_run_success("✓ Excel Report generated and stored.")
        
    except Exception as e:
        automation_context.mark_run_failed(f"⚠ Processing failed: {e}")


if __name__ == "__main__":
    import sys
    if len(sys.argv) > 1:
        # Running with Speckle Automate arguments
        execute_automate_function(automate_function, FunctionInputs)
    else:
        print("Please run 06_debug.py for local testing.")