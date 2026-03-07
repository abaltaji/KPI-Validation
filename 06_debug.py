"""
06 - Debug / Fetch Information from a Speckle Model

This script demonstrates how to fetch and explore data from
an existing Speckle model/version locally.
"""

import importlib
import os
from specklepy.transports.server import ServerTransport
from specklepy.api import operations

# Import from 05_main.py
main_module = importlib.import_module("05_main")
get_client = main_module.get_client

# Import extraction and reporting for local testing
extraction_module = importlib.import_module("03_extraction")
extract_capsule_areas = extraction_module.extract_capsule_areas

reporting_module = importlib.import_module("04_reporting")
generate_excel = reporting_module.generate_excel
update_google_sheet = reporting_module.update_google_sheet

# Import version comparison module
version_comparison_module = importlib.import_module("07_version_comparison")
load_previous_data_from_file = version_comparison_module.load_previous_data_from_file


# TODO: Replace with your project and model IDs
PROJECT_ID = "08c875bbe4"
MODEL_ID = "7631638073"


def main():
    """
    Main execution flow for fetching and exploring model data.

    Process:
    1. Authenticates with the Speckle server using local environment variables.
    2. Fetches the specified model and its latest version.
    3. Receives the model data (objects) from the server.
    4. Extracts capsule area data using the extraction module.
    5. Generates an Excel report locally.
    6. Optionally updates a Google Sheet if credentials are present.
    """
    print("=== Testing Speckle Authentication ===")
    try:
        client = get_client()
        active_user = client.active_user.get()
        print(f"✓ Logged in as {active_user.name} on {client.url}")
        
        # Fetch model data and test Excel export
        print("\n=== Testing Excel Export with Model Data ===")
        
        model_info = client.model.get(MODEL_ID, PROJECT_ID)
        print(f"✓ Model: {model_info.name}")
        
        versions = client.version.get_versions(MODEL_ID, PROJECT_ID, limit=1)
        latest_version = versions.items[0]
        print(f"  Latest version: {latest_version.id}")
        
        transport = ServerTransport(client=client, stream_id=PROJECT_ID)
        model_data = operations.receive(latest_version.referenced_object, transport)
        
        # Test Extraction
        rows = extract_capsule_areas(model_data)
        
        if rows:
            print(f"\n✓ Found {len(rows)} rows.")
            # Test Excel Generation locally
            output = generate_excel(rows)
            print(f"  Excel generated: {output}")

            # --- TEST VERSION COMPARISON ---
            print("\n=== Testing Version Comparison ===")
            try:
                # Load the baseline we just created
                previous_rows = load_previous_data_from_file(output)
                print(f"✓ Loaded baseline data: {len(previous_rows)} rows")
                
                # Simulate a change by modifying one value
                if rows and len(rows) > 0:
                    # Modify the first row's area to simulate a change
                    original_area = rows[0].get('PRG_PAR_Area', 0)
                    rows[0]['PRG_PAR_Area'] = float(original_area) + 10 if original_area else 10
                    print(f"✓ Modified first program area (+10 m²) for comparison")
                
                # Generate report with comparison
                comparison_output = generate_excel(rows, previous_rows=previous_rows)
                print(f"✓ Version Comparison report generated: {comparison_output}")
                print("  Check the 'Version Comparison' sheet for highlighted changes")
                print("  (Yellow=Modified, Green=Added, Red=Removed)")
                
            except Exception as e:
                print(f"⚠ Version comparison test failed: {e}")
                import traceback
                traceback.print_exc()

            # --- TEST GOOGLE SHEETS (Uncomment to test) ---
            print("\n=== Testing Google Sheets Export ===")
            SHEET_ID = "1JOzT3Kg3g9nDhjDb9Pp_-ZmR54kI5BxSrdqB1vzWaGQ"
            
            # Load credentials from secure file (ignored in .gitignore)
            json_path = "service_account.json"
            if os.path.exists(json_path):
                with open(json_path, "r", encoding="utf-8") as f:
                    json_key = f.read()
                url = update_google_sheet(rows, SHEET_ID, json_key)
                print(f"✓ Google Sheet updated: {url}")
            else:
                print(f"⚠ To test Sheets, create '{json_path}' with your credentials.")
        else:
            print("\n⚠ Script ran, but no mesh area data was found.")
            
    except Exception as e:
        print(f"✗ Error: {e}")
        import traceback
        traceback.print_exc()


if __name__ == "__main__":
    main()