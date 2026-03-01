"""
02 - Fetch Information from a Speckle Model

This script demonstrates how to fetch and explore data from
an existing Speckle model/version.

Use your personal tower model from session2 (project "cw-yourname")
https://app.speckle.systems/projects/YOUR_PROJECT_ID/models/YOUR_MODEL_ID
"""

from main import get_client
from specklepy.transports.server import ServerTransport
from specklepy.api import operations


# TODO: Replace with your project and model IDs
PROJECT_ID = "08c875bbe4"
MODEL_ID = "7631638073"


def main():
    # Authenticate
    client = get_client()

    # Get the model (branch) info
    model = client.model.get(MODEL_ID, PROJECT_ID)
    print(f"✓ Model: {model.name}")

    # Get the latest version (commit)
    versions = client.version.get_versions(MODEL_ID, PROJECT_ID, limit=1)
    latest_version = versions.items[0]
    print(f"  Latest version: {latest_version.id}")
    print(f"  Message: {latest_version.message}")

    # Receive the data from Speckle
    transport = ServerTransport(client=client, stream_id=PROJECT_ID)
    data = operations.receive(latest_version.referenced_object, transport)

    # Print element names
    elements = getattr(data, "elements", [])
    for element in elements:
        print(element.name)


if __name__ == "__main__":
    main()
