from enum import Enum
from pydantic import Field, SecretStr
from speckle_automate import AutomateBase


class OutputFormat(Enum):
    EXCEL = "Excel"
    GOOGLE_SHEET = "Google Sheet"


class FunctionInputs(AutomateBase):
    """These are function author-defined values."""
    output_format: OutputFormat = Field(
        default=OutputFormat.EXCEL,
        title="Output Format",
        description="Select the output format for the report.",
    )
    google_sheet_id: str = Field(
        default="",
        title="Google Sheet ID",
        description="The ID of the Google Sheet (required if Output Format is Google Sheet).",
    )
    google_service_account_json: SecretStr = Field(
        default=SecretStr(""),
        title="Google Service Account JSON",
        description="The JSON key for the Google Service Account.",
    )