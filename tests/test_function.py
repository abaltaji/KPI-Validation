"""Run integration tests with a speckle server."""

import importlib
from speckle_automate import (
    AutomationContext,
    AutomationRunData,
    AutomationStatus,
    run_function,
)
from speckle_automate.fixtures import *  # noqa: F403

inputs_module = importlib.import_module("01_inputs")
FunctionInputs = inputs_module.FunctionInputs

main_module = importlib.import_module("05_main")
automate_function = main_module.automate_function


def test_function_run(
    test_automation_run_data: AutomationRunData, test_automation_token: str
):
    """
    Integration test: Runs the automate function against a real or mocked context.

    Asserts that the function completes with a SUCCEEDED status.
    """
    automation_context = AutomationContext.initialize(
        test_automation_run_data, test_automation_token
    )
    automate_sdk = run_function(
        automation_context,
        automate_function,
        FunctionInputs(),
    )

    assert automate_sdk.run_status == AutomationStatus.SUCCEEDED
