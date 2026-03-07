"""Module for extracting data from Speckle models."""

import importlib
from typing import Any

# Import flatten_base from 02_helpers.py using importlib because of the numeric prefix
helpers = importlib.import_module("02_helpers")
flatten_base = helpers.flatten_base


def _get_attr(obj: Any, *names: str, default=None):
    """
    Safely retrieve the first available attribute or key from an object or dictionary.
    
    This helper function checks a list of potential attribute names (e.g., 'area', 'Area')
    against the provided object. It handles both dictionary-style access (obj['key'])
    and object-style access (obj.attr).
    
    Returns the value if found, otherwise returns the default value.
    """
    for name in names:
        if isinstance(obj, dict) and name in obj:
            return obj[name]
        if hasattr(obj, name):
            return getattr(obj, name)
    return default


def extract_capsule_areas(model: Any) -> list[dict]:
    """
    Traverse the Speckle model to extract capsule data and calculate areas.
    
    Process:
    1. Flattens the hierarchical Speckle model into a list of objects.
    2. Iterates through each object looking for an 'area' attribute.
    3. Validates that the area is a number.
    4. Extracts relevant metadata (Tower, Level, Capsule No, Program, Location).
    5. Returns a list of dictionaries, where each dictionary represents a row of data.
    """
    rows: list[dict] = []
    for item in flatten_base(model):
        # Extract properties dictionary
        props = _get_attr(item, "properties", "Properties")

        # INSTRUCTION: Look for the specific property 'PRG_PAR_Area' inside 'properties'
        area = None
        if props:
            area = _get_attr(props, "PRG_PAR_Area")

        if area is None:
            continue

        try:
            area_value = float(area)
        except (TypeError, ValueError):
            continue

        # Extract program: specifically look in 'properties' as requested
        program = None
        if props:
            program = _get_attr(props, "program", "Program")
        
        # Fallback to top level if not found in properties
        if not program:
            program = _get_attr(item, "program", "Program", default="Unspecified")

        # Extract additional KPI parameters
        use_ratio = 0.0
        resource_cons_ratio = 0.0
        geometry_weight = 0.0
        mean_dist_to_exit = 0.0
        ideal_dist_to_exit = 0.0

        if props:
            try:
                use_ratio = float(_get_attr(props, "PRG_PAR_UseRatio", default=0.0) or 0.0)
                resource_cons_ratio = float(_get_attr(props, "PRG_PAR_ResourceConsRatio", default=0.0) or 0.0)
                geometry_weight = float(_get_attr(props, "PRG_PAR_GeometryWeight", default=0.0) or 0.0)
                mean_dist_to_exit = float(_get_attr(props, "PRG_PAR_MeanDistToExit", default=0.0) or 0.0)
                ideal_dist_to_exit = float(_get_attr(props, "PRG_PAR_IdealDistToExit", default=0.0) or 0.0)
            except (TypeError, ValueError):
                pass

        # Build a dictionary for the current item with all required fields
        rows.append(
            {
                "tower": _get_attr(item, "tower", "Tower", default="Unknown"),
                "level": _get_attr(item, "level", "Level", default="Unspecified"),
                "capsule": _get_attr(
                    item, "capsule", "capsule_no", "CapsuleNo", default=""
                ),
                "program": program,
                "location": _get_attr(item, "location", "Location", default=""),
                "area": area_value,
                "use_ratio": use_ratio,
                "resource_cons_ratio": resource_cons_ratio,
                "geometry_weight": geometry_weight,
                "mean_dist_to_exit": mean_dist_to_exit,
                "ideal_dist_to_exit": ideal_dist_to_exit,
            }
        )
    return rows