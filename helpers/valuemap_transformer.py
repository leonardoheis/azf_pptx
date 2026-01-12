"""
ValueMap to IndustryResearch transformer module.

Transforms ValueMap BenefitTable format into the IndustryResearch table format
expected by the PowerPoint generation logic.
"""

import logging
from typing import Any

# Keys to exclude from the flattened row output (none - include all keys)
EXCLUDED_KEYS: set[str] = set()

# Default headers order (preferred column ordering)
PREFERRED_HEADER_ORDER = [
    "Challenge",
    "Description",
    "ScenarioRecordID",
    "KPI",
    "Workload",
    "BenefitFormula",
    "Inputs",
    "CalculatedBenefit",
    "CalculatedBenefitUSD",
    "BenefitCurrency",
    "Notes",
]


def is_valuemap_format(data) -> bool:
    """
    Check if the data is in ValueMap format.

    Supports multiple formats:
    - Direct array of objects: [{...}, {...}]
    - Dict with BenefitTable key: {"BenefitTable": [...]}
    - Dict with any key containing a list of objects

    Args:
        data: Data to check (dict or list)

    Returns:
        True if data contains valid table data
    """
    # Direct array of objects
    if isinstance(data, list) and data and isinstance(data[0], dict):
        return True
    # Dict with BenefitTable or any list of objects
    if isinstance(data, dict):
        if "BenefitTable" in data and isinstance(data["BenefitTable"], list):
            return True
        # Check for any key containing a list of objects
        for v in data.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return True
    return False


def _extract_benefit_table(data) -> list[dict] | None:
    """
    Extract the benefit table data from various input formats.

    Supports:
    - Direct array: [{...}, {...}]
    - Dict with BenefitTable: {"BenefitTable": [...]}
    - Dict with any list of objects (fallback)

    Args:
        data: Input data (dict or list)

    Returns:
        List of row dictionaries, or None if no valid data found
    """
    # Direct array of objects
    if isinstance(data, list):
        if data and isinstance(data[0], dict):
            return data
        return None

    # Dict with BenefitTable key (preferred)
    if isinstance(data, dict):
        if "BenefitTable" in data:
            bt = data["BenefitTable"]
            if isinstance(bt, list) and bt:
                return bt
        # Fallback: find first list of objects
        for v in data.values():
            if isinstance(v, list) and v and isinstance(v[0], dict):
                return v

    return None


def transform_valuemap_to_industry_research(
    valuemap,
    company_name: str = "",
) -> dict:
    """
    Transform ValueMap into IndustryResearch format.

    Supports multiple input formats including direct arrays and wrapped objects.

    Args:
        valuemap: Data containing benefit table (dict or list)
        company_name: Optional company name for title generation

    Returns:
        Dictionary with title, headers, and rows keys.
        Returns minimal valid structure if no data found.
    """
    benefit_table = _extract_benefit_table(valuemap)

    # Return minimal valid structure if no data found
    if not benefit_table:
        logging.warning("No valid benefit table data found in ValueMap")
        return {
            "title": _generate_title(company_name),
            "headers": [],
            "rows": [],
        }

    # Extract headers from first row, excluding complex nested objects
    headers = _extract_headers(benefit_table[0])

    # Transform rows, keeping only the headers we extracted
    rows = _transform_rows(benefit_table, headers)

    # Generate title
    title = _generate_title(company_name)

    result = {
        "title": title,
        "headers": headers,
        "rows": rows,
    }

    logging.info(
        "Transformed ValueMap to IndustryResearch: %d headers, %d rows",
        len(headers),
        len(rows),
    )

    return result


def _extract_headers(first_row: dict) -> list[str]:
    """
    Extract headers from the first row of BenefitTable.

    Excludes complex nested objects (like Inputs) and orders headers
    according to preferred order when possible.

    Args:
        first_row: First row of the BenefitTable

    Returns:
        List of header strings in preferred order
    """
    if not isinstance(first_row, dict):
        raise ValueError("BenefitTable rows must be objects")

    # Get all keys excluding complex nested objects
    raw_headers = [
        key for key in first_row.keys() if key not in EXCLUDED_KEYS and not _is_complex_value(first_row[key])
    ]

    # Sort headers: preferred order first, then remaining alphabetically
    ordered_headers = []
    remaining_headers = set(raw_headers)

    for preferred in PREFERRED_HEADER_ORDER:
        if preferred in remaining_headers:
            ordered_headers.append(preferred)
            remaining_headers.remove(preferred)

    # Add remaining headers alphabetically
    ordered_headers.extend(sorted(remaining_headers))

    return ordered_headers


def _is_complex_value(value: Any) -> bool:
    """
    Check if a value is a complex nested structure.

    Note: Complex values like dicts are now supported via flattening.

    Args:
        value: Value to check

    Returns:
        True only for list of dicts (too complex to flatten meaningfully)
    """
    # Dicts can be flattened, so they're not excluded anymore
    if isinstance(value, list) and value and isinstance(value[0], dict):
        return True
    return False


def _flatten_value(value: Any) -> str:
    """
    Convert a value to a string suitable for table display.

    Handles nested dicts by flattening to "key: value; key: value" format.

    Args:
        value: Value to convert

    Returns:
        String representation of the value
    """
    if value is None:
        return ""
    if isinstance(value, dict):
        # Flatten dict to readable string
        parts = [f"{k}: {v}" for k, v in value.items()]
        return "; ".join(parts)
    if isinstance(value, list):
        # Simple lists to comma-separated string
        return ", ".join(str(v) for v in value)
    return str(value)


def _transform_rows(benefit_table: list[dict], headers: list[str]) -> list[dict]:
    """
    Transform BenefitTable rows to include only the specified headers.

    Flattens complex values (like nested dicts) for display.

    Args:
        benefit_table: List of row dictionaries
        headers: List of header keys to include

    Returns:
        List of row dictionaries with only the specified keys
    """
    rows = []
    for row in benefit_table:
        if not isinstance(row, dict):
            logging.warning("Skipping non-dict row in BenefitTable")
            continue

        transformed_row = {}
        for header in headers:
            value = row.get(header, "")
            # Flatten complex values to string representation
            transformed_row[header] = _flatten_value(value)

        rows.append(transformed_row)

    return rows


def _generate_title(company_name: str) -> str:
    """
    Generate a title for the IndustryResearch table.

    Args:
        company_name: Company name to include in title

    Returns:
        Generated title string
    """
    if company_name and company_name.strip():
        return f"{company_name.strip()} Value Map Benefit Table"
    return "Value Map Benefit Table"
