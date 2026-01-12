import logging

from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE

from helpers.exceptions import TemplateError
from helpers.utils import (
    _add_bullet_runs,
    _add_section_header,
    _find_shape_with_token,
    _load_json,
    _replace_company_name_everywhere,
)

# Font sizes matching company_research2.py
HEADER_FONT_SIZE = 10
FIELD_FONT_SIZE = 8


def fill_company_research1(prs: Presentation, payload: dict):
    """
    Fills the CompanyResearch1 section with hierarchical bullet points.

    Outputs section headers with sub-bullets for each field:
    Profile:
      • Description: ...
      • Industry: ...
      • Core Mission: ...
    """
    token = "{{CompanyResearch1}}"

    slide, shape = _find_shape_with_token(prs, token)
    if not shape:
        raise TemplateError(f"Token '{token}' not found in any slide")

    tf = shape.text_frame
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.clear()

    if not payload:
        logging.warning("CompanyResearch1 payload is empty; slide will be left blank.")
        return

    # Process each top-level key (except Company Name)
    for key, val in payload.items():
        if key == "Company Name":
            continue

        if isinstance(val, dict):
            # Section with nested fields
            _add_section_with_nested_fields(tf, f"{key}:", val)
        elif isinstance(val, list):
            # Section with list of items
            _add_section_with_list(tf, f"{key}:", val)
        else:
            # Simple key-value at top level
            _add_field_bullet(tf, key, val, level=0)


def _add_section_with_nested_fields(tf, header: str, data: dict, level: int = 0) -> None:
    """
    Add a section header followed by sub-bullets for each field.

    Args:
        tf: Text frame to add content to
        header: Section header text
        data: Dictionary of field key-value pairs
        level: Indentation level for the header
    """
    if level == 0:
        _add_section_header(tf, header, size=HEADER_FONT_SIZE)
    else:
        _add_simple_bullet(tf, header, level=level, size=HEADER_FONT_SIZE)

    # Add sub-bullets for each field
    for key, value in data.items():
        if isinstance(value, dict):
            # Nested dict - recurse with increased level
            _add_section_with_nested_fields(tf, f"{key}:", value, level=level + 1)
        elif isinstance(value, list):
            # List of items
            _add_list_field(tf, key, value, level=level + 1)
        else:
            # Simple field
            _add_field_bullet(tf, key, value, level=level + 1)


def _add_section_with_list(tf, header: str, items: list) -> None:
    """
    Add a section header followed by list items.

    Args:
        tf: Text frame to add content to
        header: Section header text
        items: List of items (can be dicts or primitives)
    """
    _add_section_header(tf, header, size=HEADER_FONT_SIZE)

    for item in items:
        if isinstance(item, dict):
            # Each dict item gets its fields as sub-bullets
            for key, value in item.items():
                _add_field_bullet(tf, key, value, level=1)
        else:
            # Primitive item
            _add_simple_bullet(tf, str(item), level=1, size=FIELD_FONT_SIZE)


def _add_list_field(tf, key: str, items: list, level: int) -> None:
    """
    Add a field that contains a list of items.

    Args:
        tf: Text frame to add content to
        key: Field name
        items: List of items
        level: Indentation level
    """
    _add_simple_bullet(tf, f"{key}:", level=level, size=HEADER_FONT_SIZE)

    for item in items:
        if isinstance(item, dict):
            # Each dict item gets its fields as sub-bullets
            for k, v in item.items():
                _add_field_bullet(tf, k, v, level=level + 1)
        else:
            _add_simple_bullet(tf, str(item), level=level + 1, size=FIELD_FONT_SIZE)


def _add_field_bullet(tf, key: str, value, level: int) -> None:
    """
    Add a sub-bullet for a single field with optional hyperlink for URLs.

    Args:
        tf: Text frame to add content to
        key: Field name
        value: Field value
        level: Indentation level
    """
    value_str = str(value) if value is not None else ""

    # Check if value is a URL
    is_url = isinstance(value, str) and value.startswith(("http://", "https://"))

    if is_url:
        runs = [
            {"text": f"• {key}: ", "link": None},
            {"text": value_str, "link": value},
        ]
        _add_bullet_runs(tf, runs, level=level, size=FIELD_FONT_SIZE)
    else:
        runs = [{"text": f"• {key}: {value_str}", "link": None}]
        _add_bullet_runs(tf, runs, level=level, size=FIELD_FONT_SIZE)


def _add_simple_bullet(tf, text: str, level: int, size: int) -> None:
    """Add a simple bullet point without key-value formatting."""
    runs = [{"text": f"• {text}", "link": None}]
    _add_bullet_runs(tf, runs, level=level, size=size)


# --------------------------------------------------------------------
# CompanyName desde JSON externo
# --------------------------------------------------------------------
def _get_company_name_from_json(path_or_obj) -> str:
    """Accept either a loaded dict or a path to a JSON file.

    Use the helper `_load_json` which handles both cases.
    """
    try:
        data = _load_json(path_or_obj)
    except Exception as exc:
        raise ValueError(f"Failed to load company JSON: {exc}") from exc

    # Accept either unwrapped dict or {"data": [ {...} ]}
    if isinstance(data, dict) and "Company Name" in data:
        return str(data.get("Company Name", "")).strip()

    try:
        return str(data["data"][0].get("Company Name", "")).strip()
    except Exception as exc:
        raise ValueError("Company JSON missing expected 'Company Name' field") from exc


def fill_company_name_from_json(prs: Presentation, company_json_path: str):
    # Let any error propagate so callers can handle/log it appropriately
    name = _get_company_name_from_json(company_json_path)
    _replace_company_name_everywhere(prs, name)
