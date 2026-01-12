from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE

from helpers.exceptions import TemplateError
from helpers.utils import _add_bullet_runs, _add_section_header, _find_shape_with_token

# Font sizes for hierarchical layout
HEADER_FONT_SIZE = 12
FIELD_FONT_SIZE = 10


def fill_company_research2(prs: Presentation, payload: dict, company_name: str | None = None):
    """
    Fills CompanyResearch2 slide with hierarchical bullet points from the payload.

    Outputs section headers with sub-bullets for each field:
    Revenue:
      • Amount: $X
      • Fiscal Year: YYYY
      • Source: link
    """
    token = "{{CompanyResearch2}}"

    slide, shape = _find_shape_with_token(prs, token)
    if not shape:
        raise TemplateError(f"Token '{token}' not found in any slide")

    tf = shape.text_frame
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.clear()

    # --- Revenue ---
    revenue_data = payload.get("Revenue", {})
    if revenue_data:
        _add_section_with_fields(tf, "Revenue:", revenue_data)

    # --- Industry Average Gross Margin ---
    industry_gm_data = payload.get("Industry Average Gross Margin", {})
    if industry_gm_data:
        _add_section_with_fields(tf, "Industry Average Gross Margin:", industry_gm_data)

    # --- Company Gross Margin ---
    company_gm_data = payload.get("Company Gross Margin", {})
    if company_gm_data:
        _add_section_with_fields(tf, "Company Gross Margin:", company_gm_data)

    # --- Employee Count ---
    employee_data = payload.get("Employee Count", {})
    if employee_data:
        _add_section_with_fields(tf, "Employee Count:", employee_data)


def _add_section_with_fields(tf, header: str, data: dict) -> None:
    """
    Add a section header followed by sub-bullets for each field.

    Args:
        tf: Text frame to add content to
        header: Section header text (e.g., "Revenue:")
        data: Dictionary of field key-value pairs
    """
    # Add section header (bold, level 0)
    _add_section_header(tf, header, size=HEADER_FONT_SIZE)

    # Add sub-bullets for each field
    for key, value in data.items():
        _add_field_bullet(tf, key, value)


def _add_field_bullet(tf, key: str, value) -> None:
    """
    Add a sub-bullet for a single field.

    Args:
        tf: Text frame to add content to
        key: Field name
        value: Field value (can be string, number, etc.)
    """
    # Format the value appropriately
    formatted_value = _format_field_value(key, value)

    # Check if value is a URL (for Source fields)
    is_url = isinstance(value, str) and value.startswith(("http://", "https://"))

    if is_url:
        # Make source links clickable
        runs = [
            {"text": f"{key}: ", "link": None},
            {"text": formatted_value, "link": value},
        ]
        _add_bullet_runs(tf, runs, level=1, size=FIELD_FONT_SIZE)
    else:
        # Regular field
        runs = [{"text": f"{key}: {formatted_value}", "link": None}]
        _add_bullet_runs(tf, runs, level=1, size=FIELD_FONT_SIZE)


def _format_field_value(key: str, value) -> str:
    """
    Format a field value for display.

    Args:
        key: Field name (used to determine formatting)
        value: Raw value

    Returns:
        Formatted string representation
    """
    if value is None:
        return "Not available"

    # Handle numeric amounts (for Amount/Headcount fields)
    if key.lower() == "amount" and isinstance(value, (int, float)):
        return _format_currency(value)

    if key.lower() == "headcount" and isinstance(value, (int, float)):
        return f"{int(value):,}"

    # Handle percentage values
    if isinstance(value, (int, float)) and "margin" in key.lower():
        return f"{value}%"

    # Default: convert to string
    return str(value)


def _format_currency(amount) -> str:
    """
    Format currency amount with dollar sign and commas.

    Args:
        amount: Numeric amount

    Returns:
        Formatted string like "$421,300,000"
    """
    if amount is None:
        return "Not disclosed"

    try:
        amount = float(amount)
    except (TypeError, ValueError):
        return str(amount)

    return f"${int(amount):,}"
