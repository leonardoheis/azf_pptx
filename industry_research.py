from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from utils import (_find_shape_with_token, _add_section_header, _add_bullet, 
                  _fmt_currency, _parse_number, _fmt_billions_usd, _parse_percent)


def fill_industry_research(prs: Presentation, payload: dict, token="{{IndustryResearch}}"):
    """
    Fills the IndustryResearch section with industry-specific data processing and formatting.
    
    Args:
        prs: PowerPoint Presentation object
        payload: Dictionary containing industry research data
        token: Token to find and replace in the presentation
    """
    slide, shape = _find_shape_with_token(prs, token)
    if not shape:
        return

    tf = shape.text_frame
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.clear()

    # Industry-specific data processing
    for key, val in payload.items():
        if key == "Industry Name":
            continue  # Industry name might be handled separately

        if isinstance(val, (dict, list)):
            _add_section_header(tf, f"{key}:")
            
            if isinstance(val, dict):
                # Process industry dictionary data
                for k2, v2 in val.items():
                    if isinstance(v2, list):
                        _add_bullet(tf, f"{k2}:", level=0, size=14)
                        for item in v2:
                            if isinstance(item, dict):
                                # Format industry metrics and data
                                formatted_items = []
                                for kk, vv in item.items():
                                    formatted_value = _format_industry_value(kk, vv)
                                    formatted_items.append(f"{kk}: {formatted_value}")
                                line = "; ".join(formatted_items)
                                _add_bullet(tf, line, level=1, size=12)
                            else:
                                _add_bullet(tf, str(item), level=1, size=12)
                    elif isinstance(v2, dict):
                        _add_bullet(tf, f"{k2}:", level=0, size=14)
                        for kk, vv in v2.items():
                            formatted_value = _format_industry_value(kk, vv)
                            _add_bullet(tf, f"{kk}: {formatted_value}", level=1, size=12)
                    else:
                        formatted_value = _format_industry_value(k2, v2)
                        _add_bullet(tf, f"{k2}: {formatted_value}", level=0, size=14)
            else:
                # Process industry list data
                for item in val:
                    if isinstance(item, dict):
                        formatted_items = []
                        for kk, vv in item.items():
                            formatted_value = _format_industry_value(kk, vv)
                            formatted_items.append(f"{kk}: {formatted_value}")
                        line = "; ".join(formatted_items)
                        _add_bullet(tf, line, level=1, size=12)
                    else:
                        _add_bullet(tf, str(item), level=1, size=12)
        else:
            # Process simple industry values
            formatted_value = _format_industry_value(key, val)
            _add_bullet(tf, f"🏭 {key}: {formatted_value}", level=0, size=14)


def _format_industry_value(key: str, value) -> str:
    """
    Formats industry-specific values with appropriate formatting.
    
    Args:
        key: The key/field name
        value: The value to format
        
    Returns:
        Formatted string value
    """
    key_lower = key.lower()
    
    # Format monetary values
    if any(term in key_lower for term in ['revenue', 'income', 'profit', 'value', 'market size', 'cap']):
        parsed_num = _parse_number(value)
        if parsed_num is not None and parsed_num > 1000000:
            return _fmt_billions_usd(parsed_num)
        elif parsed_num is not None:
            return _fmt_currency(parsed_num)
    
    # Format percentages
    if any(term in key_lower for term in ['percent', 'rate', 'growth', 'margin', 'share']):
        parsed_percent = _parse_percent(value)
        if parsed_percent is not None:
            return f"{parsed_percent:.2f}%"
    
    # Format employee counts
    if any(term in key_lower for term in ['employees', 'workforce', 'staff']):
        parsed_num = _parse_number(value)
        if parsed_num is not None:
            return f"{int(parsed_num):,} employees"
    
    # Default formatting
    return str(value)
