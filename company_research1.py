from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from utils import _find_shape_with_token, _add_section_header, _add_bullet


def fill_company_research1(prs: Presentation, payload: dict, token="{{CompanyResearch1}}"):
    """
    Fills the CompanyResearch1 section with bullet points from objects/lists.
    
    Args:
        prs: PowerPoint Presentation object
        payload: Dictionary containing company research data
        token: Token to find and replace in the presentation
    """
    slide, shape = _find_shape_with_token(prs, token)
    if not shape:
        return

    tf = shape.text_frame
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.clear()

    # Ignore "Company Name" because the name goes through {{CompanyName}}
    for key, val in payload.items():
        if key == "Company Name":
            continue

        # Section header if the value is complex
        if isinstance(val, (dict, list)):
            _add_section_header(tf, f"{key}:")
            if isinstance(val, dict):
                # dict: sub-keys as bullets
                for k2, v2 in val.items():
                    if isinstance(v2, list):
                        _add_bullet(tf, f"{k2}:", level=0, size=14)
                        # list of dicts or primitives
                        for item in v2:
                            if isinstance(item, dict):
                                line = "; ".join(f"{kk}: {vv}" for kk, vv in item.items())
                                _add_bullet(tf, line, level=1, size=12)
                            else:
                                _add_bullet(tf, str(item), level=1, size=12)
                    elif isinstance(v2, dict):
                        # one more level
                        _add_bullet(tf, f"{k2}:", level=0, size=14)
                        for kk, vv in v2.items():
                            _add_bullet(tf, f"{kk}: {vv}", level=1, size=12)
                    else:
                        _add_bullet(tf, f"{k2}: {v2}", level=0, size=14)
            else:
                # list at root level
                for item in val:
                    if isinstance(item, dict):
                        line = "; ".join(f"{kk}: {vv}" for kk, vv in item.items())
                        _add_bullet(tf, line, level=1, size=12)
                    else:
                        _add_bullet(tf, str(item), level=1, size=12)
        else:
            # simple value
            _add_bullet(tf, f"{key}: {val}", level=0, size=14)
