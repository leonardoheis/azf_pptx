from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE

from helpers.exceptions import TemplateError
from helpers.utils import (
    _add_bullet_runs,
    _choose_link,
    _deep_find,
    _find_shape_with_token,
    _fmt_billions_usd,
    _get_first_str,
    _norm,
    _parse_date,
    _parse_number,
    _parse_percent,
)


# --------------------------------------------------------------------
# Función 2: {{CompanyResearch2}}  (tabla de métricas clave)
# --------------------------------------------------------------------
def fill_company_research2(prs: Presentation, payload: dict, company_name: str | None = None):
    """
    Genera bullets estilo narrativa, tolerante a cambios de claves/estructura.
    Detecta Revenue, Industry Avg Gross Margin, Company Gross Margin, Employee Count.
    Si falta algún dato, lo omite o degrada el texto.
    """
    token = "{{CompanyResearch2}}"

    slide, shape = _find_shape_with_token(prs, token)
    if not shape:
        raise TemplateError(f"Token '{token}' not found in any slide")

    # Inferencia de nombre si no llega por parámetro
    if not company_name:
        # prueba con claves típicas
        for k in ("Company Name", "Name", "Company"):
            if k in payload and isinstance(payload[k], str) and payload[k].strip():
                company_name = payload[k].strip()
                break
        company_name = company_name or "The company"

    tf = shape.text_frame
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.clear()

    # --- Revenue (busca por revenue/sales/total revenue) ---
    rev_obj = _deep_find(payload, ["revenue", "sales", "total revenue", "latest revenue", "annual revenue"])
    if isinstance(rev_obj, dict):
        # cantidad
        amount = None
        for k in ("amount", "value", "revenue", "sales"):
            if k in {_norm(x) for x in rev_obj.keys()}:
                for orig_k, v in rev_obj.items():
                    if _norm(orig_k) == k:
                        amount = _parse_number(v)
                        break
        if amount is None:
            # fallback: busca primer número razonable en el dict
            for v in rev_obj.values():
                n = _parse_number(v)
                if n and n > 1e6:
                    amount = n
                    break
        amount_txt = _fmt_billions_usd(amount) if amount else "an undisclosed amount"

        # fecha FY
        fy = ""
        for key in ("fiscal year close date", "fiscal year", "as of", "date"):
            fy = _get_first_str(rev_obj, [key])
            if fy:
                break
        fy_txt = _parse_date(fy) if fy else "the latest fiscal year"

        # links
        link_main = _choose_link(
            rev_obj.get("Source"),
            rev_obj.get("URL"),
            rev_obj,
        )
        link_sec = _choose_link(rev_obj.get("SEC Source"), rev_obj.get("SEC URL"))

        # bullet principal
        _add_bullet_runs(
            tf,
            [
                {
                    "text": f"{company_name} reported an annual revenue of {amount_txt} for the fiscal year ending {fy_txt} ",
                    "link": None,
                },
                *(
                    [{"text": "(", "link": None}, {"text": link_main, "link": link_main}, {"text": ").", "link": None}]
                    if link_main
                    else [{"text": ".", "link": None}]
                ),
            ],
            level=0,
            size=14,
        )

        # sub-bullet SEC opcional
        if link_sec:
            _add_bullet_runs(
                tf,
                [
                    {"text": "Additional filing: ", "link": None, "bold": False},
                    {"text": link_sec, "link": link_sec},
                ],
                level=1,
                size=12,
            )

    # --- Industry Average Gross Margin ---
    ind_gm_obj = _deep_find(payload, ["industry average gross margin", "industry gross margin", "industry avg"])
    if isinstance(ind_gm_obj, dict):
        industry = ""
        for key in ("industry", "sector"):
            s = _get_first_str(ind_gm_obj, [key])
            if s:
                industry = s
                break
        gm_avg = None
        for key in ("average gross margin", "gross margin", "avg", "average"):
            s = _get_first_str(ind_gm_obj, [key])
            if s:
                gm_avg = _parse_percent(s)
                break
        gm_txt = f"{gm_avg:.2f}%" if gm_avg is not None else "an unspecified value"

        link_ind = _choose_link(ind_gm_obj.get("Source"), ind_gm_obj.get("URL"), ind_gm_obj)

        _add_bullet_runs(
            tf,
            [
                {
                    "text": f'The industry average gross margin for the "{industry or "industry"}" industry is approximately {gm_txt} ',
                    "link": None,
                },
                *(
                    [{"text": "(", "link": None}, {"text": link_ind, "link": link_ind}, {"text": ").", "link": None}]
                    if link_ind
                    else [{"text": ".", "link": None}]
                ),
            ],
            level=0,
            size=14,
        )

    # --- Company Gross Margin ---
    comp_gm_obj = _deep_find(payload, ["company gross margin", "gross margin"])
    if isinstance(comp_gm_obj, dict):
        gm = None
        for key in ("gross margin", "margin"):
            s = _get_first_str(comp_gm_obj, [key])
            if s:
                gm = _parse_percent(s)
                break
        gm_txt = f"{gm:.2f}%" if gm is not None else "an unspecified value"

        fy = ""
        for key in ("fiscal year close date", "fiscal year", "as of", "date"):
            fy = _get_first_str(comp_gm_obj, [key])
            if fy:
                break
        fy_txt = _parse_date(fy) if fy else "the latest fiscal year"

        link_cmp = _choose_link(comp_gm_obj.get("Source"), comp_gm_obj.get("URL"), comp_gm_obj)

        # si tenemos industry avg, comparamos
        tail = ""
        try:
            ind_val = None
            for key in ("average gross margin", "gross margin", "avg", "average"):
                s = _get_first_str(ind_gm_obj or {}, [key])
                if s:
                    ind_val = _parse_percent(s)
                    break
            if ind_val is not None and gm is not None and abs(ind_val - gm) < 1e-6:
                tail = ", matching the industry average"
        except Exception:
            pass

        _add_bullet_runs(
            tf,
            [
                {
                    "text": f"The company's gross margin for the fiscal year ending {fy_txt} was {gm_txt}{tail} ",
                    "link": None,
                },
                *(
                    [{"text": "(", "link": None}, {"text": link_cmp, "link": link_cmp}, {"text": ").", "link": None}]
                    if link_cmp
                    else [{"text": ".", "link": None}]
                ),
            ],
            level=0,
            size=14,
        )

    # --- Employee Count / Headcount ---
    emp_obj = _deep_find(payload, ["employee count", "headcount", "employees"])
    if isinstance(emp_obj, dict):
        headcount = None
        # busca número
        for k in emp_obj.keys():
            nk = _norm(k)
            if any(s in nk for s in ["headcount", "employees", "employee count", "count", "total"]):
                headcount = emp_obj[k]
                break
        # formateo
        if isinstance(headcount, (int, float)):
            hc_txt = f"{int(headcount):,}"
        else:
            # intentar parsear si viene como string
            n = _parse_number(headcount)
            hc_txt = f"{int(n):,}" if n else (str(headcount) if headcount is not None else "an unspecified number")
        hc_txt = hc_txt.replace(",", ",")  # miles estándar

        asof = ""
        for key in ("as of", "date", "fiscal year close date", "fiscal year"):
            asof = _get_first_str(emp_obj, [key])
            if asof:
                break
        asof_txt = _parse_date(asof) if asof else "the stated date"

        link_emp = _choose_link(emp_obj.get("Source"), emp_obj.get("URL"), emp_obj)

        _add_bullet_runs(
            tf,
            [
                {"text": f"The company had {hc_txt} employees as of {asof_txt} ", "link": None},
                *(
                    [{"text": "(", "link": None}, {"text": link_emp, "link": link_emp}, {"text": ").", "link": None}]
                    if link_emp
                    else [{"text": ".", "link": None}]
                ),
            ],
            level=0,
            size=14,
        )
