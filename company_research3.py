from pptx import Presentation
from pptx.enum.text import MSO_AUTO_SIZE
from helpers.utils import _find_shape_with_token, _add_section_header, _add_bullet, _norm, _is_url, _extract_urls, _parse_date, _add_bullet_runs


# --------------------------------------------------------------------
# Función 3 (genérica): {{CompanyResearch3}} -> bullets jerárquicos + links
# --------------------------------------------------------------------
def fill_company_research3(prs: Presentation, payload: dict):
    
    token="{{CompanyResearch3}}"
    
    slide, shape = _find_shape_with_token(prs, token)
    if not shape:
        return

    tf = shape.text_frame
    tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
    tf.clear()

    # -------- helpers internos sin configuraciones externas --------
    def _section_items(section_value):
        """
        Extrae la lista de ítems de una sección cualquiera de forma genérica:
        - dict con alguna lista (clave que contenga 'list/items/entries/highlights/data/points') -> esa lista + meta
        - lista directa -> items=list, meta=None
        - dict sin listas -> [dict] + dict (se trata como un solo ítem)
        - primitivo -> [{'Value': primitivo}] + None
        """
        if isinstance(section_value, dict):
            # buscar una lista "natural"
            for k, v in section_value.items():
                if isinstance(v, list) and any(w in _norm(k) for w in ["list", "items", "entries", "highlights", "data", "points"]):
                    return v, section_value
            # si no hubo, tomar la primera lista que aparezca
            for v in section_value.values():
                if isinstance(v, list):
                    return v, section_value
            # dict plano
            return [section_value], section_value
        elif isinstance(section_value, list):
            return section_value, None
        else:
            return [{"Value": section_value}], None

    def _score_main_kv(k, v):
        """
        Puntúa un par clave/valor para elegir la línea principal de un ítem:
        - Prefiere strings más largas
        - Le da un pequeño bonus a claves "título-like" (title/name/headline/summary)
        - Penaliza URLs puras
        """
        if isinstance(v, str):
            if _is_url(v):
                return 0.5  # URLs no son buen título
            base = min(len(v.strip()), 200) / 200.0  # normaliza por longitud
        elif isinstance(v, (int, float)):
            base = 0.4
        elif isinstance(v, dict):
            base = 0.3
        elif isinstance(v, list):
            base = 0.35
        else:
            base = 0.2

        nk = _norm(k)
        if any(w in nk for w in ["title", "name", "headline", "subject", "summary", "objective"]):
            base += 0.3
        return base

    def _choose_main_text(item_dict):
        """
        Elige texto principal del ítem sin depender de nombres fijos:
        - Máxima puntuación por _score_main_kv
        - Si no hay strings útiles, compacta como "k: v; ..."
        """
        best_k, best_v, best_score = None, None, -1.0
        for k, v in item_dict.items():
            sc = _score_main_kv(k, v)
            if sc > best_score:
                best_k, best_v, best_score = k, v, sc

        if isinstance(best_v, str) and best_v.strip():
            return best_k, best_v.strip()
        # si el "mejor" no es string, intentar otra string decente
        for k, v in item_dict.items():
            if isinstance(v, str) and v.strip() and not _is_url(v):
                return k, v.strip()
        # último recurso: compactar el dict
        try:
            return best_k, "; ".join(f"{k}: {v}" for k, v in item_dict.items() if v not in (None, ""))
        except Exception:
            # fallback duro
            return best_k, str(next(iter(item_dict.values()), ""))

    def _key_priority(k, v):
        """
        Orden genérico de subcampos:
        0: summary/description
        1: fechas (date/as of/fiscal)
        2: valores "normales" (texto/números)
        3: URLs y fuentes
        """
        nk = _norm(k)
        if any(w in nk for w in ["summary", "description", "details", "overview"]):
            return 0
        if any(w in nk for w in ["date", "as of", "fiscal year", "fy"]):
            return 1
        if isinstance(v, str) and _is_url(v):
            return 3
        if any(w in nk for w in ["url", "link", "source", "reference"]):
            return 3
        return 2

    def _order_subkeys(item_dict, main_key_used):
        keys = [k for k in item_dict.keys() if k != main_key_used]
        # ordenar por prioridad (+ alfabético estable)
        return sorted(keys, key=lambda k: (_key_priority(k, item_dict[k]), _norm(k)))

    def _section_suffix_from_meta(meta_dict):
        """Si en meta hay FY/As Of/Date, agregar sufijo legible al header."""
        if not isinstance(meta_dict, dict):
            return ""
        # elegir la primera fecha razonable
        for k, v in meta_dict.items():
            nk = _norm(k)
            if isinstance(v, str) and any(w in nk for w in ["fiscal year", "as of", "date"]):
                nice = _parse_date(v)
                if "fiscal year" in nk or nk == "fy":
                    return f" (FY {nice})" if nice else ""
                return f" ({nice})" if nice else ""
        return ""

    def _emit_value_as_bullets(label, value, level=1):
        """Render genérico de un valor como bullets/sub-bullets."""
        if value in (None, ""):
            return

        # URL pura
        if isinstance(value, str) and _is_url(value):
            _add_bullet_runs(tf, [{"text": f"{label}: "}, {"text": value, "link": value}], level=level, size=12)
            return

        # lista
        if isinstance(value, list):
            if all(isinstance(x, (str, int, float)) for x in value):
                for x in value:
                    _add_bullet(tf, f"{label}: {x}", level=level, size=12)
            else:
                for x in value:
                    if isinstance(x, dict):
                        mk, mv = _choose_main_text(x)
                        _add_bullet(tf, f"{label}: {mv}", level=level, size=12)
                        # URLs internas
                        for u in _extract_urls(x):
                            _add_bullet_runs(tf, [{"text": "link: "}, {"text": u, "link": u}], level=level+1, size=11)
                    else:
                        _add_bullet(tf, f"{label}: {x}", level=level, size=12)
            return

        # dict
        if isinstance(value, dict):
            mk, mv = _choose_main_text(value)
            # línea principal del sub-dict
            _add_bullet(tf, f"{label}: {mv}", level=level, size=12)
            # resto de campos del sub-dict
            for sk in _order_subkeys(value, mk):
                sv = value.get(sk)
                _emit_value_as_bullets(sk, sv, level=level+1)
            return

        # string/numérico
        nk = _norm(label)
        vtxt = _parse_date(value) if isinstance(value, str) and any(s in nk for s in ["date", "as of", "fiscal year", "fy"]) else str(value)
        _add_bullet(tf, f"{label}: {vtxt}", level=level, size=12)

    # -------- recorrido genérico de secciones (en orden de aparición) --------
    for section_name, section_value in payload.items():
        items, meta = _section_items(section_value)
        suffix = _section_suffix_from_meta(meta)
        _add_section_header(tf, f"{section_name}{suffix}:")

        for it in items:
            if isinstance(it, dict):
                mk, mv = _choose_main_text(it)
                _add_bullet(tf, mv, level=0, size=14)

                # subcampos del ítem
                for sk in _order_subkeys(it, mk):
                    sv = it.get(sk)
                    _emit_value_as_bullets(sk, sv, level=1)
            elif isinstance(it, list):
                # lista de primitivas en un ítem
                for x in it:
                    _add_bullet(tf, str(x), level=0, size=14)
            else:
                # primitivo
                _add_bullet(tf, str(it), level=0, size=14)