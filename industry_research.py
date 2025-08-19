from pptx import Presentation
from pptx.enum.text import PP_ALIGN, MSO_AUTO_SIZE
from pptx.dml.color import RGBColor
from pptx.util import Pt
from utils import estimate_row_height
from config import EMU_PER_PT


def fill_industry_slides(prs: Presentation, slide, payload: dict):
    """
    Reemplaza {{IndustryResearch}} en `slide` con múltiples tablas según payload:
    - Usa `payload_data['title']` como Título de slide.
    - Usa `payload_data['headers']` como columnas.
    - Usa `payload_data['rows']` como datos de fila.
    - Corta en varias slides cuando excede espacio disponible.
    """
    token = "{{IndustryResearch}}"
    # localiza y elimina placeholder
    pos = None    
    for shape in slide.shapes:
        if not shape.has_text_frame:
            continue
        tf = shape.text_frame

        if token in tf.text:
            pos = (shape.left, shape.top, shape.width, shape.height)
            slide.shapes._spTree.remove(shape._element)
            break
    if pos is None:
        return
    left, top, width, height = pos

    # Establece título de la slide desde payload
    if slide.shapes.title:
        slide.shapes.title.text = payload.get("title", "")

    headers = payload.get("headers", [])
    rows    = payload.get("rows", [])
    if not headers or not rows:
        return

    # alturas en puntos
    total_pt   = height / EMU_PER_PT
    header_pt  = 24
    content_pt = total_pt - header_pt
    line_pt    = 12

    keys = headers
    # calcula ancho de columna en pt para wrapping
    width_pt     = width / EMU_PER_PT
    col_width_pt = width_pt / len(keys)

    # calcular altura estimada de cada fila
    row_heights = [
        estimate_row_height(e, keys, line_pt, col_width_pt)
        for e in rows
    ]

    # particionar filas en trozos que quepan en content_pt
    chunks = []
    i, n = 0, len(rows)
    while i < n:
        used, j = 0.0, i
        while j < n and used + row_heights[j] <= content_pt:
            used += row_heights[j]
            j += 1
        # al menos una fila
        if j == i:
            j = i + 1
        chunks.append(rows[i:j])
        i = j

    layout = slide.slide_layout
    for idx, chunk in enumerate(chunks):
        target = slide if idx == 0 else prs.slides.add_slide(layout)
        if idx > 0 and target.shapes.title:
            target.shapes.title.text += " (cont.)"

        rows_count, cols_count = len(chunk) + 1, len(keys)
        tbl = target.shapes.add_table(
            rows_count, cols_count, left, top, width, height
        ).table

        # fila de encabezado
        tbl.rows[0].height = Pt(header_pt)
        for col in tbl.columns:
            col.width = width // cols_count
        for c, h in enumerate(keys):
            cell = tbl.cell(0, c)
            cell.text = h
            cell.fill.solid()
            cell.fill.fore_color.rgb = RGBColor(0, 70, 122)
            para = cell.text_frame.paragraphs[0]
            para.alignment = PP_ALIGN.CENTER
            for run in para.runs:
                run.font.size = Pt(14)
                run.font.bold = True
                run.font.color.rgb = RGBColor(255, 255, 255)
            cell.text_frame.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE

        # filas de datos
        for r, entry in enumerate(chunk, start=1):
            for c, h in enumerate(keys):
                cell = tbl.cell(r, c)
                tf = cell.text_frame
                tf.clear()
                tf.auto_size = MSO_AUTO_SIZE.TEXT_TO_FIT_SHAPE
                val = entry.get(h, "")
                if isinstance(val, list) and val and isinstance(val[0], dict):
                    for item in val:
                        line = "; ".join(f"{k}: {v}" for k, v in item.items())
                        p = tf.add_paragraph()
                        p.text = line
                        p.alignment = PP_ALIGN.LEFT
                        for run in p.runs:
                            run.font.size = Pt(10)
                elif isinstance(val, list):
                    for it in val:
                        p = tf.add_paragraph()
                        p.text = f"• {it}"
                        p.alignment = PP_ALIGN.LEFT
                        for run in p.runs:
                            run.font.size = Pt(10)
                elif isinstance(val, str) and "" in val:
                    for line in val.splitlines():
                        p = tf.add_paragraph()
                        p.text = line
                        p.alignment = PP_ALIGN.LEFT
                        for run in p.runs:
                            run.font.size = Pt(10)
                else:
                    p = tf.add_paragraph()
                    p.text = str(val)
                    p.alignment = PP_ALIGN.LEFT
                    for run in p.runs:
                        run.font.size = Pt(10)
