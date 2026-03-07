"""
Artmidnet Mockup Server — app.py V3
------------------------------------
V1: Basic mockup generation (stretch + adapt modes)
V2: CORS support, health check endpoint
V3: Added /layers-report endpoint — generates a DOCX report
    from Wix page element tree data sent from Velo

Endpoints:
  POST /mockup          — generates room mockup image (existing)
  GET  /health          — health check (existing)
  POST /layers-report   — generates Layers Schema DOCX report (NEW in V3)
"""

from flask import Flask, request, jsonify, send_file
from flask_cors import CORS
import requests
import numpy as np
from PIL import Image
import io
import base64

# ── V3: ייבוא ספריות לייצור DOCX ──────────────────────────
import datetime
from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
# ────────────────────────────────────────────────────────────

app = Flask(__name__)
CORS(app)  # Allow requests from Wix


# ─────────────────────────────────────────────
# Helper: download image from URL → PIL Image
# ─────────────────────────────────────────────
def load_image_from_url(url: str) -> Image.Image:
    response = requests.get(url, timeout=15)
    response.raise_for_status()
    return Image.open(io.BytesIO(response.content)).convert("RGBA")


# ─────────────────────────────────────────────
# Helper: detect the outer black frame bounds
# Returns (left, top, right, bottom) in pixels
# ─────────────────────────────────────────────
def detect_outer_frame(img: Image.Image, dark_threshold: int = 60) -> tuple:
    arr = np.array(img.convert("RGB"))
    h, w = arr.shape[:2]

    def is_dark_row(row_pixels):
        return np.mean(row_pixels) < dark_threshold

    def is_dark_col(col_pixels):
        return np.mean(col_pixels) < dark_threshold

    top = 0
    for y in range(h):
        if is_dark_row(arr[y]):
            top = y
            break

    bottom = h - 1
    for y in range(h - 1, -1, -1):
        if is_dark_row(arr[y]):
            bottom = y
            break

    left = 0
    for x in range(w):
        if is_dark_col(arr[:, x]):
            left = x
            break

    right = w - 1
    for x in range(w - 1, -1, -1):
        if is_dark_col(arr[:, x]):
            right = x
            break

    return left, top, right, bottom


# ─────────────────────────────────────────────
# Helper: detect inner canvas area
# Scans inward from the outer frame edges
# looking for a brighter region
# ─────────────────────────────────────────────
def detect_inner_canvas(
    arr: np.ndarray,
    outer: tuple,
    bright_threshold: int = 100,
    step: int = 1
) -> tuple:
    left_o, top_o, right_o, bottom_o = outer

    inner_top = top_o
    for y in range(top_o, bottom_o):
        if np.mean(arr[y, left_o:right_o]) > bright_threshold:
            inner_top = y
            break

    inner_bottom = bottom_o
    for y in range(bottom_o, top_o, -1):
        if np.mean(arr[y, left_o:right_o]) > bright_threshold:
            inner_bottom = y
            break

    inner_left = left_o
    for x in range(left_o, right_o):
        if np.mean(arr[top_o:bottom_o, x]) > bright_threshold:
            inner_left = x
            break

    inner_right = right_o
    for x in range(right_o, left_o, -1):
        if np.mean(arr[top_o:bottom_o, x]) > bright_threshold:
            inner_right = x
            break

    return inner_left, inner_top, inner_right, inner_bottom


# ─────────────────────────────────────────────
# Helper: extract shadow strips from left + bottom
# ─────────────────────────────────────────────
def extract_shadow(img_arr: np.ndarray, outer: tuple, thickness: int = 8) -> dict:
    l, t, r, b = outer
    return {
        "left":   img_arr[t:b+1, l:l+thickness].copy(),
        "bottom": img_arr[b-thickness:b+1, l:r+1].copy(),
    }


# ─────────────────────────────────────────────
# Helper: sample wall color near the frame edges
# ─────────────────────────────────────────────
def sample_wall_color(img_arr: np.ndarray, outer: tuple, margin: int = 20) -> tuple:
    l, t, r, b = outer
    y0 = max(0, t - margin)
    y1 = t
    x0 = l + (r - l) // 3
    x1 = r - (r - l) // 3
    patch = img_arr[y0:y1, x0:x1]
    if patch.size == 0:
        return (200, 200, 200, 255)
    mean = patch.mean(axis=(0, 1))
    return tuple(int(v) for v in mean[:4])


# ─────────────────────────────────────────────
# MODE 1: STRETCH
# ─────────────────────────────────────────────
def apply_stretch(room_img: Image.Image, painting_img: Image.Image) -> Image.Image:
    arr = np.array(room_img)
    outer = detect_outer_frame(room_img)
    inner = detect_inner_canvas(arr, outer)

    il, it, ir, ib = inner
    canvas_w = ir - il
    canvas_h = ib - it

    if canvas_w <= 0 or canvas_h <= 0:
        raise ValueError("Could not detect inner canvas area")

    painting_resized = painting_img.resize(
        (canvas_w, canvas_h), Image.LANCZOS
    ).convert("RGBA")

    result = room_img.copy().convert("RGBA")
    result.paste(painting_resized, (il, it), painting_resized)

    return result.convert("RGB")


# ─────────────────────────────────────────────
# MODE 2: ADAPT
# ─────────────────────────────────────────────
def apply_adapt(room_img: Image.Image, painting_img: Image.Image) -> Image.Image:
    arr = np.array(room_img.convert("RGBA"))
    outer = detect_outer_frame(room_img)
    inner = detect_inner_canvas(arr, outer)

    lo, to, ro, bo = outer
    li, ti, ri, bi = inner

    ft_left   = li - lo
    ft_top    = ti - to
    ft_right  = ro - ri
    ft_bottom = bo - bi

    orig_canvas_w = ri - li
    paint_w, paint_h = painting_img.size
    aspect = paint_h / paint_w
    new_canvas_w = orig_canvas_w
    new_canvas_h = int(new_canvas_w * aspect)

    new_frame_w = new_canvas_w + ft_left + ft_right
    new_frame_h = new_canvas_h + ft_top  + ft_bottom

    shadow = extract_shadow(arr, outer)
    wall_color = sample_wall_color(arr, outer)

    frame_arr = np.zeros((new_frame_h, new_frame_w, 4), dtype=np.uint8)
    frame_arr[:, :] = (0, 0, 0, 255)

    painting_resized = painting_img.resize(
        (new_canvas_w, new_canvas_h), Image.LANCZOS
    ).convert("RGBA")
    p_arr = np.array(painting_resized)
    frame_arr[ft_top:ft_top+new_canvas_h, ft_left:ft_left+new_canvas_w] = p_arr

    frame_img = Image.fromarray(frame_arr, "RGBA")

    result = room_img.copy().convert("RGBA")
    paste_x = lo
    paste_y = to

    result_arr = np.array(result)
    result_arr[to:bo+1, lo:ro+1] = wall_color
    result = Image.fromarray(result_arr, "RGBA")

    result.paste(frame_img, (paste_x, paste_y), frame_img)

    result_arr2 = np.array(result)

    sh_left = shadow["left"]
    if sh_left.size > 0:
        new_sh_left = np.array(
            Image.fromarray(sh_left).resize(
                (sh_left.shape[1], new_frame_h), Image.LANCZOS
            )
        )
        x0 = paste_x - sh_left.shape[1]
        if x0 >= 0:
            result_arr2[paste_y:paste_y+new_frame_h, x0:paste_x] = new_sh_left

    sh_bot = shadow["bottom"]
    if sh_bot.size > 0:
        new_sh_bot = np.array(
            Image.fromarray(sh_bot).resize(
                (new_frame_w, sh_bot.shape[0]), Image.LANCZOS
            )
        )
        y0 = paste_y + new_frame_h
        if y0 + sh_bot.shape[0] <= result_arr2.shape[0]:
            result_arr2[y0:y0+sh_bot.shape[0], paste_x:paste_x+new_frame_w] = new_sh_bot

    return Image.fromarray(result_arr2, "RGBA").convert("RGB")


# ─────────────────────────────────────────────
# Helper: PIL Image → base64 JPEG string
# ─────────────────────────────────────────────
def image_to_base64(img: Image.Image, fmt: str = "JPEG") -> str:
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format=fmt, quality=90)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


# ─────────────────────────────────────────────
# API ENDPOINT: POST /mockup  (קיים מ-V1)
# ─────────────────────────────────────────────
@app.route("/mockup", methods=["POST"])
def mockup():
    try:
        data = request.get_json(force=True)
        room_url     = data.get("room_url")
        painting_url = data.get("painting_url")
        mode         = data.get("mode", "stretch").lower()

        if not room_url or not painting_url:
            return jsonify({"error": "room_url and painting_url are required"}), 400

        if mode not in ("stretch", "adapt"):
            return jsonify({"error": "mode must be 'stretch' or 'adapt'"}), 400

        room_img     = load_image_from_url(room_url)
        painting_img = load_image_from_url(painting_url)

        if mode == "stretch":
            result = apply_stretch(room_img, painting_img)
        else:
            result = apply_adapt(room_img, painting_img)

        return jsonify({
            "status": "ok",
            "mode": mode,
            "image_base64": image_to_base64(result)
        })

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"Failed to download image: {str(e)}"}), 400
    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Internal error: {str(e)}"}), 500


# ─────────────────────────────────────────────
# Health check  (קיים מ-V2)
# ─────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "artmidnet-mockup"})


# ═════════════════════════════════════════════
# V3 — LAYERS REPORT
# ─────────────────────────────────────────────
# פונקציות עזר לבניית ה-DOCX
# ═════════════════════════════════════════════

# מיפוי צבעים לפי סוג אלמנט — לכותרות ולטקסט בטבלה
TYPE_COLORS = {
    'Section':       'C00000',   # אדום כהה
    'Box':           '2E75B6',   # כחול
    'Text':          '375623',   # ירוק כהה
    'Button':        '7030A0',   # סגול
    'Repeater':      'C55A11',   # כתום
    'Image':         '006400',   # ירוק
    'VectorImage':   '006400',   # ירוק
    'Menu':          '595959',   # אפור
    'MenuContainer': '595959',   # אפור
}

def get_type_color(element_type):
    """מחזיר קוד צבע hex לפי סוג האלמנט"""
    return TYPE_COLORS.get(element_type, '404040')

def hex_to_rgb(hex_str):
    """ממיר hex string ל-tuple של (R, G, B)"""
    h = hex_str.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def set_cell_bg(cell, hex_color):
    """מגדיר צבע רקע לתא בטבלה"""
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'),   'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'),  hex_color)
    tcPr.append(shd)

def set_cell_font(cell, size_pt, color_hex, bold=False, indent_spaces=0):
    """מגדיר פונט לתא — גודל, צבע, bold, ואינדנטציה"""
    p = cell.paragraphs[0]
    text = p.text
    p.clear()
    run = p.add_run((' ' * indent_spaces) + text if indent_spaces else text)
    run.font.size       = Pt(size_pt)
    run.font.bold       = bold
    run.font.color.rgb  = RGBColor(*hex_to_rgb(color_hex))


# ─────────────────────────────────────────────
# API ENDPOINT: POST /layers-report  (חדש ב-V3)
#
# Body (JSON):
#   page_name — שם הדף ב-Wix (למשל "Z_LayersSchema")
#   elements  — מערך של אובייקטים:
#               { id, type, depth, parent }
#
# Response:
#   קובץ DOCX להורדה ישירה
# ─────────────────────────────────────────────
@app.route("/layers-report", methods=["POST"])
def layers_report():
    try:
        data      = request.get_json(force=True)
        page_name = data.get("page_name", "Unknown Page")
        elements  = data.get("elements", [])

        if not elements:
            return jsonify({"error": "elements array is required"}), 400

        doc = DocxDocument()

        # הגדרת שוליים לדף
        section = doc.sections[0]
        section.top_margin    = Inches(1)
        section.bottom_margin = Inches(1)
        section.left_margin   = Inches(1)
        section.right_margin  = Inches(1)

        # ── כותרת ראשית ──────────────────────────
        title = doc.add_paragraph()
        title.alignment = WD_ALIGN_PARAGRAPH.LEFT
        run = title.add_run('Artmidnet — Layers Report')
        run.bold           = True
        run.font.size      = Pt(20)
        run.font.color.rgb = RGBColor(0x1F, 0x38, 0x64)

        # שורת תת-כותרת: שם דף + תאריך + מספר אלמנטים
        sub = doc.add_paragraph()
        run2 = sub.add_run(
            f'Page: {page_name}  |  '
            f'{datetime.date.today().strftime("%d/%m/%Y")}  |  '
            f'{len(elements)} elements'
        )
        run2.font.size      = Pt(10)
        run2.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        run2.italic         = True

        doc.add_paragraph()  # רווח

        # ── סיכום לפי סוג ────────────────────────
        h2 = doc.add_paragraph()
        r  = h2.add_run('Summary by Type')
        r.bold           = True
        r.font.size      = Pt(13)
        r.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

        # ספירת כמות אלמנטים לכל סוג
        type_counts = {}
        for el in elements:
            t = el.get('type', 'Unknown')
            type_counts[t] = type_counts.get(t, 0) + 1

        # בניית טבלת סיכום
        summary_table = doc.add_table(rows=1, cols=2)
        summary_table.style = 'Table Grid'

        # כותרת הטבלה — רקע כהה + טקסט לבן
        hdr = summary_table.rows[0].cells
        for i, txt in enumerate(['Type', 'Count']):
            hdr[i].text = txt
            set_cell_bg(hdr[i], '1F3864')
            run = hdr[i].paragraphs[0].runs[0]
            run.bold           = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.size      = Pt(10)

        # שורה לכל סוג — ממוין מהנפוץ לנדיר
        for t, count in sorted(type_counts.items(), key=lambda x: -x[1]):
            row   = summary_table.add_row().cells
            color = get_type_color(t)
            rgb   = hex_to_rgb(color)
            row[0].text = t
            row[1].text = str(count)
            for cell in row:
                set_cell_bg(cell, 'F8F8F8')
                p = cell.paragraphs[0]
                if p.runs:
                    p.runs[0].font.color.rgb = RGBColor(*rgb)
                    p.runs[0].font.size      = Pt(10)

        doc.add_paragraph()  # רווח

        # ── טבלה מלאה ────────────────────────────
        h2b = doc.add_paragraph()
        r2  = h2b.add_run('Full Elements Tree')
        r2.bold           = True
        r2.font.size      = Pt(13)
        r2.font.color.rgb = RGBColor(0x2E, 0x75, 0xB6)

        note = doc.add_paragraph()
        rn = note.add_run('ID is indented by depth level. Depth = hierarchy level in Wix Layers.')
        rn.font.size      = Pt(9)
        rn.font.color.rgb = RGBColor(0x88, 0x88, 0x88)
        rn.italic         = True

        doc.add_paragraph()

        # בניית טבלה ראשית עם 4 עמודות
        main_table = doc.add_table(rows=1, cols=4)
        main_table.style = 'Table Grid'

        # כותרת הטבלה
        hdr2 = main_table.rows[0].cells
        for i, txt in enumerate(['ID', 'Type', 'Parent', 'Depth']):
            hdr2[i].text = txt
            set_cell_bg(hdr2[i], '1F3864')
            run = hdr2[i].paragraphs[0].runs[0]
            run.bold           = True
            run.font.color.rgb = RGBColor(0xFF, 0xFF, 0xFF)
            run.font.size      = Pt(10)

        # שורה לכל אלמנט
        for el in elements:
            el_id     = el.get('id',     '')
            el_type   = el.get('type',   '')
            el_parent = el.get('parent') or '—'
            el_depth  = int(el.get('depth', 1))
            color     = get_type_color(el_type)
            rgb       = hex_to_rgb(color)

            # אינדנטציה חזותית לפי עומק — רווחים לפני ה-ID
            indent  = '    ' * (el_depth - 1)
            row_bg  = 'EBF3FB' if el_depth % 2 == 0 else 'FFFFFF'

            row = main_table.add_row().cells
            row[0].text = indent + '#' + el_id
            row[1].text = el_type
            row[2].text = ('#' + el_parent) if el_parent != '—' else '—'
            row[3].text = str(el_depth)

            for cell in row:
                set_cell_bg(cell, row_bg)
                p = cell.paragraphs[0]
                if p.runs:
                    p.runs[0].font.size      = Pt(9)
                    p.runs[0].font.color.rgb = RGBColor(*rgb)

        # שמירת המסמך ל-buffer בזיכרון ושליחה כ-attachment
        buf = io.BytesIO()
        doc.save(buf)
        buf.seek(0)

        filename = f'Layers_Report_{page_name}_{datetime.date.today()}.docx'
        return send_file(
            buf,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        )

    except Exception as e:
        return jsonify({"error": f"Internal error: {str(e)}"}), 500


# ─────────────────────────────────────────────
# Run
# ─────────────────────────────────────────────
if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
