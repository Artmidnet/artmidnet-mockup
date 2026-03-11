"""
Artmidnet Mockup Server — app.py V18
------------------------------------
V1: Basic mockup generation (stretch + adapt modes)
V2: CORS support, health check endpoint
V3: Added /layers-report endpoint — generates DOCX, returns file directly
V4: /layers-report now returns base64 JSON instead of file download
V5: Page name displayed prominently as main title in report header
V6: Added /cms-report endpoint — generates CMS Schema DOCX, returns base64 JSON
V7: MERGE — restored /noframe, /zoom, /rect from V2 (were missing in V6)
V8: Fixed apply_adapt — shadow paste array shape mismatch (broadcast error)
V9: Fixed apply_zoom — use detect_outer_frame+detect_inner_canvas instead of detect_white_area
V10: Fixed apply_zoom — clip painting to inner canvas bounds (horizontal overflow fix)
V11: Fixed apply_zoom — constrain painting size to inner canvas (no overflow in any orientation)
V12: New approach for zoom+rect — detect frame, create white canvas with painting AR, add border+shadow
V13: Test — only steps A+B+C (frame detection + white canvas), no painting/border/shadow
V14: Fix detect_outer_frame — scan single pixel at center column/row instead of full row/col mean
V15: Fix detect_outer_frame — use narrow center strip (5%) mean instead of single pixel, more robust
V16: New detect_outer_frame — find largest bright canvas region, expand to cover black border
V17: New approach — detect red dot (ImagePoint) in mockup, place white canvas using size_px + painting AR
V18: שלב D — הלבשת תמונת המקור על המשטח הלבן

Endpoints:
  GET  /health          — health check
  POST /mockup          — generates room mockup image (stretch / adapt)
  POST /noframe         — BuildMockupNoframe — painting centered on light background
  POST /zoom            — BuildMockupZoom    — painting fills 80% of mockup
  POST /rect            — BuildMockupRect    — adapts frame AR to painting
  POST /layers-report   — generates Layers Schema DOCX, returns base64 JSON
  POST /cms-report      — generates CMS Schema DOCX, returns base64 JSON
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import requests
import numpy as np
from PIL import Image
import io
import base64
import datetime

from docx import Document as DocxDocument
from docx.shared import Pt, RGBColor, Inches
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

app = Flask(__name__)
CORS(app)


# ─────────────────────────────────────────────
# Helper: download image from URL → PIL Image
# ─────────────────────────────────────────────
def load_image_from_url(url: str) -> Image.Image:
    response = requests.get(url, timeout=30)
    response.raise_for_status()
    return Image.open(io.BytesIO(response.content)).convert("RGBA")


# ─────────────────────────────────────────────
# Helper: detect outer black frame bounds
# ─────────────────────────────────────────────
def detect_red_dot(img: Image.Image) -> tuple:
    """
    V17: Find center of red dot (ImagePoint) in mockup template.
    Red dot criteria: R > 180, G < 80, B < 80.
    Returns (cx, cy) center of the red region.
    """
    arr = np.array(img.convert("RGB"))
    red_mask = (arr[:, :, 0] > 180) & (arr[:, :, 1] < 80) & (arr[:, :, 2] < 80)
    ys, xs = np.where(red_mask)
    if len(xs) == 0:
        h, w = arr.shape[:2]
        cx, cy = w // 2, h // 2
        print(f"V17 detect_red_dot: NOT FOUND — using center ({cx},{cy})")
        return cx, cy
    cx = int(np.mean(xs))
    cy = int(np.mean(ys))
    print(f"V17 detect_red_dot: found {len(xs)} red pixels | center=({cx},{cy})")
    return cx, cy


# ─────────────────────────────────────────────
# Helper: detect inner canvas (bright area)
# ─────────────────────────────────────────────
def detect_inner_canvas(arr: np.ndarray, outer: tuple, bright_threshold: int = 100) -> tuple:
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
# Helper: detect white area in mockup image
# ─────────────────────────────────────────────
def detect_white_area(img: Image.Image, white_threshold: int = 240) -> tuple:
    arr = np.array(img.convert("RGB"))
    h, w = arr.shape[:2]

    white_mask = np.all(arr > white_threshold, axis=2)

    rows = np.any(white_mask, axis=1)
    cols = np.any(white_mask, axis=0)

    row_indices = np.where(rows)[0]
    col_indices = np.where(cols)[0]

    if len(row_indices) == 0 or len(col_indices) == 0:
        return w // 4, h // 4, 3 * w // 4, 3 * h // 4

    return int(col_indices[0]), int(row_indices[0]), int(col_indices[-1]), int(row_indices[-1])


# ─────────────────────────────────────────────
# Helper: extract shadow strips
# ─────────────────────────────────────────────
def extract_shadow(img_arr: np.ndarray, outer: tuple, thickness: int = 8) -> dict:
    l, t, r, b = outer
    return {
        "left":   img_arr[t:b + 1, l:l + thickness].copy(),
        "bottom": img_arr[b - thickness:b + 1, l:r + 1].copy(),
    }


# ─────────────────────────────────────────────
# Helper: sample wall color near frame
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
# Helper: sample background color from corners
# ─────────────────────────────────────────────
def sample_corner_color(img: Image.Image, corner_size: int = 60) -> tuple:
    arr = np.array(img.convert("RGBA"))
    h, w = arr.shape[:2]
    cs = min(corner_size, h // 4, w // 4)
    corners = [
        arr[:cs, :cs],
        arr[:cs, w - cs:],
        arr[h - cs:, :cs],
        arr[h - cs:, w - cs:]
    ]
    mean = np.mean([c.mean(axis=(0, 1)) for c in corners], axis=0)
    return tuple(int(v) for v in mean)


# ─────────────────────────────────────────────
# Helper: PIL Image → base64 JPEG
# ─────────────────────────────────────────────
def image_to_base64(img: Image.Image, fmt: str = "JPEG") -> str:
    buf = io.BytesIO()
    img.convert("RGB").save(buf, format=fmt, quality=90)
    buf.seek(0)
    return base64.b64encode(buf.read()).decode("utf-8")


# ─────────────────────────────────────────────
# MODE: STRETCH
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

    painting_resized = painting_img.resize((canvas_w, canvas_h), Image.LANCZOS).convert("RGBA")
    result = room_img.copy().convert("RGBA")
    result.paste(painting_resized, (il, it), painting_resized)
    return result.convert("RGB")


# ─────────────────────────────────────────────
# MODE: ADAPT
# V8: Fixed shadow paste — clip to image bounds to prevent shape mismatch
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

    painting_resized = painting_img.resize((new_canvas_w, new_canvas_h), Image.LANCZOS).convert("RGBA")
    p_arr = np.array(painting_resized)
    frame_arr[ft_top:ft_top + new_canvas_h, ft_left:ft_left + new_canvas_w] = p_arr

    frame_img = Image.fromarray(frame_arr, "RGBA")
    result = room_img.copy().convert("RGBA")

    result_arr = np.array(result)
    img_h, img_w = result_arr.shape[:2]

    result_arr[to:bo + 1, lo:ro + 1] = wall_color
    result = Image.fromarray(result_arr, "RGBA")
    result.paste(frame_img, (lo, to), frame_img)

    result_arr2 = np.array(result)

    # V8: הגנה על shadow paste — חיתוך לגבולות התמונה
    sh_left = shadow["left"]
    if sh_left.size > 0:
        new_sh_left = np.array(Image.fromarray(sh_left).resize(
            (sh_left.shape[1], new_frame_h), Image.LANCZOS))
        x0 = lo - sh_left.shape[1]
        if x0 >= 0:
            y_end = min(to + new_frame_h, img_h)
            actual_h = y_end - to
            result_arr2[to:y_end, x0:lo] = new_sh_left[:actual_h, :]

    sh_bot = shadow["bottom"]
    if sh_bot.size > 0:
        new_sh_bot = np.array(Image.fromarray(sh_bot).resize(
            (new_frame_w, sh_bot.shape[0]), Image.LANCZOS))
        y0 = to + new_frame_h
        if y0 < img_h:
            y_end = min(y0 + sh_bot.shape[0], img_h)
            actual_h = y_end - y0
            x_end = min(lo + new_frame_w, img_w)
            actual_w = x_end - lo
            result_arr2[y0:y_end, lo:x_end] = new_sh_bot[:actual_h, :actual_w]

    return Image.fromarray(result_arr2, "RGBA").convert("RGB")


# ─────────────────────────────────────────────
# MODE: NOFRAME
# Centers painting on 2000x2000 light background
# ─────────────────────────────────────────────
def apply_noframe(painting_img: Image.Image) -> Image.Image:
    pw, ph = painting_img.size
    if max(pw, ph) != 2000:
        scale = 2000 / max(pw, ph)
        pw = int(pw * scale)
        ph = int(ph * scale)
        painting_img = painting_img.resize((pw, ph), Image.LANCZOS)

    arr = np.array(painting_img.convert("RGB")).reshape(-1, 3).astype(float)
    indices = np.random.choice(len(arr), min(300, len(arr)), replace=False)
    chosen = arr[indices[np.random.randint(len(indices))]]
    light = tuple(int(230 + (c / 255.0) * 20) for c in chosen) + (255,)

    canvas = Image.new("RGBA", (2000, 2000), light)

    x = (2000 - pw) // 2
    y = (2000 - ph) // 2

    painting_rgba = painting_img.convert("RGBA")
    canvas.paste(painting_rgba, (x, y), painting_rgba)

    return canvas.convert("RGB")


# ─────────────────────────────────────────────
# MODE: ZOOM
# Painting fills 80% of mockup image
# ─────────────────────────────────────────────
def apply_new_mockup(painting_img: Image.Image, mockup_img: Image.Image, size_px: int = 800) -> Image.Image:
    """
    V17: Steps A+B+C using red dot (ImagePoint) and size_px.
    A) Find red dot center = ImagePoint (cx, cy)
    B) Calculate painting AR = h/w
    C) Build white canvas: longest side = size_px, keep AR
       Center it on ImagePoint, paste on mockup
    """
    # ── A: find red dot ──
    img_cx, img_cy = detect_red_dot(mockup_img)

    # ── B: painting AR ──
    pw, ph = painting_img.size
    ar = ph / pw  # height / width
    print(f"V17 painting: {pw}x{ph} AR={ar:.3f} | size_px={size_px}")

    # ── C: white canvas sized by size_px and AR ──
    if ar <= 1.0:
        # landscape or square: width is longest side
        wc_w = size_px
        wc_h = int(size_px * ar)
    else:
        # portrait: height is longest side
        wc_h = size_px
        wc_w = int(size_px / ar)

    print(f"V18 white canvas: {wc_w}x{wc_h} | ImagePoint=({img_cx},{img_cy})")

    # ── D: resize painting to fit white canvas, paste ──
    painting_resized = painting_img.convert("RGBA").resize((wc_w, wc_h), Image.LANCZOS)

    # paste painting centered on ImagePoint
    result = mockup_img.copy().convert("RGBA")
    img_w, img_h = result.size
    paste_x = img_cx - wc_w // 2
    paste_y = img_cy - wc_h // 2
    paste_x = max(0, min(paste_x, img_w - wc_w))
    paste_y = max(0, min(paste_y, img_h - wc_h))

    result.paste(painting_resized, (paste_x, paste_y), painting_resized)
    print(f"V18 painting pasted at ({paste_x},{paste_y})")
    return result.convert("RGB")


# ─────────────────────────────────────────────
# MODE: ZOOM — painting fills mockup frame
# ─────────────────────────────────────────────
def apply_zoom(painting_img: Image.Image, mockup_img: Image.Image, size_px: int = 800) -> Image.Image:
    return apply_new_mockup(painting_img, mockup_img, size_px)


# ─────────────────────────────────────────────
# MODE: RECT — adapts frame AR to painting
# ─────────────────────────────────────────────
def apply_rect(painting_img: Image.Image, mockup_img: Image.Image, size_px: int = 800) -> Image.Image:
    return apply_new_mockup(painting_img, mockup_img, size_px)


# ═════════════════════════════════════════════
# DOCX Helpers (layers-report + cms-report)
# ═════════════════════════════════════════════

TYPE_COLORS = {
    'Section': 'C00000', 'Box': '2E75B6', 'Text': '375623',
    'Button': '7030A0', 'Repeater': 'C55A11', 'Image': '006400',
    'VectorImage': '006400', 'Menu': '595959', 'MenuContainer': '595959',
}

FIELD_TYPE_COLORS = {
    'TEXT':          '1F3864',
    'NUMBER':        '375623',
    'BOOLEAN':       'C55A11',
    'DATE':          '7030A0',
    'IMAGE':         '006400',
    'MEDIA_GALLERY': '006400',
    'REFERENCE':     'C00000',
    'ARRAY_STRING':  '2E75B6',
    'ARRAY':         '2E75B6',
    'OBJECT':        '595959',
    'RICH_TEXT':     '404040',
}

def get_type_color(element_type):
    return TYPE_COLORS.get(element_type, '404040')

def get_field_type_color(field_type):
    return FIELD_TYPE_COLORS.get(field_type, '404040')

def hex_to_rgb(hex_str):
    h = hex_str.lstrip('#')
    return tuple(int(h[i:i+2], 16) for i in (0, 2, 4))

def set_cell_bg(cell, hex_color):
    tc = cell._tc
    tcPr = tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), hex_color)
    tcPr.append(shd)


# ═════════════════════════════════════════════
# ENDPOINTS: IMAGE MOCKUPS
# ═════════════════════════════════════════════

@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "artmidnet-mockup", "version": "V18"})


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
        result = apply_stretch(room_img, painting_img) if mode == "stretch" else apply_adapt(room_img, painting_img)

        return jsonify({"status": "ok", "mode": mode, "image_base64": image_to_base64(result)})

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"Failed to download image: {str(e)}"}), 400
    except ValueError as e:
        return jsonify({"error": str(e)}), 422
    except Exception as e:
        return jsonify({"error": f"Internal error: {str(e)}"}), 500


@app.route("/noframe", methods=["POST"])
def noframe():
    try:
        data = request.get_json(force=True)
        painting_url = data.get("painting_url")
        if not painting_url:
            return jsonify({"error": "painting_url is required"}), 400

        painting_img = load_image_from_url(painting_url)
        result = apply_noframe(painting_img)

        return jsonify({"status": "ok", "image_base64": image_to_base64(result)})

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"Failed to download image: {str(e)}"}), 400
    except Exception as e:
        return jsonify({"error": f"Internal error: {str(e)}"}), 500


@app.route("/zoom", methods=["POST"])
def zoom():
    try:
        data = request.get_json(force=True)
        painting_url = data.get("painting_url")
        mockup_url   = data.get("mockup_url")
        if not painting_url or not mockup_url:
            return jsonify({"error": "painting_url and mockup_url are required"}), 400

        painting_img = load_image_from_url(painting_url)
        mockup_img   = load_image_from_url(mockup_url)
        size_px = int(data.get("size_px", 800))
        result = apply_zoom(painting_img, mockup_img, size_px)

        return jsonify({"status": "ok", "image_base64": image_to_base64(result)})

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"Failed to download image: {str(e)}"}), 400
    except Exception as e:
        return jsonify({"error": f"Internal error: {str(e)}"}), 500


@app.route("/rect", methods=["POST"])
def rect():
    try:
        data = request.get_json(force=True)
        painting_url = data.get("painting_url")
        mockup_url   = data.get("mockup_url")
        if not painting_url or not mockup_url:
            return jsonify({"error": "painting_url and mockup_url are required"}), 400

        painting_img = load_image_from_url(painting_url)
        mockup_img   = load_image_from_url(mockup_url)
        size_px = int(data.get("size_px", 800))
        result = apply_rect(painting_img, mockup_img, size_px)

        return jsonify({"status": "ok", "image_base64": image_to_base64(result)})

    except requests.exceptions.RequestException as e:
        return jsonify({"error": f"Failed to download image: {str(e)}"}), 400
    except Exception as e:
        return jsonify({"error": f"Internal error: {str(e)}"}), 500


# ═════════════════════════════════════════════
# ENDPOINTS: DOCX REPORTS
# ═════════════════════════════════════════════

@app.route("/layers-report", methods=["POST"])
def layers_report():
    try:
        data      = request.get_json(force=True)
        page_name = data.get("page_name", "Unknown Page")
        elements  = data.get("elements", [])
        if not elements:
            return jsonify({"error": "elements array is required"}), 400

        doc = DocxDocument()
        section = doc.sections[0]
        section.top_margin = section.bottom_margin = Inches(1)
        section.left_margin = section.right_margin = Inches(1)

        site_label = doc.add_paragraph()
        r0 = site_label.add_run('Artmidnet — Layers Report')
        r0.font.size = Pt(11); r0.font.color.rgb = RGBColor(0x88,0x88,0x88); r0.italic = True

        page_title = doc.add_paragraph()
        page_title.paragraph_format.space_before = Pt(4)
        page_title.paragraph_format.space_after  = Pt(2)
        r1 = page_title.add_run(f'Page: {page_name}')
        r1.bold = True; r1.font.size = Pt(26); r1.font.color.rgb = RGBColor(0x1F,0x38,0x64)
        pPr = page_title._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom = OxmlElement('w:bottom')
        bottom.set(qn('w:val'), 'single'); bottom.set(qn('w:sz'), '6')
        bottom.set(qn('w:space'), '1'); bottom.set(qn('w:color'), '1F3864')
        pBdr.append(bottom); pPr.append(pBdr)

        meta = doc.add_paragraph()
        meta.paragraph_format.space_before = Pt(6)
        r2 = meta.add_run(f'Date: {datetime.date.today().strftime("%d/%m/%Y")}   |   Total elements: {len(elements)}')
        r2.font.size = Pt(10); r2.font.color.rgb = RGBColor(0x88,0x88,0x88); r2.italic = True

        doc.add_paragraph()

        h2 = doc.add_paragraph()
        r = h2.add_run('Summary by Type')
        r.bold = True; r.font.size = Pt(13); r.font.color.rgb = RGBColor(0x2E,0x75,0xB6)

        type_counts = {}
        for el in elements:
            t = el.get('type', 'Unknown')
            type_counts[t] = type_counts.get(t, 0) + 1

        summary_table = doc.add_table(rows=1, cols=2)
        summary_table.style = 'Table Grid'
        hdr = summary_table.rows[0].cells
        for i, txt in enumerate(['Type', 'Count']):
            hdr[i].text = txt; set_cell_bg(hdr[i], '1F3864')
            run = hdr[i].paragraphs[0].runs[0]
            run.bold = True; run.font.color.rgb = RGBColor(0xFF,0xFF,0xFF); run.font.size = Pt(10)

        for t, count in sorted(type_counts.items(), key=lambda x: -x[1]):
            row = summary_table.add_row().cells
            color = get_type_color(t); rgb = hex_to_rgb(color)
            row[0].text = t; row[1].text = str(count)
            for cell in row:
                set_cell_bg(cell, 'F8F8F8')
                p = cell.paragraphs[0]
                if p.runs:
                    p.runs[0].font.color.rgb = RGBColor(*rgb); p.runs[0].font.size = Pt(10)

        doc.add_paragraph()
        h2b = doc.add_paragraph()
        r2b = h2b.add_run('Full Elements Tree')
        r2b.bold = True; r2b.font.size = Pt(13); r2b.font.color.rgb = RGBColor(0x2E,0x75,0xB6)

        note = doc.add_paragraph()
        rn = note.add_run('Depth column represents hierarchy level. ID is indented accordingly.')
        rn.font.size = Pt(9); rn.font.color.rgb = RGBColor(0x88,0x88,0x88); rn.italic = True
        doc.add_paragraph()

        main_table = doc.add_table(rows=1, cols=4)
        main_table.style = 'Table Grid'
        hdr2 = main_table.rows[0].cells
        for i, txt in enumerate(['ID', 'Type', 'Parent', 'Depth']):
            hdr2[i].text = txt; set_cell_bg(hdr2[i], '1F3864')
            run = hdr2[i].paragraphs[0].runs[0]
            run.bold = True; run.font.color.rgb = RGBColor(0xFF,0xFF,0xFF); run.font.size = Pt(10)

        for el in elements:
            el_id = el.get('id',''); el_type = el.get('type','')
            el_parent = el.get('parent') or '—'; el_depth = int(el.get('depth', 1))
            color = get_type_color(el_type); rgb = hex_to_rgb(color)
            indent = '    ' * (el_depth - 1); row_bg = 'EBF3FB' if el_depth % 2 == 0 else 'FFFFFF'
            row = main_table.add_row().cells
            row[0].text = indent + '#' + el_id; row[1].text = el_type
            row[2].text = ('#' + el_parent) if el_parent != '—' else '—'; row[3].text = str(el_depth)
            for cell in row:
                set_cell_bg(cell, row_bg)
                p = cell.paragraphs[0]
                if p.runs:
                    p.runs[0].font.size = Pt(9); p.runs[0].font.color.rgb = RGBColor(*rgb)

        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        docx_base64 = base64.b64encode(buf.read()).decode("utf-8")
        filename = f'Layers_Report_{page_name}_{datetime.date.today()}.docx'
        return jsonify({"status": "ok", "base64": docx_base64, "filename": filename})

    except Exception as e:
        return jsonify({"error": f"Internal error: {str(e)}"}), 500


@app.route("/cms-report", methods=["POST"])
def cms_report():
    try:
        data        = request.get_json(force=True)
        collections = data.get("collections", [])

        if not collections:
            return jsonify({"error": "collections array is required"}), 400

        total_fields = sum(len(col.get('fields', [])) for col in collections)

        doc = DocxDocument()
        section = doc.sections[0]
        section.top_margin = section.bottom_margin = Inches(1)
        section.left_margin = section.right_margin = Inches(1)

        site_label = doc.add_paragraph()
        r0 = site_label.add_run('Artmidnet — CMS Schema Report')
        r0.bold = True; r0.font.size = Pt(20); r0.font.color.rgb = RGBColor(0x1F,0x38,0x64)

        meta = doc.add_paragraph()
        meta.paragraph_format.space_before = Pt(2)
        r_meta = meta.add_run(
            f'Collections: {len(collections)}  |  {total_fields} Fields  |  '
            f'Generated: {datetime.date.today().strftime("%d/%m/%Y")}'
        )
        r_meta.font.size = Pt(10); r_meta.font.color.rgb = RGBColor(0x88,0x88,0x88); r_meta.italic = True

        pPr = meta._p.get_or_add_pPr()
        pBdr = OxmlElement('w:pBdr')
        bottom_bdr = OxmlElement('w:bottom')
        bottom_bdr.set(qn('w:val'), 'single'); bottom_bdr.set(qn('w:sz'), '6')
        bottom_bdr.set(qn('w:space'), '1'); bottom_bdr.set(qn('w:color'), '1F3864')
        pBdr.append(bottom_bdr); pPr.append(pBdr)

        doc.add_paragraph()

        for idx, col in enumerate(collections, 1):
            col_id   = col.get('collectionId', '')
            col_name = col.get('displayName', col_id)
            fields   = col.get('fields', [])

            col_heading = doc.add_paragraph()
            col_heading.paragraph_format.space_before = Pt(14)
            col_heading.paragraph_format.space_after  = Pt(2)
            r_col = col_heading.add_run(f'{idx}. Collection = {col_name}')
            r_col.bold = True; r_col.font.size = Pt(13); r_col.font.color.rgb = RGBColor(0x1F,0x38,0x64)

            col_sub = doc.add_paragraph()
            col_sub.paragraph_format.space_after = Pt(4)
            r_id = col_sub.add_run(f'id = {col_id}')
            r_id.font.size = Pt(10); r_id.font.color.rgb = RGBColor(0x88,0x88,0x88)

            col_count = doc.add_paragraph()
            col_count.paragraph_format.space_after = Pt(6)
            r_count = col_count.add_run(f'{len(fields)} fields')
            r_count.font.size = Pt(10); r_count.font.color.rgb = RGBColor(0x55,0x55,0x55)
            r_count.italic = True

            if fields:
                tbl = doc.add_table(rows=1, cols=4)
                tbl.style = 'Table Grid'

                hdr_cells = tbl.rows[0].cells
                for i, txt in enumerate(['Field Display Name', 'Field Type', 'Field Key', 'System']):
                    hdr_cells[i].text = txt
                    set_cell_bg(hdr_cells[i], '1F3864')
                    run = hdr_cells[i].paragraphs[0].runs[0]
                    run.bold = True
                    run.font.color.rgb = RGBColor(0xFF,0xFF,0xFF)
                    run.font.size = Pt(10)

                for f_idx, field in enumerate(fields):
                    f_name   = field.get('displayName', '')
                    f_type   = field.get('type', '')
                    f_key    = field.get('key', '')
                    f_system = 'Yes' if field.get('systemField') else 'No'
                    row_bg   = 'F8F8F8' if f_idx % 2 == 0 else 'FFFFFF'
                    type_rgb = hex_to_rgb(get_field_type_color(f_type))

                    row = tbl.add_row().cells
                    row[0].text = f_name
                    row[1].text = f_type
                    row[2].text = f_key
                    row[3].text = f_system

                    for c_idx, cell in enumerate(row):
                        set_cell_bg(cell, row_bg)
                        p = cell.paragraphs[0]
                        if p.runs:
                            p.runs[0].font.size = Pt(10)
                            if c_idx == 1:
                                p.runs[0].font.color.rgb = RGBColor(*type_rgb)
                            else:
                                p.runs[0].font.color.rgb = RGBColor(0x33,0x33,0x33)

            doc.add_paragraph()

        buf = io.BytesIO(); doc.save(buf); buf.seek(0)
        docx_base64 = base64.b64encode(buf.read()).decode("utf-8")
        filename = f'CMS_Schema_Report_{datetime.date.today()}.docx'

        return jsonify({"status": "ok", "base64": docx_base64, "filename": filename})

    except Exception as e:
        return jsonify({"error": f"Internal error: {str(e)}"}), 500


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
