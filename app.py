"""
Artmidnet Mockup Server
-----------------------
Flask API that replaces a painting inside a room frame image.
Supports two modes:
  - stretch: scales the painting to fit the existing frame exactly
  - adapt:   resizes the frame to match the painting's aspect ratio
              while preserving frame thickness and shadow
"""

from flask import Flask, request, jsonify
from flask_cors import CORS
import requests
import numpy as np
from PIL import Image
import io
import base64

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

    # Scan from each edge inward until pixels become bright
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
    # Sample from a small patch above the frame
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
# Scales the painting to fit exactly inside the
# detected inner canvas area
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

    # Resize painting to canvas dimensions
    painting_resized = painting_img.resize(
        (canvas_w, canvas_h), Image.LANCZOS
    ).convert("RGBA")

    # Paste onto room image
    result = room_img.copy().convert("RGBA")
    result.paste(painting_resized, (il, it), painting_resized)

    return result.convert("RGB")


# ─────────────────────────────────────────────
# MODE 2: ADAPT
# Resizes the frame to match painting's aspect
# ratio; rebuilds frame + shadow around painting
# ─────────────────────────────────────────────
def apply_adapt(room_img: Image.Image, painting_img: Image.Image) -> Image.Image:
    arr = np.array(room_img.convert("RGBA"))
    outer = detect_outer_frame(room_img)
    inner = detect_inner_canvas(arr, outer)

    lo, to, ro, bo = outer
    li, ti, ri, bi = inner

    # Frame thickness on each side
    ft_left   = li - lo
    ft_top    = ti - to
    ft_right  = ro - ri
    ft_bottom = bo - bi

    # Target canvas size = painting's natural size (scaled to fit original canvas width)
    orig_canvas_w = ri - li
    paint_w, paint_h = painting_img.size
    aspect = paint_h / paint_w
    new_canvas_w = orig_canvas_w
    new_canvas_h = int(new_canvas_w * aspect)

    # New frame outer size
    new_frame_w = new_canvas_w + ft_left + ft_right
    new_frame_h = new_canvas_h + ft_top  + ft_bottom

    # Extract shadow from original
    shadow = extract_shadow(arr, outer)
    wall_color = sample_wall_color(arr, outer)

    # Build new frame as black rectangle
    frame_arr = np.zeros((new_frame_h, new_frame_w, 4), dtype=np.uint8)
    # Fill with black (frame)
    frame_arr[:, :] = (0, 0, 0, 255)

    # Fill inner canvas with painting
    painting_resized = painting_img.resize(
        (new_canvas_w, new_canvas_h), Image.LANCZOS
    ).convert("RGBA")
    p_arr = np.array(painting_resized)
    frame_arr[ft_top:ft_top+new_canvas_h, ft_left:ft_left+new_canvas_w] = p_arr

    frame_img = Image.fromarray(frame_arr, "RGBA")

    # Compose onto a wall-colored background
    # Background size = same as original room image
    result = room_img.copy().convert("RGBA")

    # Position: align top-left of new frame to same top-left as original outer frame
    paste_x = lo
    paste_y = to

    # If new frame is smaller/larger, we need to patch the wall color around it
    # First, fill the old frame area with wall color
    result_arr = np.array(result)
    result_arr[to:bo+1, lo:ro+1] = wall_color
    result = Image.fromarray(result_arr, "RGBA")

    # Paste new frame
    result.paste(frame_img, (paste_x, paste_y), frame_img)

    # Re-apply shadow strips (stretched to new frame width/height)
    result_arr2 = np.array(result)

    # Left shadow
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

    # Bottom shadow
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
# API ENDPOINT: POST /mockup
# Body (JSON):
#   room_url    — URL of the room image with frame
#   painting_url — URL of the painting image
#   mode        — "stretch" | "adapt"
# Response (JSON):
#   image_base64 — base64-encoded JPEG of the result
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
# Health check — Render uses this to verify
# the service is alive
# ─────────────────────────────────────────────
@app.route("/health", methods=["GET"])
def health():
    return jsonify({"status": "ok", "service": "artmidnet-mockup"})


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=False)
