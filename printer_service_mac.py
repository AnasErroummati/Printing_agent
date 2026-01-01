#!/usr/bin/env python3
import os
import json
import base64
import io
import logging
import subprocess
import tempfile

from flask import Flask, jsonify, request
from flask_cors import CORS
from PIL import Image, ImageDraw, ImageFont

# ---------------- CONFIG & PATHS ----------------

HTTP_PORT = 8000

# Use Application Support on macOS
HOME = os.path.expanduser("~")
DATA_DIR = os.path.join(HOME, "Library", "Application Support", "PrintAgentService")
os.makedirs(DATA_DIR, exist_ok=True)

LOG_FILE = os.path.join(DATA_DIR, "printer_service.log")
SELECTED_FILE = os.path.join(DATA_DIR, "selected_printer.json")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

# ---------------- PERSISTENCE ----------------


def save_selected(printer_name: str):
    try:
        with open(SELECTED_FILE, "w", encoding="utf-8") as f:
            json.dump({"printer": printer_name}, f)
        logging.info("Saved selected printer: %s", printer_name)
    except Exception:
        logging.exception("Failed to save selected printer")


def load_selected():
    if not os.path.exists(SELECTED_FILE):
        return None
    try:
        with open(SELECTED_FILE, encoding="utf-8") as f:
            return json.load(f).get("printer")
    except Exception:
        logging.exception("Failed to load selected printer")
        return None


# ---------------- CUPS HELPERS ----------------

try:
    import cups

    cups_available = True
    conn = cups.Connection()
except ImportError:
    cups_available = False
    conn = None
    logging.warning("pycups not installed – falling back to lp/lpstat commands only.")


def list_cups_printers():
    """
    Return a list of printer names.
    Uses pycups if available, otherwise uses lpstat.
    """
    if cups_available and conn:
        printers = conn.getPrinters()
        return list(printers.keys())

    # Fallback: parse lpstat
    try:
        out = subprocess.check_output(["lpstat", "-p"], text=True)
        names = []
        for line in out.splitlines():
            if line.startswith("printer "):
                parts = line.split()
                if len(parts) >= 2:
                    names.append(parts[1])
        return names
    except Exception:
        logging.exception("Failed to list printers via lpstat")
        return []


def send_raw_to_printer(printer_name: str, raw_bytes: bytes):
    """
    Send raw data to a printer on macOS.
    NOTE: This is mainly used for /print/drawer, and will only work
    if the CUPS queue is truly RAW/ESC-POS capable.
    """
    try:
        subprocess.run(
            ["lp", "-d", printer_name, "-o", "raw"],
            input=raw_bytes,
            check=True,
        )
    except Exception:
        logging.exception("Failed to send raw data to printer %s", printer_name)
        raise


# ---------------- TEXT+LOGO → RASTER IMAGE ----------------


def render_receipt_to_image(
    plain_text: str, logo_b64: str | None, max_width: int = 576
) -> str:
    """
    Build a single PNG ticket (logo + text) and return the temp file path.

    - plain_text: utf-8 text content of the receipt (multi-line string)
    - logo_b64: 'data:image/png;base64,...' or bare base64, or None
    - max_width: target pixel width (80mm printers often ~576px)
    """

    # 1) Decode logo if present
    logo_img = None
    if logo_b64:
        try:
            if "," in logo_b64:
                logo_b64 = logo_b64.split(",", 1)[1]
            logo_data = base64.b64decode(logo_b64)
            logo_img = Image.open(io.BytesIO(logo_data)).convert("RGBA")

            if logo_img.width > max_width:
                ratio = max_width / logo_img.width
                logo_img = logo_img.resize(
                    (max_width, int(logo_img.height * ratio)), Image.LANCZOS
                )
        except Exception:
            logging.exception("Failed to decode logo image")
            logo_img = None

    # 2) Split text into lines
    lines = (plain_text or "").split("\n")
    # You can tweak line height & paddings to adjust density on the ticket
    line_height = 18
    padding_x = 20
    padding_top = 20
    padding_bottom = 20
    spacing_logo_text = 20

    logo_height = logo_img.height if logo_img else 0
    text_height = line_height * max(len(lines), 1)

    width = max_width
    height = (
        padding_top
        + logo_height
        + (spacing_logo_text if logo_img else 0)
        + text_height
        + padding_bottom
    )

    # 3) Create white canvas
    canvas = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(canvas)

    # Optional: use a monospace font; default font is fine if you don't care
    # font_path = "/System/Library/Fonts/Menlo.ttc"
    # font = ImageFont.truetype(font_path, 12)
    font: ImageFont.FreeTypeFont | ImageFont.ImageFont | None = None

    # 4) Paste logo centered
    current_y = padding_top
    if logo_img:
        x_logo = (width - logo_img.width) // 2
        canvas.paste(logo_img, (x_logo, current_y), logo_img)
        current_y += logo_height + spacing_logo_text

    # 5) Draw text lines
    for line in lines:
        draw.text((padding_x, current_y), line, fill="black", font=font)
        current_y += line_height

    # 6) Save to temp PNG
    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp_path = tmp.name
    tmp.close()
    canvas.save(tmp_path, "PNG")

    logging.info("Rendered receipt image to %s", tmp_path)
    return tmp_path


def print_raster_ticket(printer_name: str, plain_text: str, logo_b64: str | None):
    """
    High-level helper: render (logo + text) to PNG and send to CUPS like Chrome does.
    NO '-o raw' here. Let the driver do its job.
    """
    png_path = render_receipt_to_image(plain_text, logo_b64)
    try:
        subprocess.run(["lp", "-d", printer_name, png_path], check=True)
    finally:
        try:
            os.remove(png_path)
        except Exception:
            logging.warning("Could not delete temp file %s", png_path)


# ---------------- FLASK APP ----------------

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})


@app.route("/printers", methods=["GET"])
def list_printers():
    logging.info("Endpoint /printers called")
    return jsonify(printers=list_cups_printers())


@app.route("/selected", methods=["GET"])
def get_selected():
    return jsonify(selected=load_selected())


@app.route("/select-printer", methods=["POST"])
def select_printer():
    data = request.get_json() or {}
    name = data.get("name")
    if not name:
        return jsonify(error="Printer name required"), 400

    # Optional: validate that printer exists
    available = list_cups_printers()
    if name not in available:
        return jsonify(error=f"Printer '{name}' not found on this Mac"), 404

    save_selected(name)
    return jsonify(selected=name)


@app.route("/status", methods=["GET"])
def status():
    printer = load_selected()
    connected = False
    if printer:
        # Simple check: see if it shows in list
        connected = printer in list_cups_printers()
    return jsonify(selected=printer, connected=connected)


@app.route("/print/drawer", methods=["POST"])
def open_drawer():
    """
    Attempts ESC/POS drawer kick via raw.
    This will only work if you later create a real RAW ESC/POS queue.
    """
    printer = load_selected()
    if not printer:
        return jsonify(error="No printer selected"), 404

    try:
        drawer_cmd = b"\x1b\x70\x00\x19\xfa"
        send_raw_to_printer(printer, drawer_cmd)
        return jsonify(result="drawer opened")
    except Exception:
        logging.exception("Drawer open failed")
        return jsonify(error="Failed to open drawer"), 500


@app.route("/print/test", methods=["POST"])
def test_print():
    """
    Simple test: print a line of text using normal CUPS path (non-raw).
    """
    printer = load_selected()
    if not printer:
        return jsonify(error="No printer selected"), 404

    try:
        text = "Success from macOS\n"
        subprocess.run(
            ["lp", "-d", printer],
            input=text.encode("utf-8"),
            check=True,
        )
        return jsonify(result="test printed")
    except Exception:
        logging.exception("Test print failed")
        return jsonify(error="Failed to test print"), 500


@app.route("/print/raw", methods=["POST"])
def print_raw_bytes():
    """
    MAC VERSION of /print/raw

    Payload shape (backward-compatible):

      {
        "data": "<base64 of ESC/POS or plain text bytes>",   # legacy / Windows field
        "logo": "<base64 image>",
        "printLogo": true/false,
        "plainTextReceipt": "<optional raw text for mac>"    # NEW, from TS generateReceiptPayload
      }

    On macOS we:
      - Prefer 'plainTextReceipt' if provided.
      - Otherwise, try to decode 'data' as UTF-8 text.
      - Then render (logo + text) into a PNG and print via CUPS (no -o raw).
    """
    print("\n--- NEW PRINT REQUEST (macOS raster) ---")
    printer = load_selected()
    if not printer:
        return jsonify(error="No printer selected"), 404

    req_data = request.get_json() or {}

    # NEW field from TS generator
    plain_text = req_data.get("plainTextReceipt")

    # Legacy field (still sent by old code / used by Windows)
    b64_text = req_data.get("data")

    b64_logo = req_data.get("logo")
    should_print_logo = req_data.get("printLogo", False)

    # 1) Determine the text we will render
    if not plain_text:
        # No explicit plainTextReceipt → fallback to decoding 'data'
        if not b64_text:
            return jsonify(error="No data provided"), 400

        try:
            raw_bytes = base64.b64decode(b64_text)
            plain_text = raw_bytes.decode("utf-8", errors="ignore")
        except Exception:
            plain_text = ""

    if plain_text is None or plain_text.strip() == "":
        return jsonify(error="Empty receipt text"), 400

    # 2) Decide whether to use logo
    logo_for_render = b64_logo if (should_print_logo and b64_logo) else None

    try:
        logging.info(
            "Printing raster ticket: %d chars, logo=%s",
            len(plain_text),
            bool(logo_for_render),
        )

        print_raster_ticket(printer, plain_text, logo_for_render)

        print("--- SUCCESS (macOS raster) ---\n")
        return jsonify(result="success")
    except Exception as e:
        logging.exception("Critical print error on macOS")
        return jsonify(error=str(e)), 500


@app.route("/initialize", methods=["POST"])
def initialize():
    try:
        if os.path.exists(SELECTED_FILE):
            os.remove(SELECTED_FILE)
        return jsonify(result="initialized")
    except Exception:
        logging.exception("Initialization failed")
        return jsonify(error="Failed to initialize"), 500


if __name__ == "__main__":
    print(f"Print agent running on http://localhost:{HTTP_PORT}")
    app.run(host="0.0.0.0", port=HTTP_PORT, use_reloader=False)
