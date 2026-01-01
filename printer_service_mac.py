#!/usr/bin/env python3
import os
import json
import base64
import io
import logging
import subprocess
import tempfile

from typing import Optional
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

# ---- LOGGING SETUP: file + console ----
logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
)

logger = logging.getLogger()
logger.setLevel(logging.INFO)

# Add console logging too (so you see everything in Terminal)
console_handler = logging.StreamHandler()
console_handler.setLevel(logging.INFO)
console_handler.setFormatter(logging.Formatter("%(asctime)s [%(levelname)s] %(message)s"))
logger.addHandler(console_handler)

logger.info("=== Starting PrintAgentService (macOS) ===")

# ---------------- PERSISTENCE ----------------


def save_selected(printer_name: str):
    try:
        with open(SELECTED_FILE, "w", encoding="utf-8") as f:
            json.dump({"printer": printer_name}, f)
        logger.info("Saved selected printer: %s", printer_name)
    except Exception:
        logger.exception("Failed to save selected printer")


def load_selected():
    if not os.path.exists(SELECTED_FILE):
        logger.info("No selected printer file found yet")
        return None
    try:
        with open(SELECTED_FILE, encoding="utf-8") as f:
            printer = json.load(f).get("printer")
            logger.info("Loaded selected printer from disk: %s", printer)
            return printer
    except Exception:
        logger.exception("Failed to load selected printer")
        return None


# ---------------- CUPS HELPERS ----------------

try:
    import cups

    cups_available = True
    conn = cups.Connection()
    logger.info("pycups available, using CUPS API")
except ImportError:
    cups_available = False
    conn = None
    logger.warning("pycups NOT installed – falling back to lp/lpstat commands only.")


def list_cups_printers():
    if cups_available and conn:
        try:
            printers = conn.getPrinters()
            names = list(printers.keys())
            logger.info("CUPS printers via pycups: %s", names)
            return names
        except Exception:
            logger.exception("Error while listing printers via pycups")

    # Fallback: parse lpstat
    try:
        out = subprocess.check_output(["lpstat", "-p"], text=True)
        names = []
        for line in out.splitlines():
            if line.startswith("printer "):
                parts = line.split()
                if len(parts) >= 2:
                    names.append(parts[1])
        logger.info("CUPS printers via lpstat: %s", names)
        return names
    except Exception:
        logger.exception("Failed to list printers via lpstat")
        return []


def send_raw_to_printer(printer_name: str, raw_bytes: bytes):
    """
    Raw ESC/POS send. Only works if printer queue is truly RAW.
    Used for /print/drawer.
    """
    logger.info(
        "send_raw_to_printer → printer=%s, bytes=%d",
        printer_name,
        len(raw_bytes),
    )
    try:
        result = subprocess.run(
            ["lp", "-d", printer_name, "-o", "raw"],
            input=raw_bytes,
            capture_output=True,
            check=False,
        )
        logger.info(
            "lp (raw) returncode=%s stdout=%s stderr=%s",
            result.returncode,
            result.stdout.decode("utf-8", errors="ignore")
            if isinstance(result.stdout, (bytes, bytearray))
            else result.stdout,
            result.stderr.decode("utf-8", errors="ignore")
            if isinstance(result.stderr, (bytes, bytearray))
            else result.stderr,
        )
        if result.returncode != 0:
            raise RuntimeError(f"lp raw failed with code {result.returncode}")
    except Exception:
        logger.exception("Failed to send raw data to printer %s", printer_name)
        raise


# ---------------- TEXT+LOGO → RASTER IMAGE ----------------


def render_receipt_to_image(
    plain_text: str, logo_b64: Optional[str], max_width: int = 576
) -> str:
    """
    Build a single PNG ticket (logo + text) and return the temp file path.
    """

    logger.info(
        "render_receipt_to_image → text_len=%d, logo_present=%s, max_width=%d",
        len(plain_text or ""),
        bool(logo_b64),
        max_width,
    )

    logo_img = None
    if logo_b64:
        try:
            logger.info("Decoding logo base64 (first 60 chars): %s...", logo_b64[:60])
            if "," in logo_b64:
                logo_b64 = logo_b64.split(",", 1)[1]
            logo_data = base64.b64decode(logo_b64)
            logo_img = Image.open(io.BytesIO(logo_data)).convert("RGBA")
            logger.info("Logo image decoded: size=%sx%s", logo_img.width, logo_img.height)

            if logo_img.width > max_width:
                ratio = max_width / logo_img.width
                new_size = (max_width, int(logo_img.height * ratio))
                logo_img = logo_img.resize(new_size, Image.LANCZOS)
                logger.info("Logo resized to: %sx%s", logo_img.width, logo_img.height)
        except Exception:
            logger.exception("Failed to decode logo image")
            logo_img = None

    lines = (plain_text or "").split("\n")
    logger.info("Number of text lines: %d", len(lines))

    line_height = 18      # tweak if receipt is too tall/short
    padding_x = 20        # left padding
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

    logger.info("Creating canvas: width=%d, height=%d", width, height)
    canvas = Image.new("RGB", (width, height), "white")
    draw = ImageDraw.Draw(canvas)

    font = None  # default font

    current_y = padding_top
    if logo_img:
        x_logo = (width - logo_img.width) // 2
        canvas.paste(logo_img, (x_logo, current_y), logo_img)
        logger.info(
            "Pasted logo at x=%d, y=%d (logo size=%dx%d)",
            x_logo,
            current_y,
            logo_img.width,
            logo_img.height,
        )
        current_y += logo_height + spacing_logo_text

    for idx, line in enumerate(lines):
        draw.text((padding_x, current_y), line, fill="black", font=font)
        logger.debug("Drawing line %d at y=%d: %s", idx, current_y, line)
        current_y += line_height

    tmp = tempfile.NamedTemporaryFile(suffix=".png", delete=False)
    tmp_path = tmp.name
    tmp.close()
    canvas.save(tmp_path, "PNG")

    logger.info("Rendered receipt image to %s", tmp_path)
    return tmp_path


def print_raster_ticket(printer_name: str, plain_text: str, logo_b64: Optional[str]):
    """
    Render (logo + text) to PNG and print via lp (non-raw).
    """
    logger.info(
        "print_raster_ticket → printer=%s, text_len=%d, logo_present=%s",
        printer_name,
        len(plain_text or ""),
        bool(logo_b64),
    )

    png_path = render_receipt_to_image(plain_text, logo_b64)
    logger.info("About to send PNG to lp: %s", png_path)

    try:
        result = subprocess.run(
            ["lp", "-d", printer_name, png_path],
            capture_output=True,
            check=False,
        )
        stdout = (
            result.stdout.decode("utf-8", errors="ignore")
            if isinstance(result.stdout, (bytes, bytearray))
            else result.stdout
        )
        stderr = (
            result.stderr.decode("utf-8", errors="ignore")
            if isinstance(result.stderr, (bytes, bytearray))
            else result.stderr
        )
        logger.info(
            "lp (image) returncode=%s stdout=%s stderr=%s",
            result.returncode,
            stdout,
            stderr,
        )
        if result.returncode != 0:
            raise RuntimeError(f"lp image failed with code {result.returncode}")
    finally:
        # For debugging, we KEEP the PNG (so you can open it and see).
        # If you want to delete it later, uncomment the remove.
        # try:
        #     os.remove(png_path)
        #     logger.info("Deleted temp PNG %s", png_path)
        # except Exception:
        #     logger.warning("Could not delete temp file %s", png_path)
        logger.info("Keeping temp PNG at %s for inspection", png_path)


# ---------------- FLASK APP ----------------

app = Flask(__name__)
CORS(app, resources={r"/*": {"origins": "*"}})


@app.route("/printers", methods=["GET"])
def list_printers_route():
    logger.info("Endpoint /printers called")
    printers = list_cups_printers()
    return jsonify(printers=printers)


@app.route("/selected", methods=["GET"])
def get_selected():
    printer = load_selected()
    logger.info("GET /selected → %s", printer)
    return jsonify(selected=printer)


@app.route("/select-printer", methods=["POST"])
def select_printer():
    data = request.get_json() or {}
    name = data.get("name")
    logger.info("POST /select-printer with name=%s", name)

    if not name:
        return jsonify(error="Printer name required"), 400

    available = list_cups_printers()
    if name not in available:
        logger.warning(
            "Requested printer '%s' not in available list: %s", name, available
        )
        return jsonify(error=f"Printer '{name}' not found on this Mac"), 404

    save_selected(name)
    return jsonify(selected=name)


@app.route("/status", methods=["GET"])
def status():
    printer = load_selected()
    available = list_cups_printers()
    connected = printer in available if printer else False
    logger.info(
        "GET /status → selected=%s, connected=%s, available=%s",
        printer,
        connected,
        available,
    )
    return jsonify(selected=printer, connected=connected)


@app.route("/print/drawer", methods=["POST"])
def open_drawer():
    logger.info("POST /print/drawer")
    printer = load_selected()
    if not printer:
        logger.warning("Drawer requested but no printer selected")
        return jsonify(error="No printer selected"), 404

    try:
        drawer_cmd = b"\x1b\x70\x00\x19\xfa"
        send_raw_to_printer(printer, drawer_cmd)
        logger.info("Drawer command sent successfully")
        return jsonify(result="drawer opened")
    except Exception:
        logger.exception("Drawer open failed")
        return jsonify(error="Failed to open drawer"), 500


@app.route("/print/test", methods=["POST"])
def test_print():
    logger.info("POST /print/test")
    printer = load_selected()
    if not printer:
        logger.warning("Test print requested but no printer selected")
        return jsonify(error="No printer selected"), 404

    try:
        text = "Success from macOS\n"
        logger.info(
            "Sending simple text test to lp: printer=%s, text=%r", printer, text
        )
        result = subprocess.run(
            ["lp", "-d", printer],
            input=text.encode("utf-8"),
            capture_output=True,
            check=False,
        )
        stdout = (
            result.stdout.decode("utf-8", errors="ignore")
            if isinstance(result.stdout, (bytes, bytearray))
            else result.stdout
        )
        stderr = (
            result.stderr.decode("utf-8", errors="ignore")
            if isinstance(result.stderr, (bytes, bytearray))
            else result.stderr
        )
        logger.info(
            "lp (test) returncode=%s stdout=%s stderr=%s",
            result.returncode,
            stdout,
            stderr,
        )
        if result.returncode != 0:
            raise RuntimeError(f"lp test failed with code {result.returncode}")
        return jsonify(result="test printed")
    except Exception:
        logger.exception("Test print failed")
        return jsonify(error="Failed to test print"), 500


@app.route("/print/raw", methods=["POST"])
def print_raw_bytes():
    logger.info("POST /print/raw → NEW PRINT REQUEST (macOS raster)")
    printer = load_selected()
    if not printer:
        logger.warning("Print requested but no printer selected")
        return jsonify(error="No printer selected"), 404

    req_data = request.get_json() or {}
    logger.info("Incoming JSON keys: %s", list(req_data.keys()))

    plain_text = req_data.get("plainTextReceipt")
    b64_text = req_data.get("data")
    b64_logo = req_data.get("logo")
    should_print_logo = req_data.get("printLogo", False)

    logger.info(
        "Payload summary: plainText_len=%s, has_data=%s, has_logo=%s, printLogo=%s",
        len(plain_text) if plain_text else None,
        b64_text is not None,
        b64_logo is not None,
        should_print_logo,
    )

    # Determine text to render
    if not plain_text:
        logger.info("No plainTextReceipt provided, falling back to base64 'data'")
        if not b64_text:
            logger.warning("No data provided in /print/raw")
            return jsonify(error="No data provided"), 400
        try:
            raw_bytes = base64.b64decode(b64_text)
            plain_text = raw_bytes.decode("utf-8", errors="ignore")
            logger.info(
                "Decoded plain_text from base64 'data', length=%d",
                len(plain_text or ""),
            )
        except Exception:
            logger.exception("Failed to decode base64 'data' as UTF-8 text")
            plain_text = ""

    if plain_text is None or plain_text.strip() == "":
        logger.warning("Final plain_text is empty after decoding")
        return jsonify(error="Empty receipt text"), 400

    if len(plain_text) > 300:
        logger.info("First 300 chars of plain_text:\n%s", plain_text[:300])
    else:
        logger.info("plain_text:\n%s", plain_text)

    logo_for_render = b64_logo if (should_print_logo and b64_logo) else None
    logger.info("Logo will be used: %s", bool(logo_for_render))

    try:
        logger.info("Calling print_raster_ticket...")
        print_raster_ticket(printer, plain_text, logo_for_render)
        logger.info("--- SUCCESS (macOS raster) ---")
        return jsonify(result="success")
    except Exception as e:
        logger.exception("Critical print error on macOS in /print/raw")
        return jsonify(error=str(e)), 500


@app.route("/initialize", methods=["POST"])
def initialize():
    logger.info("POST /initialize")
    try:
        if os.path.exists(SELECTED_FILE):
            os.remove(SELECTED_FILE)
            logger.info("Deleted selected_printer.json")
        return jsonify(result="initialized")
    except Exception:
        logger.exception("Initialization failed")
        return jsonify(error="Failed to initialize"), 500


if __name__ == "__main__":
    logger.info(f"Print agent running on http://localhost:{HTTP_PORT}")
    print(f"Print agent running on http://localhost:{HTTP_PORT}")
    app.run(host="0.0.0.0", port=HTTP_PORT, use_reloader=False)
