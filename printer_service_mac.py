#!/usr/bin/env python3
import os
import json
import base64
import io
import logging
import subprocess

from flask import Flask, jsonify, request
from flask_cors import CORS
from PIL import Image

# ---------------- CONFIG & PATHS ----------------

HTTP_PORT = 8000

# Use Application Support on macOS
HOME = os.path.expanduser("~")
DATA_DIR = os.path.join(HOME, "Library", "Application Support", "PrintAgentService")
os.makedirs(DATA_DIR, exist_ok=True)

LOG_FILE      = os.path.join(DATA_DIR, "printer_service.log")
SELECTED_FILE = os.path.join(DATA_DIR, "selected_printer.json")

logging.basicConfig(
    filename=LOG_FILE,
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s"
)

# ---------------- PERSISTENCE ----------------

def save_selected(printer_name: str):
    try:
        with open(SELECTED_FILE, "w", encoding="utf-8") as f:
            json.dump({"printer": printer_name}, f)
        logging.info(f"Saved selected printer: {printer_name}")
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
    Send raw ESC/POS bytes to a printer on macOS.
    Uses 'lp -o raw'.
    """
    try:
        proc = subprocess.run(
            ["lp", "-d", printer_name, "-o", "raw"],
            input=raw_bytes,
            check=True
        )
    except Exception as e:
        logging.exception(f"Failed to send raw data to printer {printer_name}")
        raise


# ---------------- IMAGE → ESC/POS ----------------

def convert_image_to_escpos(base64_string: str) -> bytes:
    """
    Same logic as your Windows version, adapted 1:1.
    """
    print(f"DEBUG: Logo Input Length: {len(base64_string)}")
    print(f"DEBUG: Logo Start: {base64_string[:50]}...")

    try:
        # Remove header if present
        if "," in base64_string:
            base64_string = base64_string.split(",", 1)[1]

        image_data = base64.b64decode(base64_string)
        im = Image.open(io.BytesIO(image_data))

        # Resize to ~384px width (typical 58mm thermal)
        MAX_WIDTH = 384
        if im.width > MAX_WIDTH:
            ratio = MAX_WIDTH / float(im.width)
            new_height = int(im.height * ratio)
            if hasattr(Image, "Resampling"):
                resample_method = Image.Resampling.LANCZOS
            else:
                resample_method = Image.LANCZOS
            im = im.resize((MAX_WIDTH, new_height), resample_method)

        # Handle transparency -> white background
        if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
            bg = Image.new("RGB", im.size, (255, 255, 255))
            im = im.convert("RGBA")
            alpha = im.split()[3]
            bg.paste(im, mask=alpha)
            im = bg
        else:
            im = im.convert("RGB")

        # Grayscale
        im = im.convert("L")

        # Threshold
        threshold = 190
        im = im.point(lambda p: 255 if p > threshold else 0)
        im = im.convert("1")  # 1-bit

        width_bytes = (im.width + 7) // 8
        header = (
            b"\x1d\x76\x30\x00"
            + width_bytes.to_bytes(2, "little")
            + im.height.to_bytes(2, "little")
        )

        img_bytes = bytearray(im.tobytes())
        for i in range(len(img_bytes)):
            img_bytes[i] ^= 0xFF  # invert bits

        return header + bytes(img_bytes)

    except Exception as e:
        print(f"Logo Conversion Error: {e}")
        return b""


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
    printer = load_selected()
    if not printer:
        return jsonify(error="No printer selected"), 404

    try:
        # ESC/POS drawer kick
        drawer_cmd = b"\x1b\x70\x00\x19\xfa"
        send_raw_to_printer(printer, drawer_cmd)
        return jsonify(result="drawer opened")
    except Exception:
        logging.exception("Drawer open failed")
        return jsonify(error="Failed to open drawer"), 500


@app.route("/print/test", methods=["POST"])
def test_print():
    printer = load_selected()
    if not printer:
        return jsonify(error="No printer selected"), 404

    try:
        data = b"Success from macOS\n"
        send_raw_to_printer(printer, data)
        return jsonify(result="test printed")
    except Exception:
        logging.exception("Test print failed")
        return jsonify(error="Failed to test print"), 500


@app.route("/print/raw", methods=["POST"])
def print_raw_bytes():
    print("\n--- NEW PRINT REQUEST (macOS) ---")
    printer = load_selected()
    if not printer:
        return jsonify(error="No printer selected"), 404

    req_data = request.get_json() or {}
    b64_text = req_data.get("data")        # Receipt text (base64)
    b64_logo = req_data.get("logo")        # Logo image (base64)
    should_print_logo = req_data.get("printLogo", False)

    if not b64_text:
        return jsonify(error="No data provided"), 400

    try:
        print("--> Decoding Text...")
        receipt_bytes = base64.b64decode(b64_text)

        drawer_cmd = b"\x1b\x70\x00\x19\xfa"
        chunks = []

        # Kick drawer
        print("--> ACTION: Kick Drawer")
        chunks.append(drawer_cmd)

        # Logo
        if should_print_logo and b64_logo:
            print("--> ACTION: Processing Logo...")
            logo_bytes = convert_image_to_escpos(b64_logo)
            if logo_bytes:
                # Center align
                chunks.append(b"\x1b\x61\x01")
                chunks.append(logo_bytes)
                # Back to left
                chunks.append(b"\x1b\x61\x00")
                print("    Logo bytes prepared.")
            else:
                print("    Logo conversion failed or empty.")

        # Receipt text
        print("--> ACTION: Print Text")
        chunks.append(receipt_bytes)

        # Concatenate and send once
        final_bytes = b"".join(chunks)
        send_raw_to_printer(printer, final_bytes)

        print("--- SUCCESS (macOS) ---\n")
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
