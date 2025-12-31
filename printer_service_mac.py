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


def print_file_via_cups(printer_name: str, file_path: str):
    """
    Print a file (PNG, PDF, etc.) using normal CUPS pipeline.
    No -o raw, so the driver renders it (like Chrome).
    """
    try:
        subprocess.run(
            ["lp", "-d", printer_name, file_path],
            check=True
        )
    except Exception:
        logging.exception(f"Failed to print file via CUPS on printer {printer_name}")
        raise


# ---------------- RASTER RECEIPT RENDERING ----------------

def print_receipt_raster(printer_name: str, plain_text: str):
    """
    Render a receipt image:

        Reçu Client  (centered header)
        <plain_text lines...>

    Then print it via CUPS as a PNG.
    """
    MAX_WIDTH = 576  # typical width for 80mm thermal printers (adjust if needed)
    padding = 20
    line_height = 18

    # Split text into lines
    lines = plain_text.split("\n")

    # Header
    header_text = "Reçu Client"
    header_height = 40

    # Total height = padding + header + padding + text + padding
    text_height = len(lines) * line_height
    height = padding + header_height + padding + text_height + padding

    # Create white canvas
    canvas = Image.new("RGB", (MAX_WIDTH, height), "white")
    draw = ImageDraw.Draw(canvas)

    # Try to use Menlo (monospace) if available
    try:
        header_font = ImageFont.truetype("/System/Library/Fonts/Menlo.ttc", 18)
        body_font = ImageFont.truetype("/System/Library/Fonts/Menlo.ttc", 12)
    except Exception:
        header_font = None
        body_font = None

    # Draw centered header
    w, _ = draw.textsize(header_text, font=header_font)
    draw.text(
        ((MAX_WIDTH - w) // 2, padding),
        header_text,
        fill="black",
        font=header_font
    )

    # Draw receipt text below header
    y = padding + header_height
    for line in lines:
        draw.text((20, y), line, fill="black", font=body_font)
        y += line_height

    # Save to temporary PNG and send to CUPS
    with tempfile.NamedTemporaryFile(suffix=".png", delete=False) as tmp:
        canvas.save(tmp.name, "PNG")
        tmp_path = tmp.name

    print_file_via_cups(printer_name, tmp_path)


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
        connected = printer in list_cups_printers()
    return jsonify(selected=printer, connected=connected)


@app.route("/print/test", methods=["POST"])
def test_print():
    """
    Simple test endpoint:
    prints a small test receipt with 'Reçu Client' + 2 lines.
    """
    printer = load_selected()
    if not printer:
        return jsonify(error="No printer selected"), 404

    try:
        test_text = "Test ligne 1\nTest ligne 2\nMerci d'utiliser Ramti 😉"
        print_receipt_raster(printer, test_text)
        return jsonify(result="test printed")
    except Exception:
        logging.exception("Test print failed")
        return jsonify(error="Failed to test print"), 500


@app.route("/print/raw", methods=["POST"])
def print_raw():
    """
    macOS-specific print endpoint.

    Expects JSON:
    {
      "plainTextReceipt": "Ramti Salon\nItem A ...\nItem B ...\nTotal: 120 DH"
    }

    It will print:

        Reçu Client
        Ramti Salon
        Item A ...
        ...
    """
    printer = load_selected()
    if not printer:
        return jsonify(error="No printer selected"), 404

    data = request.get_json() or {}
    plain_text = data.get("plainTextReceipt")

    if not plain_text:
        return jsonify(error="plainTextReceipt is required on macOS"), 400

    try:
        print_receipt_raster(printer, plain_text)
        return jsonify(result="success")
    except Exception as e:
        logging.exception("Mac raster print failed")
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
