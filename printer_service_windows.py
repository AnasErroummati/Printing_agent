# printer_service.py
"""
Windows Service wrapping a Flask-based print agent.
"""
import os
import json
import logging
import sys
import threading
import concurrent.futures
import time
import servicemanager
import win32serviceutil
import base64
from PIL import Image, ImageWin
import win32print
import win32ui
import win32con
from PIL import Image
import io
# Configuration
HTTP_PORT           = 8000
SERVICE_NAME        = "PrintAgentService"
SERVICE_DISPLAY_NAME= "Python Print Agent"

# Persist data under ProgramData so it survives reboots
import os
PROGRAM_DATA = os.environ.get("PROGRAMDATA", r"C:\ProgramData")
DATA_DIR     = os.path.join(PROGRAM_DATA, "PrintAgentService")
os.makedirs(DATA_DIR, exist_ok=True)

LOG_FILE      = os.path.join(DATA_DIR, "printer_service.log")
SELECTED_FILE = os.path.join(DATA_DIR, "selected_printer.json")
# Persistence helpers
def save_selected(printer_name):
    try:
        with open(SELECTED_FILE, 'w') as f:
            json.dump({'printer': printer_name}, f)
        logging.info(f"Saved selected printer: {printer_name}")
    except Exception:
        logging.exception("Failed to save selected printer")


def load_selected():
    if not os.path.exists(SELECTED_FILE):
        return None
    try:
        with open(SELECTED_FILE) as f:
            return json.load(f).get('printer')
    except Exception:
        logging.exception("Failed to load selected printer")
        return None

# Flask app launcher
def run_service():
    from flask import Flask, jsonify, request
    try:
        from flask_cors import CORS
        cors_available = True
    except ImportError:
        cors_available = False
    import win32print

    app = Flask(__name__)
    if cors_available:
        CORS(app, resources={r"/*": {"origins": "*"}})
    else:
        @app.after_request
        def _add_cors_headers(response):
            response.headers.update({
                'Access-Control-Allow-Origin': '*',
                'Access-Control-Allow-Methods': 'GET,POST,OPTIONS',
                'Access-Control-Allow-Headers': 'Content-Type'
            })
            return response

    @app.route('/printers', methods=['GET'])
    def list_printers():
        logging.info("Endpoint /printers called")
        try:
            flags = win32print.PRINTER_ENUM_LOCAL | win32print.PRINTER_ENUM_CONNECTIONS
            names = [name for _, _, name, _ in win32print.EnumPrinters(flags)]
            return jsonify(printers=names)
        except Exception:
            logging.exception("Failed to enumerate printers")
            return jsonify(error="Failed to enumerate printers"), 500

    @app.route('/selected', methods=['GET'])
    def get_selected():
        return jsonify(selected=load_selected())

    @app.route('/select-printer', methods=['POST'])
    def select_printer():
        data = request.get_json() or {}
        name = data.get('name')
        if not name:
            return jsonify(error="Printer name required"), 400
        save_selected(name)
        return jsonify(selected=name)

    @app.route('/status', methods=['GET'])
    def status():
        printer = load_selected()
        connected = False
        if printer:
            try:
                ph = win32print.OpenPrinter(printer)
                win32print.ClosePrinter(ph)
                connected = True
            except:
                connected = False
        return jsonify(selected=printer, connected=connected)

    @app.route('/print/drawer', methods=['POST'])
    def open_drawer():
        printer = load_selected()
        if not printer:
            return jsonify(error="No printer selected"), 404
        try:
            cmd = b"\x1b\x70\x00\x19\xfa"
            ph = win32print.OpenPrinter(printer)
            win32print.StartDocPrinter(ph, 1, ("Drawer", None, "RAW"))
            win32print.StartPagePrinter(ph)
            win32print.WritePrinter(ph, cmd)
            win32print.EndPagePrinter(ph)
            win32print.EndDocPrinter(ph)
            win32print.ClosePrinter(ph)
            return jsonify(result="drawer opened")
        except Exception:
            logging.exception("Drawer open failed")
            return jsonify(error="Failed to open drawer"), 500

    @app.route('/print/test', methods=['POST'])
    def test_print():
        printer = load_selected()
        if not printer:
            return jsonify(error="No printer selected"), 404
        try:
            data = b"Success\n"
            ph = win32print.OpenPrinter(printer)
            win32print.StartDocPrinter(ph, 1, ("Test", None, "RAW"))
            win32print.StartPagePrinter(ph)
            win32print.WritePrinter(ph, data)
            win32print.EndPagePrinter(ph)
            win32print.EndDocPrinter(ph)
            win32print.ClosePrinter(ph)
            return jsonify(result="test printed")
        except Exception:
            logging.exception("Test print failed")
            return jsonify(error="Failed to test print"), 500
    # --- HELPER: CONVERT IMAGE TO ESC/POS BITMAP ---

    def convert_image_to_escpos(base64_string: str) -> bytes:
        """
        Converts a Base64 image string into raw ESC/POS 'GS v 0' commands.
        Keeps background white and makes logo/text black.
        """
        print(f"DEBUG: Logo Input Length: {len(base64_string)}")
        print(f"DEBUG: Logo Start: {base64_string[:50]}...")  # Check if looks like base64

        try:
            # 1. Remove header if present: data:image/png;base64,...
            if "," in base64_string:
                base64_string = base64_string.split(",", 1)[1]

            # 2. Decode and open
            image_data = base64.b64decode(base64_string)
            im = Image.open(io.BytesIO(image_data))

            # 3. Resize for standard thermal printer
            MAX_WIDTH = 384
            if im.width > MAX_WIDTH:
                ratio = MAX_WIDTH / float(im.width)
                new_height = int(im.height * ratio)
                if hasattr(Image, "Resampling"):
                    resample_method = Image.Resampling.LANCZOS
                else:
                    resample_method = Image.LANCZOS
                im = im.resize((MAX_WIDTH, new_height), resample_method)

            # 4. Handle transparency → paste onto WHITE background
            if im.mode in ("RGBA", "LA") or (im.mode == "P" and "transparency" in im.info):
                bg = Image.new("RGB", im.size, (255, 255, 255))
                im = im.convert("RGBA")
                alpha = im.split()[3]
                bg.paste(im, mask=alpha)
                im = bg
            else:
                im = im.convert("RGB")

            # 5. Grayscale
            im = im.convert("L")

            # 6. Threshold to make logo darker (black) & background white
            threshold = 190  # tweak 180–200 if needed
            im = im.point(lambda p: 255 if p > threshold else 0)
            im = im.convert("1")  # 1-bit: 0=black, 255=white (in Pillow)

            # 7. Build ESC/POS header for GS v 0
            width_bytes = (im.width + 7) // 8
            header = (
                b"\x1d\x76\x30\x00"
                + width_bytes.to_bytes(2, "little")
                + im.height.to_bytes(2, "little")
            )

            # 8. INVERT BITS so ESC/POS prints black where Pillow had black
            # Pillow "1" mode: bit=0 black, 1 white
            # ESC/POS raster: bit=1 black dot, 0 white
            img_bytes = bytearray(im.tobytes())
            for i in range(len(img_bytes)):
                img_bytes[i] ^= 0xFF  # flip all bits

            return header + bytes(img_bytes)

        except Exception as e:
            print(f"Logo Conversion Error: {e}")
            return b""
    # --- MAIN PRINT ENDPOINT ---
    @app.route('/print/raw', methods=['POST'])
    def print_raw_bytes():
        print("\n--- NEW PRINT REQUEST ---")
        
        printer = load_selected()
        if not printer:
            return jsonify(error="No printer selected"), 404
        
        req_data = request.get_json() or {}
        
        # 1. Extract Data
        b64_text = req_data.get('data')             # The Receipt Text
        b64_logo = req_data.get('logo')             # The Logo Image
        should_print_logo = req_data.get('printLogo', False) # Boolean Flag
        
        if not b64_text:
            return jsonify(error="No data provided"), 400

        try:
            # 2. Prepare Bytes
            print("--> Decoding Text...")
            receipt_bytes = base64.b64decode(b64_text)
            
            drawer_cmd = b"\x1b\x70\x00\x19\xfa"

            # 3. Open Printer
            print(f"--> Opening Printer: {printer}")
            ph = win32print.OpenPrinter(printer)
            
            try:
                win32print.StartDocPrinter(ph, 1, ("React_Receipt", None, "RAW"))
                win32print.StartPagePrinter(ph)
                
                # --- SEQUENCE START ---
                
                # A: KICK DRAWER (First priority)
                print("--> ACTION: Kick Drawer")
                win32print.WritePrinter(ph, drawer_cmd)
                
                # B: PRINT LOGO (If requested and available)
                if should_print_logo and b64_logo:
                    print("--> ACTION: Processing Logo...")
                    logo_bytes = convert_image_to_escpos(b64_logo)
                    
                    if logo_bytes:
                        # 1. Center Align
                        win32print.WritePrinter(ph, b"\x1b\x61\x01")
                        # 2. Print Image
                        win32print.WritePrinter(ph, logo_bytes)
                        # 3. Reset to Left Align (Crucial for text)
                        win32print.WritePrinter(ph, b"\x1b\x61\x00")
                        print("    Logo sent to printer.")
                    else:
                        print("    Logo conversion failed or empty.")
                # C: PRINT TEXT (The receipt content)
                print("--> ACTION: Print Text")
                win32print.WritePrinter(ph, receipt_bytes)
                
                # --- SEQUENCE END ---

                win32print.EndPagePrinter(ph)
            finally:
                win32print.EndDocPrinter(ph)
                win32print.ClosePrinter(ph)
            
            print("--- SUCCESS ---\n")
            return jsonify(result="success")
            
        except Exception as e:
            print(f"!!! CRITICAL ERROR !!! {str(e)}")
            import traceback
            traceback.print_exc()
            return jsonify(error=str(e)), 500

    logging.info(f"Starting Flask on port {HTTP_PORT}")
    app.run(host='0.0.0.0', port=HTTP_PORT, use_reloader=False)

    @app.route('/initialize', methods=['POST'])
    def initialize():
        try:
            if os.path.exists(SELECTED_FILE):
                os.remove(SELECTED_FILE)
            return jsonify(result="initialized")
        except Exception:
            logging.exception("Initialization failed")
            return jsonify(error="Failed to initialize"), 500

    logging.info(f"Starting Flask on port {HTTP_PORT}")
    app.run(host='0.0.0.0', port=HTTP_PORT, use_reloader=False)

# Worker thread to host Flask
class ServiceThread(threading.Thread):
    def __init__(self, stop_event):
        super().__init__()
        self.stop_event = stop_event

    def run(self):
        # Start Flask in background
        executor = concurrent.futures.ThreadPoolExecutor(max_workers=1)
        executor.submit(run_service)
        # Keep alive until stopped
        while not self.stop_event.is_set():
            time.sleep(1)

class FlaskService(win32serviceutil.ServiceFramework):
    _svc_name_ = SERVICE_NAME
    _svc_display_name_ = SERVICE_DISPLAY_NAME
    _svc_description_ = "Python-based printing agent service"

    def __init__(self, args):
        super().__init__(args)
        self.stop_event = threading.Event()
        self.thread = ServiceThread(self.stop_event)

    def SvcStop(self):
        logging.info("Service stop requested")
        self.stop_event.set()

    def SvcDoRun(self):
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(FlaskService)
        self.thread.start()
        self.stop_event.wait()
        self.thread.join()

if __name__ == '__main__':
    # --- NEW DEBUG LOGIC ---
    if len(sys.argv) > 1 and sys.argv[1] == 'debug':
        # 1. Configure logging to show up in the TERMINAL (stdout)
        #    so you don't have to dig into the log file.
        root = logging.getLogger()
        root.addHandler(logging.StreamHandler(sys.stdout))
        root.setLevel(logging.INFO)
        
        print("----------------------------------------------------------------")
        print("  DEBUG MODE ACTIVE")
        print("  - Running as a normal script (Not a Service)")
        print(f"  - API accessible at http://localhost:{HTTP_PORT}")
        print("  - Press CTRL+C to stop")
        print("----------------------------------------------------------------")
        
        # 2. Run the Flask App directly
        run_service()
        
    # --- ORIGINAL SERVICE LOGIC ---
    elif len(sys.argv) == 1:
        servicemanager.Initialize()
        servicemanager.PrepareToHostSingle(FlaskService)
        servicemanager.StartServiceCtrlDispatcher()
    else:
        win32serviceutil.HandleCommandLine(FlaskService)
