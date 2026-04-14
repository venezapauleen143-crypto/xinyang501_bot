#!/usr/bin/env python3
"""
RVC HTTP server for Jay Chou voice conversion.
POST /convert  - body: WAV bytes, response: converted WAV bytes
GET  /health   - response: "ok"
"""

import io
import os
import sys
import tempfile
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

PORT = 5679
BOT_DIR = Path(__file__).parent
MODEL_PTH = str(BOT_DIR / "rvc_models" / "jay_chou.pth")
MODEL_INDEX = str(BOT_DIR / "rvc_models" / "jay_chou.index")

rvc_instance = None


def load_rvc():
    global rvc_instance
    from rvc_python.infer import RVCInference
    print("[RVC] Loading RVCInference...")
    rvc_instance = RVCInference(device="cpu")
    rvc_instance.load_model(MODEL_PTH, version="v2", index_path=MODEL_INDEX)
    rvc_instance.set_params(
        f0method="harvest",
        f0up_key=0,
        index_rate=0.6,
        filter_radius=3,
        rms_mix_rate=1,
        protect=0.33,
    )
    print("[RVC] Model loaded OK")


class RVCHandler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # suppress request logs

    def do_GET(self):
        if self.path == "/health":
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"ok")
        else:
            self.send_response(404)
            self.end_headers()

    def do_POST(self):
        if self.path != "/convert":
            self.send_response(404)
            self.end_headers()
            return

        length = int(self.headers.get("Content-Length", 0))
        wav_bytes = self.rfile.read(length)

        try:
            with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as fin:
                fin.write(wav_bytes)
                in_path = fin.name

            out_path = in_path.replace(".wav", "_rvc.wav")
            rvc_instance.infer_file(in_path, out_path)

            with open(out_path, "rb") as f:
                result = f.read()

            os.unlink(in_path)
            os.unlink(out_path)

            self.send_response(200)
            self.send_header("Content-Type", "audio/wav")
            self.send_header("Content-Length", str(len(result)))
            self.end_headers()
            self.wfile.write(result)

        except Exception as e:
            print(f"[RVC] Error: {e}")
            os.unlink(in_path) if os.path.exists(in_path) else None
            self.send_response(500)
            self.end_headers()
            self.wfile.write(str(e).encode())


if __name__ == "__main__":
    sys.path.insert(0, str(Path(__file__).parent.parent / "fairseq_src" / "fairseq-0.12.2"))
    load_rvc()
    server = HTTPServer(("127.0.0.1", PORT), RVCHandler)
    print(f"[RVC] Server running on port {PORT}")
    server.serve_forever()
