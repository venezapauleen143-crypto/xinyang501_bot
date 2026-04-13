"""
XTTS v2 本地語音伺服器 (Python 3.11)
監聽 localhost:5678，POST /tts {"text": "..."} → 回傳 WAV bytes
啟動：C:\Python311\python.exe xtts_server.py
"""
import io
import json
import subprocess
import tempfile
import threading
from http.server import BaseHTTPRequestHandler, HTTPServer
from pathlib import Path

PORT = 5678
REFERENCE_WAV = str(Path(__file__).parent / "xtts_reference.wav")
LANGUAGE = "zh-cn"

print("[XTTS] 載入模型中，請稍候...")
from TTS.api import TTS

tts_model = TTS("tts_models/multilingual/multi-dataset/xtts_v2", gpu=False)
tts_lock = threading.Lock()
print(f"[XTTS] 模型載入完成，監聽 port {PORT}")


class Handler(BaseHTTPRequestHandler):
    def log_message(self, format, *args):
        pass  # 靜音 access log

    def do_POST(self):
        if self.path != "/tts":
            self.send_error(404)
            return

        length = int(self.headers.get("Content-Length", 0))
        body = json.loads(self.rfile.read(length))
        text = body.get("text", "").strip()
        if not text:
            self.send_error(400, "empty text")
            return

        try:
            with tts_lock:
                tmp_wav = tempfile.mktemp(suffix=".wav")
                tts_model.tts_to_file(
                    text=text,
                    speaker_wav=REFERENCE_WAV,
                    language=LANGUAGE,
                    file_path=tmp_wav,
                )
            wav_bytes = Path(tmp_wav).read_bytes()
            Path(tmp_wav).unlink(missing_ok=True)

            self.send_response(200)
            self.send_header("Content-Type", "audio/wav")
            self.send_header("Content-Length", str(len(wav_bytes)))
            self.end_headers()
            self.wfile.write(wav_bytes)
        except Exception as e:
            print(f"[XTTS] 生成失敗：{e}")
            self.send_error(500, str(e))

    def do_GET(self):
        if self.path == "/health":
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b"ok")
        else:
            self.send_error(404)


if __name__ == "__main__":
    server = HTTPServer(("127.0.0.1", PORT), Handler)
    print(f"[XTTS] 伺服器啟動 http://127.0.0.1:{PORT}")
    server.serve_forever()
