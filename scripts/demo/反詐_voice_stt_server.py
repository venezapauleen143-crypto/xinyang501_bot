"""語音轉文字守護進程（subprocess 隔離 paddle 跟 cuDNN 衝突）

工作流程：
  啟動 → 載入 faster-whisper large-v3 GPU 模式
       → 印 __READY__ 標記
       → 等 stdin 收 wav 路徑（一行一個）
       → 每個 wav 轉文字後印一行 JSON 到 stdout
       → 收到 __EXIT__ 或 stdin 關閉就退出

通訊協定：
  stdin  : 一行一個 wav 絕對路徑（utf-8）
  stdout : 一行一個 JSON {"ok": true, "text": "...", "language": "zh", "confidence": 0.99}
           失敗回 {"ok": false, "error": "..."}
  stderr : 紀錄 log（不參與通訊）
"""
import sys
import io
import json
import time
import traceback

# stdin/stdout 強制 utf-8 + line buffering
sys.stdin = io.TextIOWrapper(sys.stdin.buffer, encoding="utf-8", errors="replace")
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace", write_through=True)
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace", write_through=True)

print("[stt-server] 啟動 — 載入 faster-whisper large-v3 GPU 模式...", file=sys.stderr, flush=True)
t0 = time.time()

try:
    from faster_whisper import WhisperModel
    model = WhisperModel("large-v3", device="cuda", compute_type="float16")
    device = "cuda/float16"
except Exception as e:
    print(f"[stt-server] GPU 載入失敗（{type(e).__name__}: {e}），改 CPU int8", file=sys.stderr, flush=True)
    from faster_whisper import WhisperModel
    model = WhisperModel("large-v3", device="cpu", compute_type="int8")
    device = "cpu/int8"

elapsed = time.time() - t0
print(f"[stt-server] 模型載入完成（{elapsed:.1f}s, device={device}）", file=sys.stderr, flush=True)
print("__READY__", flush=True)  # 通知主進程已就緒

# 主迴圈：每行一個 wav 路徑
for line in sys.stdin:
    wav_path = line.strip()
    if not wav_path:
        continue
    if wav_path == "__EXIT__":
        print("[stt-server] 收到 __EXIT__，退出", file=sys.stderr, flush=True)
        break

    try:
        t0 = time.time()
        segments, info = model.transcribe(wav_path, language="zh", beam_size=5)
        segments = list(segments)
        text = "".join(seg.text for seg in segments).strip()
        result = {
            "ok": True,
            "text": text,
            "language": info.language,
            "confidence": float(info.language_probability),
            "elapsed_sec": round(time.time() - t0, 2),
        }
        print(f"[stt-server] {wav_path} → '{text[:50]}'（{result['elapsed_sec']}s）", file=sys.stderr, flush=True)
    except Exception as e:
        result = {"ok": False, "error": f"{type(e).__name__}: {e}", "traceback": traceback.format_exc()[:500]}
        print(f"[stt-server] FAIL {wav_path}: {e}", file=sys.stderr, flush=True)

    print(json.dumps(result, ensure_ascii=False), flush=True)

print("[stt-server] 主迴圈結束", file=sys.stderr, flush=True)
