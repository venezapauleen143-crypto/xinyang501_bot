"""
反詐 demo 語音轉文字模組

提供 3 個高層 API：
  - record_loopback(duration_sec) -> wav_path
  - transcribe_wav(wav_path) -> text
  - record_and_transcribe(duration_sec, label=None) -> text  （一鍵）

Whisper model 全域 singleton（首次呼叫載入，之後重複用，避免每次 7 秒載入）。

抓 LINE 程式的播放音訊用 WASAPI loopback（系統音訊回送），抓「預設輸出裝置」
所收到的所有聲音。點 LINE 語音時請確保預設裝置就是你聽聲音用的喇叭。
"""

import os
import sys
import io
import re
import time
import wave
from pathlib import Path
from datetime import datetime

# stdout utf-8（避免 emoji / 中文亂碼）
if sys.stdout.encoding and sys.stdout.encoding.lower() != "utf-8":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
    sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

import numpy as np
import pyaudiowpatch as pyaudio
# faster_whisper 改 lazy import（避免跟 paddle 同時載 cuDNN DLL 撞）

# 暫存音檔目錄
_VOICE_DIR = Path(os.environ.get("USERPROFILE", ".")) / ".cache" / "fanzha_voice"
_VOICE_DIR.mkdir(parents=True, exist_ok=True)

# 🔴 STT 守護進程 singleton（subprocess 隔離 paddle 跟 ctranslate2 cuDNN 衝突）
# 業界 2026 共識（來源：CTranslate2 issue / PaddleOCR-VL doc）
_STT_PROCESS = None
_STT_LOCK = None  # 多 thread 同時請求時序列化（pyaudio + STT 各跑各的）


def _get_stt_process():
    """惰性啟動守護進程 + 等模型載完。失敗回 None。"""
    global _STT_PROCESS, _STT_LOCK
    import subprocess
    import threading
    if _STT_LOCK is None:
        _STT_LOCK = threading.Lock()

    # 已有 process 且還活著 → 直接用
    if _STT_PROCESS is not None and _STT_PROCESS.poll() is None:
        return _STT_PROCESS

    server_script = str(Path(__file__).parent / "反詐_voice_stt_server.py")
    print(f"[反詐_voice] 啟動 STT 守護進程 → {server_script}", flush=True)
    env = os.environ.copy()
    env["PYTHONIOENCODING"] = "utf-8"
    _STT_PROCESS = subprocess.Popen(
        [sys.executable, "-u", server_script],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        env=env,
        encoding="utf-8",
        errors="replace",
        bufsize=1,  # line buffered
    )

    # 等 __READY__ 標記（模型載完）
    t0 = time.time()
    while True:
        line = _STT_PROCESS.stdout.readline()
        if not line:
            err = _STT_PROCESS.stderr.read() if _STT_PROCESS.stderr else "(no stderr)"
            print(f"[反詐_voice] STT 守護進程啟動失敗：{err[:500]}", flush=True)
            _STT_PROCESS = None
            return None
        line = line.strip()
        if line == "__READY__":
            print(f"[反詐_voice] STT 守護進程就緒（{time.time()-t0:.1f}s）", flush=True)
            return _STT_PROCESS
        if time.time() - t0 > 90:
            print(f"[反詐_voice] STT 守護進程 90 秒未就緒，放棄", flush=True)
            try:
                _STT_PROCESS.terminate()
            except Exception:
                pass
            _STT_PROCESS = None
            return None


def _shutdown_stt_process():
    """主進程退出時清理守護進程"""
    global _STT_PROCESS
    if _STT_PROCESS is None:
        return
    try:
        if _STT_PROCESS.poll() is None:
            _STT_PROCESS.stdin.write("__EXIT__\n")
            _STT_PROCESS.stdin.flush()
            _STT_PROCESS.wait(timeout=5)
    except Exception:
        try:
            _STT_PROCESS.terminate()
        except Exception:
            pass
    _STT_PROCESS = None


import atexit
atexit.register(_shutdown_stt_process)


def _get_loopback_device(p):
    """取得預設輸出裝置對應的 loopback device info"""
    default = p.get_default_output_device_info()
    for lb in p.get_loopback_device_info_generator():
        if default["name"] in lb["name"]:
            return lb
    raise RuntimeError(f"找不到預設輸出 '{default['name']}' 對應的 loopback 裝置")


def record_loopback(duration_sec, label=None):
    """錄系統音訊回送 duration_sec 秒，存 wav 到暫存目錄並回傳路徑。

    label: 可選的識別字串（顯示在 log + 檔名）
    """
    label = label or datetime.now().strftime("%H%M%S")
    wav_path = _VOICE_DIR / f"voice_{datetime.now().strftime('%Y%m%d_%H%M%S')}_{label}.wav"

    p = pyaudio.PyAudio()
    try:
        device = _get_loopback_device(p)
        sample_rate = int(device["defaultSampleRate"])
        channels = device["maxInputChannels"]

        print(f"[反詐_voice] 開始錄音 {duration_sec}s（label={label}）", flush=True)
        frames = []

        def callback(in_data, frame_count, time_info, status):
            frames.append(in_data)
            return (in_data, pyaudio.paContinue)

        stream = p.open(
            format=pyaudio.paInt16,
            channels=channels,
            rate=sample_rate,
            frames_per_buffer=1024,
            input=True,
            input_device_index=device["index"],
            stream_callback=callback,
        )
        time.sleep(duration_sec)
        stream.stop_stream()
        stream.close()

        wf = wave.open(str(wav_path), "wb")
        wf.setnchannels(channels)
        wf.setsampwidth(pyaudio.get_sample_size(pyaudio.paInt16))
        wf.setframerate(sample_rate)
        wf.writeframes(b"".join(frames))
        wf.close()
        kb = wav_path.stat().st_size / 1024
        print(f"[反詐_voice] 錄音完成 → {wav_path.name}（{kb:.0f} KB）", flush=True)
        return str(wav_path)
    finally:
        p.terminate()


def transcribe_wav(wav_path, language="zh"):
    """把 wav 路徑透過守護進程轉文字。失敗回空字串。

    通訊：寫 wav 路徑到守護進程 stdin，讀一行 JSON 結果。
    """
    import json
    proc = _get_stt_process()
    if proc is None:
        print(f"[反詐_voice] STT 守護進程不可用，跳過 {wav_path}", flush=True)
        return ""

    # 多 thread 同時呼叫要序列化（守護進程一次只能處理一個請求）
    with _STT_LOCK:
        try:
            proc.stdin.write(f"{wav_path}\n")
            proc.stdin.flush()
            line = proc.stdout.readline()
        except (BrokenPipeError, OSError) as e:
            print(f"[反詐_voice] STT 通訊失敗：{e}", flush=True)
            return ""

    if not line:
        print(f"[反詐_voice] STT 守護進程無回應（可能已死）", flush=True)
        return ""

    try:
        result = json.loads(line.strip())
    except Exception as e:
        print(f"[反詐_voice] STT 回傳 JSON 解析失敗（{e}）：{line[:120]!r}", flush=True)
        return ""

    if not result.get("ok"):
        print(f"[反詐_voice] STT 失敗：{result.get('error', '?')}", flush=True)
        return ""

    text = result.get("text", "")
    print(f"[反詐_voice] STT 完成（{result.get('elapsed_sec', '?')}s, 語言={result.get('language')}信心{result.get('confidence', 0):.2f}）: {text[:80]}", flush=True)
    return text


def record_and_transcribe(duration_sec, label=None, language="zh"):
    """一鍵：錄 + 轉。回傳文字（轉失敗回空字串）"""
    try:
        wav_path = record_loopback(duration_sec, label=label)
        return transcribe_wav(wav_path, language=language)
    except Exception as e:
        print(f"[反詐_voice] 錄音/轉文字失敗：{type(e).__name__}: {e}", flush=True)
        return ""


# ============================================================
# 語音訊息偵測 + 三層 ▶ 按鈕定位（方案 C）
# ============================================================

# 語音指紋：LINE 把「儲存|另存新檔|分享|傳送至Keep筆記」當每個語音訊息的選單列
# 三重容錯比對（PaddleOCR 偶爾辨識變形）
def _is_voice_fingerprint(text):
    """判斷一段 OCR 文字是否為語音訊息底下的選單指紋"""
    if not text:
        return False
    has_save = "儲存" in text
    has_other = ("另存" in text) or ("Keep" in text) or ("Kep" in text)
    return has_save and has_other


def find_voice_markers(messages):
    """從 ocr_scan_chat 返回的 messages list 中找出對方語音訊息位置。

    參數：
        messages: [{"text", "sender", "y", "conf"}]

    返回：
        [{"marker_y": int, "marker_text": str}] — 每個對方語音訊息一個元素
    """
    markers = []
    for m in messages:
        if m.get("sender") != "them":
            continue
        if _is_voice_fingerprint(m.get("text", "")):
            markers.append({"marker_y": m["y"], "marker_text": m["text"]})
    return markers


# ▶ 按鈕模板路徑（從 chat_voice_ocr.png 切出的 32x32 圖）
_TEMPLATE_PATH = Path(__file__).parent / "assets" / "voice_play_button.png"
_TEMPLATE_CACHE = None


def _get_play_template():
    """惰性載入 ▶ 模板（numpy array）"""
    global _TEMPLATE_CACHE
    if _TEMPLATE_CACHE is None:
        if not _TEMPLATE_PATH.exists():
            return None
        from PIL import Image
        _TEMPLATE_CACHE = np.array(Image.open(_TEMPLATE_PATH).convert("RGB"))
    return _TEMPLATE_CACHE


def detect_play_button(chat_img_arr, marker_y):
    """三層偵測 ▶ 按鈕位置（chat_area 內相對座標）。

    參數：
        chat_img_arr: chat_area 截圖 (H, W, 3) numpy array
        marker_y: 該語音的指紋 y（從 find_voice_markers 拿）

    返回：
        (x, y, method) — method ∈ {"pixel", "template"}；找不到回 (None, None, None)
    """
    h, w = chat_img_arr.shape[:2]

    # ─── 第一層：像素掃描（先排頭像 x>=60，找不到再放寬到 x>=30 支援省略頭像情境）
    for x_left_min in (60, 30):
        y_top = max(0, marker_y - 50)
        y_bot = max(y_top + 1, marker_y - 5)
        x_right = min(w, 280)
        region = chat_img_arr[y_top:y_bot, x_left_min:x_right]
        dark = (region[:, :, 0] < 90) & (region[:, :, 1] < 90) & (region[:, :, 2] < 90)
        ys_d, xs_d = np.where(dark)
        if len(xs_d) < 30:
            continue
        x_min = xs_d.min()
        play_mask = xs_d < (x_min + 30)
        if play_mask.any():
            play_xs = xs_d[play_mask] + x_left_min
            play_ys = ys_d[play_mask] + y_top
            return int(play_xs.mean()), int(play_ys.mean()), "pixel"

    # ─── 第二層：模板比對 ───
    template = _get_play_template()
    if template is None:
        return None, None, None
    try:
        import cv2
        # 限制搜尋範圍到指紋上方 60px 內，避免誤抓
        y_top = max(0, marker_y - 60)
        y_bot = max(y_top + template.shape[0] + 1, marker_y)
        sub = chat_img_arr[y_top:y_bot, :, :]
        res = cv2.matchTemplate(sub, template, cv2.TM_CCOEFF_NORMED)
        min_v, max_v, min_loc, max_loc = cv2.minMaxLoc(res)
        if max_v >= 0.6:  # 信心夠才採用
            tx, ty = max_loc
            cx = tx + template.shape[1] // 2
            cy = ty + y_top + template.shape[0] // 2
            return int(cx), int(cy), "template"
    except Exception as e:
        print(f"[反詐_voice] 模板比對失敗：{e}", flush=True)

    # ─── 第三層：放棄 ───
    return None, None, None


def verify_playing_started(before_arr, after_arr, play_xy, threshold=15):
    """驗證點擊播放後，氣泡是否進入「播放中」狀態。

    LINE 播放中的氣泡會出現進度條（深色細線），氣泡內像素變化量會增加。

    參數：
        before_arr, after_arr: 點擊前/後的 chat_area numpy array
        play_xy: 播放按鈕中心 (x, y)
        threshold: 平均像素差異閾值（預設 15）

    返回：
        True = 已開始播放
    """
    if before_arr.shape != after_arr.shape:
        return False
    cx, cy = play_xy
    # 取播放按鈕右側區域（進度條位置）30x40
    x1 = max(0, cx + 15)
    x2 = min(before_arr.shape[1], cx + 100)
    y1 = max(0, cy - 8)
    y2 = min(before_arr.shape[0], cy + 8)
    if x2 <= x1 or y2 <= y1:
        return False
    diff = np.abs(before_arr[y1:y2, x1:x2].astype(int) - after_arr[y1:y2, x1:x2].astype(int))
    avg = diff.mean()
    return avg > threshold


if __name__ == "__main__":
    print("=== 反詐_voice 自我測試 ===")
    print("6 秒倒數後開始錄音 8 秒，請在錄音開始時播放 LINE 語音訊息")
    for i in range(6, 0, -1):
        print(f"  準備... {i}", flush=True)
        time.sleep(1)
    print(">>> 開始錄音！現在點 LINE 語音 <<<", flush=True)
    text = record_and_transcribe(8, label="selftest")
    print(f"\n結果：{text!r}")
