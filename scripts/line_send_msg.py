"""
LINE 傳訊息腳本
用法：python line_send_msg.py <好友名稱> <訊息內容>
範例：python line_send_msg.py 佳瑩 你好啊

流程：
1. 定位 LINE 視窗 + 偵測當前頁面
2. 切到好友頁（如果不在的話）
3. 點搜尋欄 → 輸入好友名稱
4. PaddleOCR 或 Vision 從搜尋結果找到好友 → 點擊
5. 確認進入該好友的對話視窗
6. 點輸入框 → 打字 → Enter 發送
"""
import sys
import io
import os
import re
import json
import time
import base64
import ctypes
import signal
import atexit
import gc

# DPI + 編碼
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(0)
except Exception:
    pass

if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import win32gui
import pyautogui
import pyperclip
from PIL import Image
from dotenv import load_dotenv
from pathlib import Path

# 把 scripts 目錄加入 path
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SCRIPT_DIR.parent))

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"


# ============================================================
# GPU 記憶體清理（atexit + 信號處理）
# ============================================================
def _cleanup_gpu():
    """退出時釋放 GPU 記憶體"""
    try:
        gc.collect()
        import paddle
        paddle.device.cuda.empty_cache()
    except Exception:
        pass


atexit.register(_cleanup_gpu)


def _signal_handler(signum, frame):
    _cleanup_gpu()
    sys.exit(0)


signal.signal(signal.SIGINT, _signal_handler)
if hasattr(signal, "SIGBREAK"):
    signal.signal(signal.SIGBREAK, _signal_handler)


def send_message(contact_name, message, monitor=None):
    """完整流程：定位 LINE → 搜尋好友 → 點進對話 → 發送訊息
    所有定位邏輯由 line_locate.py 處理"""

    from line_locate import (
        locate_line_regions, switch_page, find_line_window,
        search_friend_and_scan, enter_chat_from_search,
    )

    # Step 1: 置前 LINE 視窗（不動大小）
    print(f"[Step 1] 定位 LINE 視窗...", flush=True)
    line = find_line_window()
    if not line:
        print("[ERROR] 找不到 LINE 視窗", flush=True)
        return False
    import win32con
    try:
        SWP_FLAGS = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
        win32gui.SetWindowPos(line[0], win32con.HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        win32gui.SetWindowPos(line[0], win32con.HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        win32gui.SetWindowPos(line[0], win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              SWP_FLAGS | win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(line[0])
    except Exception as e:
        print(f"[Step 1] 置前失敗但繼續: {e}", flush=True)
    time.sleep(0.5)

    # Step 2: 定位 + 切到好友頁
    regions = locate_line_regions(monitor)
    print(f"[Step 2] 當前頁面: {regions['current_page']}", flush=True)
    if regions["current_page"] != "friend":
        regions = switch_page(regions, "friend", monitor)

    # Step 3: 搜尋好友 + OCR 掃描搜尋結果（line_locate.py 處理）
    print(f"[Step 3] 搜尋好友: {contact_name}", flush=True)
    friend_pos = search_friend_and_scan(regions, contact_name, monitor)
    if friend_pos is None:
        print("[ERROR] 找不到好友", flush=True)
        return False

    # Step 4: 點擊好友 → 判斷有無聊天 → 進入對話（line_locate.py 處理）
    print(f"[Step 4] 點擊好友進入對話...", flush=True)
    regions = enter_chat_from_search(friend_pos, regions, monitor)
    if regions is None:
        print("[ERROR] 無法進入聊天視窗", flush=True)
        return False

    # Step 5: 點輸入框 → 輸入訊息 → Enter 發送
    print(f"[Step 5] 發送訊息...", flush=True)
    ix, iy = regions["input_box"]["center"]
    pyautogui.click(ix, iy)
    time.sleep(0.3)
    pyperclip.copy(message)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1)

    print(f"[Step 5] 已送出: {message}", flush=True)
    print(f"完成", flush=True)
    return True


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        print("用法：python line_send_msg.py <好友名稱> <訊息內容>")
        sys.exit(0)

    contact = sys.argv[1]
    msg = " ".join(sys.argv[2:])
    send_message(contact, msg)
