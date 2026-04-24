"""
LINE 多客戶自動回覆腳本（SOP 話術引擎）
用法：python line_auto_multi.py <監控時間HH:MM> [SOP設定檔.json]
範例：python line_auto_multi.py 23:30 scripts/line_sop/織夢小棧.json

流程：
1. 開啟聊天頁
2. 截圖偵測有綠色未讀標記的對話
3. 點第一個綠色標記 → 進入聊天
4. OCR 讀取對話內容 → 判斷 SOP 步驟
5. Claude AI 生成回覆 → 發送
6. 回到聊天列表
7. 點下一個綠色標記 → 重複 3-6
8. 全部處理完 → 等待新的未讀
9. 持續循環直到監控時間結束
"""
import sys
import io
import os
import re
import json
import time
import signal
import ctypes
import hashlib
import types
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
from datetime import datetime
from dotenv import load_dotenv
from pathlib import Path
import anthropic

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SCRIPT_DIR.parent))

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
client = anthropic.Anthropic()
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"
DEFAULT_SOP = "C:/Users/blue_/claude-telegram-bot/scripts/line_sop/織夢小棧.json"

# ============================================================
# GPU 記憶體清理
# ============================================================
def _cleanup_gpu():
    try:
        gc.collect()
        import paddle
        paddle.device.cuda.empty_cache()
        print("[Cleanup] GPU 記憶體已釋放", flush=True)
    except Exception:
        pass

atexit.register(_cleanup_gpu)

# ============================================================
# 優雅停止機制
# ============================================================
STOP_FILE = "C:/Users/blue_/Desktop/測試檔案/.stop_line_multi"
_should_stop = False

def _signal_handler(signum, frame):
    global _should_stop
    _should_stop = True
    print(f"\n[STOP] 收到信號 {signum}，準備停止...", flush=True)
    _cleanup_gpu()

signal.signal(signal.SIGINT, _signal_handler)
if hasattr(signal, "SIGBREAK"):
    signal.signal(signal.SIGBREAK, _signal_handler)

def should_stop():
    global _should_stop
    if _should_stop:
        return True
    if os.path.exists(STOP_FILE):
        try:
            os.remove(STOP_FILE)
        except OSError:
            pass
        _should_stop = True
        print("[STOP] 偵測到停止旗標檔案，準備停止...", flush=True)
        return True
    return False

# ============================================================
# SOP 設定檔（從 line_auto_chat.py 引用）
# ============================================================
from line_auto_chat import (
    load_sop, build_system_prompt, generate_reply,
    filter_reply, send_reply, send_multi_reply, send_image,
    grab_chat_area, ocr_extract_messages, chat_hash,
)

# ============================================================
# 時間判斷
# ============================================================
def time_to_minutes(t_str):
    h, m = map(int, t_str.split(":"))
    return h * 60 + m

def is_before_stop_time(stop_time):
    now_min = time_to_minutes(datetime.now().strftime("%H:%M"))
    stop_min = time_to_minutes(stop_time)
    return now_min < stop_min

# ============================================================
# 偵測聊天列表中有未讀標記的對話
# ============================================================
def find_unread_conversations(monitor=2):
    """
    切到聊天頁，用 line_locate.py 的 find_unread_badges 偵測綠色未讀標記。

    回傳：(regions, unread_list)
        regions: 當前定位
        unread_list: [{"y": int, "center": (x, y)}] 有未讀的對話座標
    """
    from line_locate import (
        locate_line_regions, switch_page, find_unread_badges,
    )

    regions = locate_line_regions(monitor)

    # 確保在聊天頁
    if regions["current_page"] != "chat":
        regions = switch_page(regions, "chat", monitor)

    # 用定位腳本的像素偵測找綠色徽章（LINE16 方式）
    unread_list = find_unread_badges(monitor)

    return regions, unread_list


# ============================================================
# 處理單一客戶的對話
# ============================================================
def handle_one_customer(conv, regions, system_prompt, sop, all_histories, monitor=2):
    """
    點進一個客戶的對話，讀取內容，判斷 SOP 步驟，回覆。

    參數：
        conv: {"name", "unread", "center"} 從聊天列表取得
        regions: 當前的 regions
        system_prompt: SOP system prompt
        sop: SOP JSON
        all_histories: {客戶名: [conversation_history]} 所有客戶的對話歷史
        monitor: 螢幕編號

    回傳：更新後的 regions
    """
    from line_locate import locate_line_regions, ocr_scan_panel, screenshot_line

    cx, cy = conv["center"]
    print(f"\n[Customer] 點擊未讀對話 at ({cx}, {cy})", flush=True)

    # Step 1: 點擊該對話進入聊天
    pyautogui.click(cx, cy)
    time.sleep(1.5)

    # Step 2: 重新定位（進入聊天後）
    regions = locate_line_regions(monitor)

    # 從 chat_title 區域讀取客戶名稱
    ct = regions.get("chat_title", {})
    # 用 OCR 讀 chat_title
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
    sx_r = full_img.size[0] / mon["width"]
    sy_r = full_img.size[1] / mon["height"]
    ct_il = max(0, int((ct.get("left", 0) - mon["left"]) * sx_r) - il)
    ct_it = max(0, int((ct.get("top", 0) - mon["top"]) * sy_r) - it)
    ct_ir = min(line_crop.size[0], int((ct.get("right", 0) - mon["left"]) * sx_r) - il)
    ct_ib = min(line_crop.size[1], int((ct.get("bottom", 0) - mon["top"]) * sy_r) - it)
    if ct_ir > ct_il and ct_ib > ct_it:
        title_crop = line_crop.crop((ct_il, ct_it, ct_ir, ct_ib))
        title_items = ocr_scan_panel(title_crop)
        name = " ".join(item["text"] for item in title_items).strip() if title_items else "unknown"
    else:
        name = "unknown"
    print(f"[Customer] 客戶名稱: {name}", flush=True)

    # Step 3: 截圖對話區 + OCR 讀取內容
    chat_img = grab_chat_area(regions, monitor)
    current_messages = ocr_extract_messages(chat_img)
    print(f"[Customer] OCR 讀到 {len(current_messages)} 條訊息", flush=True)

    # Step 4: 取得或建立該客戶的對話歷史
    if name not in all_histories:
        all_histories[name] = []
    history = all_histories[name]

    # Step 5: 找出新的對方訊息
    # 如果歷史是空的，用 OCR 讀到的所有對方訊息
    if not history:
        new_them = [m["text"] for m in current_messages if m["sender"] == "them"]
        # 把所有訊息加入歷史
        for m in current_messages:
            history.append(m)
    else:
        # 用最後一條已知訊息定位新訊息
        from difflib import SequenceMatcher
        last_known = history[-1]
        last_text = last_known["text"]

        match_idx = -1
        for i in range(len(current_messages) - 1, -1, -1):
            ratio = SequenceMatcher(None, current_messages[i]["text"], last_text).ratio()
            if ratio > 0.6:
                match_idx = i
                break

        if match_idx >= 0 and match_idx < len(current_messages) - 1:
            new_msgs = current_messages[match_idx + 1:]
            new_them = [m["text"] for m in new_msgs if m["sender"] == "them"]
            for m in new_msgs:
                history.append(m)
        else:
            # 找不到匹配，用所有對方訊息
            new_them = [m["text"] for m in current_messages if m["sender"] == "them"]

    if not new_them:
        print(f"[Customer] 沒有新的對方訊息，跳過", flush=True)
        return regions

    print(f"[Customer] 新訊息: {new_them}", flush=True)

    # Step 6: Claude AI 生成回覆
    reply = generate_reply(system_prompt, history, new_them)

    if reply and len(reply) > 1:
        # 偵測 [END] 標記
        if "[END]" in reply:
            print(f"[Customer] SOP 結束（{name}）", flush=True)
            return regions

        # 過濾 AI 分析內容
        reply = filter_reply(reply)
        if not reply:
            print(f"[Customer] 過濾後為空，跳過", flush=True)
            return regions

        print(f"[Customer] 回覆: {reply[:80]}", flush=True)

        # 發送回覆
        send_multi_reply(reply, regions)

        # 回覆加入歷史
        for part in reply.split("|||"):
            part = part.strip()
            if part and not (part.startswith("{") and part.endswith("}")):
                history.append({"text": part, "sender": "me", "y": 0})

        # 偵測到編號 → 分享溫妮好友資訊
        if "編號" in reply:
            print(f"[Customer] 偵測到編號，分享溫妮好友資訊給 {name}", flush=True)
            from line_locate import share_contact_card
            share_contact_card(regions, "溫妮", name, monitor)

        time.sleep(2)

    return regions


# ============================================================
# 主流程
# ============================================================
def main(stop_time, sop_path=DEFAULT_SOP, monitor=2):
    from line_locate import (
        locate_line_regions, switch_page, find_line_window,
        screenshot_line_window,
    )
    import win32con

    # 清殘留旗標
    if os.path.exists(STOP_FILE):
        try:
            os.remove(STOP_FILE)
        except OSError:
            pass

    print("=" * 50, flush=True)
    print(f"LINE 多客戶自動回覆", flush=True)
    print(f"監控到：{stop_time}", flush=True)
    print(f"SOP：{sop_path}", flush=True)
    print("=" * 50, flush=True)

    # 載入 SOP
    sop = load_sop(sop_path)
    system_prompt = build_system_prompt(sop)
    print(f"[Init] SOP: {sop['name']}（{len(sop['steps'])} 步驟）", flush=True)

    # 置前 LINE
    line = find_line_window()
    if not line:
        print("[ERROR] 找不到 LINE 視窗", flush=True)
        return False

    SWP_FLAGS = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
    try:
        win32gui.SetWindowPos(line[0], win32con.HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        win32gui.SetWindowPos(line[0], win32con.HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        win32gui.SetWindowPos(line[0], win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              SWP_FLAGS | win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(line[0])
    except Exception:
        pass
    time.sleep(0.5)

    # 所有客戶的對話歷史
    all_histories = {}

    # 主迴圈
    POLL_INTERVAL = 10  # 每 10 秒檢查一次

    print(f"\n[Monitor] 開始監控未讀訊息...", flush=True)
    print(f"[Monitor] 停止方式：touch {STOP_FILE}", flush=True)

    while is_before_stop_time(stop_time):
        if should_stop():
            break

        try:
            # Step 1: 切到聊天頁，用定位腳本偵測綠色未讀標記
            regions, unread_list = find_unread_conversations(monitor)

            if not unread_list:
                # 沒有未讀，等待
                for _ in range(POLL_INTERVAL):
                    if should_stop():
                        break
                    time.sleep(1)
                continue

            print(f"\n[Monitor] 偵測到 {len(unread_list)} 個未讀對話", flush=True)

            # Step 2: 逐一處理每個未讀對話
            for conv in unread_list:
                if should_stop():
                    break

                # 處理這個客戶
                regions = handle_one_customer(
                    conv, regions, system_prompt, sop, all_histories, monitor
                )

                if should_stop():
                    break

                # 回到聊天列表
                from line_locate import switch_page
                regions = switch_page(regions, "chat", monitor)
                time.sleep(1)

            # 全部處理完，等待新的未讀
            print(f"[Monitor] 本輪處理完畢，等待新訊息...", flush=True)
            for _ in range(POLL_INTERVAL):
                if should_stop():
                    break
                time.sleep(1)

        except Exception as e:
            print(f"[ERR] {e}", flush=True)
            import traceback
            traceback.print_exc()
            time.sleep(5)

    reason = "收到停止信號" if _should_stop else "到達結束時間"
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 監控結束（{reason}）", flush=True)
    print(f"共處理 {len(all_histories)} 個客戶", flush=True)
    return True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        print("用法：python line_auto_multi.py <監控時間HH:MM> [SOP設定檔.json]")
        sys.exit(0)

    stop = sys.argv[1]
    sop = sys.argv[2] if len(sys.argv) > 2 else DEFAULT_SOP
    main(stop, sop)
