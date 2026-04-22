"""
LINE 自動回覆腳本（SOP 話術引擎）
用法：python line_auto_chat.py <好友名稱> <監控時間HH:MM> [SOP設定檔.json]
範例：python line_auto_chat.py "仁輝 JAMES" 23:30 scripts/line_sop/織夢小棧.json

流程：
1. 讀取 SOP 設定檔（JSON）
2. 搜尋好友 → 進入對話視窗
3. PaddleOCR 監控對話區 → 偵測新訊息
4. Claude AI 根據 SOP + 對話上下文判斷回覆
5. 自動發送回覆
6. 監控到指定時間或收到停止信號
"""
import sys
import io
import os
import re
import json
import time
import base64
import hashlib
import ctypes
import signal
import types
import numpy as np

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
import mss
from PIL import Image
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
# GPU 記憶體清理（atexit + 信號處理）
# ============================================================
import atexit
import gc


def _cleanup_gpu():
    """退出時釋放 GPU 記憶體"""
    try:
        gc.collect()
        import paddle
        paddle.device.cuda.empty_cache()
        print("[Cleanup] GPU 記憶體已釋放", flush=True)
    except Exception:
        pass


atexit.register(_cleanup_gpu)


# ============================================================
# 啟動前清理殘留 GPU 進程
# ============================================================
def _kill_stale_gpu_processes():
    """殺掉殘留的同名 Python 腳本進程（防止 VRAM 雙倍佔用）"""
    import subprocess as _sp
    my_pid = os.getpid()
    script_name = "line_auto_chat.py"
    try:
        # 只找 python 進程，排除 bash/wmic/其他
        result = _sp.run(
            ["wmic", "process", "where",
             f"name like 'python%' and commandline like '%{script_name}%' and processid!='{my_pid}'",
             "get", "processid"],
            capture_output=True, text=True, timeout=5
        )
        for line in result.stdout.strip().split("\n"):
            line = line.strip()
            if line.isdigit():
                pid = int(line)
                if pid != my_pid:
                    try:
                        os.kill(pid, signal.SIGTERM)
                        print(f"[Cleanup] 殺掉殘留進程 PID={pid}", flush=True)
                    except (OSError, ProcessLookupError):
                        pass
    except Exception:
        pass


# ============================================================
# 優雅停止機制
# ============================================================
STOP_FILE = "C:/Users/blue_/Desktop/測試檔案/.stop_line_auto"
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


# PaddleOCR 由 line_locate.py 統一管理，不在這裡建立


# ============================================================
# SOP 設定檔載入
# ============================================================
def load_sop(sop_path):
    """載入 SOP 設定檔"""
    with open(sop_path, "r", encoding="utf-8") as f:
        return json.load(f)


def build_system_prompt(sop):
    """從 SOP 設定檔建構 Claude 的 System Prompt"""
    today = datetime.now().strftime("%Y年%m月%d日")

    # 課程資訊
    course = sop["course_info"]
    course_text = (
        f"課程名稱：{course['name']}\n"
        f"課程類型：{course['type']}\n"
        f"課程內容：{', '.join(course['content'])}\n"
        f"課程特色：{', '.join(course['features'])}\n"
        f"課程時間：{course['schedule']}\n"
        f"沒有課的日子：{course['no_class_day']}\n"
        f"免費堂數：{course['free_lessons']}堂\n"
        f"後續收費：{course['paid_price']}\n"
        f"工具材料：{course['tools']}\n"
        f"上課地點：{course['location']}\n"
    )

    # 規則
    rules = sop["rules"]
    rules_text = (
        f"編號生成方式：{rules['id_generation']}\n"
        f"資料傳給：{rules['forward_to']}\n"
        f"絕對不能自己編造地址：{rules['never_fabricate_address']}\n"
        f"絕對不能承諾具體開課日期：{rules['never_promise_exact_date']}\n"
        f"只有實體課程：{rules['only_physical_class']}\n"
        f"客戶拒絕時繼續推：{rules['keep_pushing_if_refused']}\n"
    )

    # SOP 步驟
    steps_text = ""
    for step in sop["steps"]:
        replies = "\n".join(step.get("replies", []))
        steps_text += f"\n【{step['id']}】{step['description']}\n回覆：{replies}\n"
        if "expect" in step:
            steps_text += f"等待客戶：{step['expect']}\n"
        if "next" in step:
            steps_text += f"下一步：{step['next']}\n"

    # FAQ
    faq_text = ""
    for category, items in sop["faq"].items():
        faq_text += f"\n【{category}】\n"
        for q, a in items.items():
            faq_text += f"  Q: {q}\n  A: {a}\n"

    # 追問話術
    follow = sop["follow_up"]
    follow_text = "\n".join(f"- {k}: {v}" for k, v in follow.items())

    # few-shot 範例
    examples_text = ""
    for ex in sop["few_shot_examples"]:
        examples_text += f"\n場景：{ex['scenario']}\n"
        if "customer" in ex:
            examples_text += f"客戶：{ex['customer']}\n"
        if "reply" in ex:
            examples_text += f"回覆：{ex['reply']}\n"
        if "action" in ex:
            examples_text += f"動作：{ex['action']}\n"
        if "customer_correction" in ex:
            examples_text += f"客戶更正：{ex['customer_correction']}\n"
            examples_text += f"回覆更正：{ex['reply_correction']}\n"

    prompt = (
        f"今天是{today}。你是「{sop['name']}」的客服小編。\n"
        f"人設：{sop['persona']}\n\n"
        f"====== 課程資訊 ======\n{course_text}\n"
        f"====== 規則（必須遵守）======\n{rules_text}\n"
        f"====== SOP 流程步驟 ======\n{steps_text}\n"
        f"====== 常見問題 FAQ ======\n{faq_text}\n"
        f"====== 追問話術 ======\n{follow_text}\n\n"
        f"====== 真實對話範例（模仿這個風格）======\n{examples_text}\n\n"
        f"====== 回覆規則（最高優先，違反任何一條都不合格）======\n"
        f"1. 你要判斷客戶目前在 SOP 的哪一步，然後回覆對應的話術\n"
        f"2. 如果客戶問了 FAQ 裡有的問題，用 FAQ 的答案回\n"
        f"3. 如果客戶問了 FAQ 裡沒有的問題，用你的判斷靈活回答，但不要離開 SOP 流程\n"
        f"4. 客戶用貼圖、OK、好、可以等正面回應 → 理解為同意，繼續流程\n"
        f"5. 客戶拒絕 → 追問原因，不輕易放棄\n"
        f"6. 客戶沒回應 → 用追問話術\n"
        f"7. 客戶資料不完整 → 提醒補齊缺少的項目\n"
        f"8. 當客戶提供完整資料（姓名+生日+電話）→ 自動用電話後五碼算出編號\n"
        f"9. 回覆要簡短自然，像真人客服打字，不要一次回太多\n"
        f"10. 每次只回覆一段話（不要一次回多段，除非是固定話術需要分段發送）\n\n"
        f"====== 絕對禁止（違反 = 暴露 AI 身份）======\n"
        f"1. 自己的對話回覆最多 2 條訊息、每條不超過 50 字。但 SOP 固定話術（報名表、課程介紹、編號通知）照原文完整發送，不限字數\n"
        f"2. 絕對不能說「不明白」「看不懂」「請問這是什麼意思」「我無法判斷」\n"
        f"3. 絕對不能把你的分析、判斷過程、確認內容發給客戶\n"
        f"4. 客戶給資料就直接收下登記，不要覆述整理、不要逐項確認、不要解釋你怎麼理解的\n"
        f"5. 資料解析規則：4碼數字=月日生日（如0922=09/22），09開頭10碼=手機號碼，中文=姓名\n"
        f"6. 數字可能黏在一起（如09200936523698），自己拆：前4碼=生日0920，後10碼=電話0936523698\n"
        f"7. 解析有疑慮就自己判斷，判斷錯了也比問客戶好，絕對不能暴露你在猜\n\n"
        f"回覆格式：只回覆要發送的文字，不要加任何格式、解釋、或標記。\n"
        f"如果需要分多條訊息發送，用 ||| 分隔，例如：第一條訊息|||第二條訊息|||第三條訊息\n"
    )
    return prompt


# ============================================================
# 對話區截圖 + OCR 訊息提取（全部呼叫 line_locate.py）
# ============================================================
def grab_chat_area(regions, monitor=2):
    """截取 LINE 對話區（呼叫 line_locate.py）"""
    from line_locate import screenshot_chat_area
    return screenshot_chat_area(regions, monitor)


def chat_hash(chat_img):
    """對話區截圖的 hash，偵測畫面變化"""
    small = chat_img.resize((80, 40))
    return hashlib.md5(small.tobytes()).hexdigest()


def ocr_extract_messages(chat_img):
    """OCR 提取對話訊息（呼叫 line_locate.py）"""
    from line_locate import ocr_scan_chat
    return ocr_scan_chat(chat_img)


def detect_new_messages(conversation_history, current_msgs):
    """比對 conversation_history（完整記錄）和 current_msgs（OCR 掃到的），找新增的對方訊息。
    用 conversation_history 最後一條訊息在 current_msgs 裡定位，之後的就是新的。
    """
    if not current_msgs:
        return []

    if not conversation_history:
        return [m["text"] for m in current_msgs if m["sender"] == "them"]

    # 拿 conversation_history 最後一條的文字，在 current_msgs 裡從後往前找
    from difflib import SequenceMatcher
    last_known = conversation_history[-1]
    last_text = last_known["text"]

    match_idx = -1
    for i in range(len(current_msgs) - 1, -1, -1):
        ratio = SequenceMatcher(None, current_msgs[i]["text"], last_text).ratio()
        if ratio > 0.6:
            match_idx = i
            break

    if match_idx >= 0 and match_idx < len(current_msgs) - 1:
        new_msgs = current_msgs[match_idx + 1:]
        return [m["text"] for m in new_msgs if m["sender"] == "them"]

    # 找不到匹配（可能最後一條已經滾出畫面），用全部 current 的最後幾條跟歷史比
    # 取 current 最後面的對方訊息，檢查是不是已經在歷史裡
    history_texts = [m["text"] for m in conversation_history[-10:]]
    new_them = []
    for msg in current_msgs:
        if msg["sender"] != "them":
            continue
        is_known = False
        for ht in history_texts:
            if SequenceMatcher(None, msg["text"], ht).ratio() > 0.6:
                is_known = True
                break
        if not is_known:
            new_them.append(msg["text"])

    return new_them


# ============================================================
# Claude AI 回覆（根據 SOP）
# ============================================================
def generate_reply(system_prompt, conversation_history, new_messages_text):
    """Claude AI 根據 SOP + 對話歷史 + 新訊息生成回覆"""
    # 組對話歷史
    history_lines = []
    for msg in conversation_history[-20:]:
        label = "[客戶]" if msg["sender"] == "them" else "[小編]"
        history_lines.append(f"{label} {msg['text']}")
    history_text = "\n".join(history_lines)

    new_text = "\n".join(f"• {m}" for m in new_messages_text)

    user_content = (
        f"以下是目前的對話紀錄：\n\n{history_text}\n\n"
        f"客戶剛發了新訊息：\n{new_text}\n\n"
        f"請根據 SOP 流程判斷目前在哪一步，然後回覆客戶。\n"
        f"如果客戶提供了完整資料（姓名+生日+電話），請用電話後五碼算出編號。\n\n"
        f"【最重要】你的回覆會直接發送給客戶看到，所以：\n"
        f"- 只寫要發送的文字，一個字的分析、判斷、解釋都不能有\n"
        f"- 不能出現「姓名=」「生日=」「電話=」「編號=」「解析」「資料」這種格式\n"
        f"- 客戶提供資料後，直接回「好的已幫您登記✅」然後給編號\n"
        f"- 多條訊息用 ||| 分隔"
    )

    r = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=500,
        system=system_prompt,
        messages=[{"role": "user", "content": user_content}]
    )
    return r.content[0].text.strip()


# ============================================================
# 發送訊息（點輸入框 → 打字 → Enter）
# ============================================================
def send_reply(msg, regions):
    """發送一條訊息"""
    ix, iy = regions["input_box"]["center"]
    pyautogui.click(ix, iy)
    time.sleep(0.3)
    pyperclip.copy(msg)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.5)


def send_multi_reply(reply_text, regions):
    """發送回覆（支援多條訊息，用 ||| 分隔）"""
    parts = [p.strip() for p in reply_text.split("|||") if p.strip()]
    for part in parts:
        # 跳過 {send_image} 和 {contact_card} 等佔位符（後續版本再實作）
        if part.startswith("{") and part.endswith("}"):
            print(f"[Reply] 跳過佔位符: {part}", flush=True)
            continue
        send_reply(part, regions)
        print(f"[Reply] → {part[:60]}", flush=True)
        time.sleep(1)


# ============================================================
# 搜尋好友 + 進入對話（全部呼叫 line_locate.py）
# ============================================================
def enter_conversation(contact_name, monitor=2):
    """搜尋好友並進入對話視窗，所有定位邏輯由 line_locate.py 處理"""
    from line_locate import (
        locate_line_regions, switch_page, screenshot_line,
        find_line_window, search_friend_and_scan, enter_chat_from_search,
    )

    # Step 1: 置前 LINE 視窗（不動大小）
    print(f"[Setup] 定位 LINE 視窗...", flush=True)
    line = find_line_window()
    if not line:
        print("[ERROR] 找不到 LINE 視窗", flush=True)
        return None
    import win32con
    try:
        SWP_FLAGS = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
        win32gui.SetWindowPos(line[0], win32con.HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        win32gui.SetWindowPos(line[0], win32con.HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        win32gui.SetWindowPos(line[0], win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              SWP_FLAGS | win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(line[0])
    except Exception as e:
        print(f"[Setup] 置前失敗但繼續: {e}", flush=True)
    time.sleep(0.5)

    # Step 2: 定位 + 切到好友頁
    regions = locate_line_regions(monitor)
    if should_stop():
        return None
    if regions["current_page"] != "friend":
        regions = switch_page(regions, "friend", monitor)
    if should_stop():
        return None

    # Step 3: 搜尋好友 + OCR 掃描搜尋結果（line_locate.py 處理）
    print(f"[Setup] 搜尋好友: {contact_name}", flush=True)
    friend_pos = search_friend_and_scan(regions, contact_name, monitor)
    if should_stop():
        return None
    if friend_pos is None:
        print("[ERROR] 找不到好友", flush=True)
        return None

    # Step 4: 點擊好友 → 判斷有無聊天 → 進入對話（line_locate.py 處理）
    print(f"[Setup] 點擊好友進入對話...", flush=True)
    regions = enter_chat_from_search(friend_pos, regions, monitor)
    if regions is None:
        print("[ERROR] 無法進入聊天視窗", flush=True)
        return None

    print(f"[Setup] 已進入對話視窗", flush=True)
    return regions


# ============================================================
# 監控 + 自動回覆主迴圈
# ============================================================
COOLDOWN_SECONDS = 15
POLL_INTERVAL = 5
WAIT_COMPLETE = 2
WAIT_COMPLETE_ROUNDS = 2


def time_to_minutes(t_str):
    """把 HH:MM 轉成分鐘數，用於數值比較（解決跨午夜問題）"""
    h, m = map(int, t_str.split(":"))
    return h * 60 + m


def is_before_stop_time(stop_time):
    """判斷現在是否還沒到停止時間"""
    now_min = time_to_minutes(datetime.now().strftime("%H:%M"))
    stop_min = time_to_minutes(stop_time)
    return now_min < stop_min


def monitor_and_reply(regions, stop_time, system_prompt, conversation_history, monitor=2):
    """監控對話區，偵測新訊息後自動回覆。
    conversation_history 是完整的對話記錄（包含歡迎詞），會在這裡持續累加。
    """

    # 初始截圖 hash（用來偵測畫面變化）
    chat = grab_chat_area(regions, monitor)
    last_hash = chat_hash(chat)
    cooldown_until = 0

    print(f"[Monitor] 開始監控 → {stop_time}", flush=True)
    print(f"[Monitor] 冷卻：{COOLDOWN_SECONDS}s / 輪詢：{POLL_INTERVAL}s", flush=True)
    print(f"[Monitor] 對話歷史：{len(conversation_history)} 條", flush=True)
    print(f"[Monitor] 停止方式：touch {STOP_FILE}", flush=True)

    while is_before_stop_time(stop_time):
        if should_stop():
            break

        for _ in range(POLL_INTERVAL):
            if should_stop():
                break
            time.sleep(1)
        if should_stop():
            break

        try:
            if time.time() < cooldown_until:
                continue

            chat = grab_chat_area(regions, monitor)
            h = chat_hash(chat)

            if h == last_hash:
                continue

            t = datetime.now().strftime("%H:%M:%S")

            # 畫面有變化，先等 3 秒確認不是自己發的訊息造成的
            time.sleep(3)
            chat = grab_chat_area(regions, monitor)
            h2 = chat_hash(chat)

            # 再等一輪確認對方說完
            latest_chat = chat
            latest_hash = h2
            for wait_round in range(WAIT_COMPLETE_ROUNDS):
                time.sleep(WAIT_COMPLETE)
                check = grab_chat_area(regions, monitor)
                check_h = chat_hash(check)
                if check_h != latest_hash:
                    print(f"[{t}] 對方還在打字（第{wait_round+1}輪），等...", flush=True)
                    latest_chat = check
                    latest_hash = check_h
                else:
                    break

            # OCR 提取訊息
            current_messages = ocr_extract_messages(latest_chat)

            # 用 conversation_history 比對，找新增的對方訊息
            new_them_msgs = detect_new_messages(conversation_history, current_messages)

            if not new_them_msgs:
                last_hash = latest_hash
                continue

            print(f"[{t}] 偵測到 {len(new_them_msgs)} 條新訊息：{new_them_msgs}", flush=True)

            # 新訊息加入 conversation_history
            for msg_text in new_them_msgs:
                conversation_history.append({"text": msg_text, "sender": "them", "y": 0})

            # Claude AI 生成回覆
            reply = generate_reply(system_prompt, conversation_history, new_them_msgs)

            if reply and len(reply) > 1:
                print(f"[{t}] AI 回覆：{reply[:100]}", flush=True)

                # 發送回覆
                send_multi_reply(reply, regions)

                # 回覆加入 conversation_history（記原始回覆文字，不是 OCR 掃回來的）
                for part in reply.split("|||"):
                    part = part.strip()
                    if part and not (part.startswith("{") and part.endswith("}")):
                        conversation_history.append({"text": part, "sender": "me", "y": 0})

                cooldown_until = time.time() + COOLDOWN_SECONDS
                print(f"[{t}] 冷卻 {COOLDOWN_SECONDS} 秒", flush=True)

                # 冷卻期間每秒檢查停止
                for _ in range(COOLDOWN_SECONDS):
                    if should_stop():
                        break
                    time.sleep(1)

                # 冷卻結束後更新 hash（這時畫面已穩定，自己的訊息已顯示完）
                chat = grab_chat_area(regions, monitor)
                last_hash = chat_hash(chat)
            else:
                print(f"[{t}] 回覆為空，跳過", flush=True)
                last_hash = latest_hash

        except Exception as e:
            print(f"[ERR] {e}", flush=True)
            import traceback
            traceback.print_exc()
            time.sleep(5)

    reason = "收到停止信號" if _should_stop else "到達結束時間"
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 監控結束（{reason}）", flush=True)


# ============================================================
# 主流程
# ============================================================
def main(contact_name, stop_time, sop_path=DEFAULT_SOP, monitor=2):
    import threading

    # 啟動前殺殘留進程（防 VRAM 雙倍佔用）
    _kill_stale_gpu_processes()

    # 清殘留旗標
    if os.path.exists(STOP_FILE):
        try:
            os.remove(STOP_FILE)
        except OSError:
            pass

    print("=" * 50, flush=True)
    print(f"LINE 自動回覆（SOP 引擎）", flush=True)
    print(f"好友：{contact_name}", flush=True)
    print(f"監控到：{stop_time}", flush=True)
    print(f"SOP：{sop_path}", flush=True)
    print("=" * 50, flush=True)

    # 載入 SOP
    print(f"\n[Init] 載入 SOP 設定檔...", flush=True)
    sop = load_sop(sop_path)
    system_prompt = build_system_prompt(sop)
    print(f"[Init] SOP: {sop['name']}（{len(sop['steps'])} 步驟, {len(sop['few_shot_examples'])} 範例）", flush=True)

    # 背景預載 OCR（用 line_locate.py 的引擎）
    def _preload_ocr():
        try:
            from line_locate import _get_ocr_engine
            _get_ocr_engine()
            print("[Preload] PaddleOCR GPU 已載入", flush=True)
        except Exception as e:
            print(f"[Preload] PaddleOCR 載入失敗：{e}", flush=True)

    ocr_thread = threading.Thread(target=_preload_ocr, daemon=True)
    ocr_thread.start()

    # 搜尋好友 + 進入對話
    regions = enter_conversation(contact_name, monitor)
    if regions is None:
        print("[ERROR] 無法進入對話，中止", flush=True)
        return False

    if should_stop():
        return False

    # 等 OCR 預載完成
    for _ in range(15):
        if should_stop():
            return False
        if not ocr_thread.is_alive():
            break
        time.sleep(1)

    # 建立 conversation_history（完整對話記錄，發的收的都記）
    conversation_history = []

    # 驗證 input_box 在 LINE 視窗範圍內才發歡迎詞
    ib = regions["input_box"]
    ix, iy = ib["center"]
    chat_left = regions["chat_area"].get("left", 0)
    valid_input = (ix != 0 and iy != 0 and
                   ix >= chat_left and
                   ib["bottom"] - ib["top"] > 5)

    if not valid_input:
        print(f"[Init] input_box 座標異常 ({ix},{iy})，跳過歡迎詞避免發到錯誤位置", flush=True)
    else:
        # 發送 SOP 歡迎詞（主動開場）
        print(f"\n[Init] 發送 SOP 歡迎詞...", flush=True)
        welcome_step = None
        ask_area_step = None
        for step in sop["steps"]:
            if step["id"] == "welcome":
                welcome_step = step
            elif step["id"] == "ask_area":
                ask_area_step = step

        if welcome_step:
            for reply in welcome_step.get("replies", []):
                if reply.startswith("{") and reply.endswith("}"):
                    print(f"[Init] 跳過佔位符: {reply}", flush=True)
                    continue
                send_reply(reply, regions)
                conversation_history.append({"text": reply, "sender": "me", "y": 0})
                print(f"[Init] → {reply[:60]}", flush=True)
                time.sleep(1)

        if ask_area_step:
            for reply in ask_area_step.get("replies", []):
                send_reply(reply, regions)
                conversation_history.append({"text": reply, "sender": "me", "y": 0})
                print(f"[Init] → {reply[:60]}", flush=True)
                time.sleep(1)

        print(f"[Init] 歡迎詞已發送（{len(conversation_history)} 條記錄），等待客戶回覆...", flush=True)
        time.sleep(2)

    # 開始監控（傳入 conversation_history）
    print(f"\n[Monitor] 開始自動回覆...", flush=True)
    monitor_and_reply(regions, stop_time, system_prompt, conversation_history, monitor)

    print("\n完成", flush=True)
    return True


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        print("用法：python line_auto_chat.py <好友名稱> <監控時間HH:MM> [SOP設定檔.json]")
        sys.exit(0)

    contact = sys.argv[1]
    stop = sys.argv[2]
    sop = sys.argv[3] if len(sys.argv) > 3 else DEFAULT_SOP
    main(contact, stop, sop)
