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
# 優雅停止機制
# ============================================================
STOP_FILE = "C:/Users/blue_/Desktop/測試檔案/.stop_line_auto"
_should_stop = False


def _signal_handler(signum, frame):
    global _should_stop
    _should_stop = True
    print(f"\n[STOP] 收到信號 {signum}，準備停止...", flush=True)


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
# PaddleOCR GPU 全域引擎
# ============================================================
_ocr_engine = None


def _get_ocr_engine():
    global _ocr_engine
    if _ocr_engine is None:
        if "modelscope" not in sys.modules:
            fake = types.ModuleType("modelscope")
            fake.__version__ = "0.0.0"
            sys.modules["modelscope"] = fake
            sys.modules["modelscope.utils"] = types.ModuleType("modelscope.utils")
            sys.modules["modelscope.utils.import_utils"] = types.ModuleType("modelscope.utils.import_utils")
        os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"
        from paddleocr import PaddleOCR
        _ocr_engine = PaddleOCR(
            lang="chinese_cht",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
        )
    return _ocr_engine


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
        f"====== 回覆規則 ======\n"
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
        f"回覆格式：只回覆要發送的文字，不要加任何格式、解釋、或標記。\n"
        f"如果需要分多條訊息發送，用 ||| 分隔，例如：第一條訊息|||第二條訊息|||第三條訊息\n"
    )
    return prompt


# ============================================================
# 對話區截圖 + OCR 訊息提取
# ============================================================
def grab_chat_area(regions, monitor=2):
    """截取 LINE 對話區"""
    from line_locate import screenshot_line
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)

    ca = regions["chat_area"]
    sx_ratio = full_img.size[0] / mon["width"]
    sy_ratio = full_img.size[1] / mon["height"]

    ca_il = int((ca["left"] - mon["left"]) * sx_ratio) - il
    ca_it = int((ca["top"] - mon["top"]) * sy_ratio) - it
    ca_ir = int((ca["right"] - mon["left"]) * sx_ratio) - il
    ca_ib = int((ca["bottom"] - mon["top"]) * sy_ratio) - it

    lw, lh = line_crop.size
    ca_il = max(0, ca_il)
    ca_it = max(0, ca_it)
    ca_ir = min(lw, ca_ir)
    ca_ib = min(lh, ca_ib)

    return line_crop.crop((ca_il, ca_it, ca_ir, ca_ib))


def chat_hash(chat_img):
    """對話區截圖的 hash，偵測畫面變化"""
    small = chat_img.resize((80, 40))
    return hashlib.md5(small.tobytes()).hexdigest()


def ocr_extract_messages(chat_img):
    """
    PaddleOCR 提取對話區文字 + x 位置判斷 sender。
    LINE 的訊息：靠左 = 對方（them），靠右 = 自己（me）
    回傳: [{"text": "...", "sender": "them/me/unknown", "y": int, "conf": float}]
    """
    arr = np.array(chat_img)
    ch, cw = arr.shape[:2]
    mid_x = cw // 2  # 對話區中線

    ocr = _get_ocr_engine()
    results = ocr.predict(arr)

    raw_items = []
    if not results:
        return []

    for item in results:
        texts = item.get("rec_texts", [])
        scores = item.get("rec_scores", [])
        boxes = item.get("dt_polys", [])

        for i, (text, conf) in enumerate(zip(texts, scores)):
            t = text.strip()
            if not t or conf < 0.3:
                continue

            box = boxes[i] if i < len(boxes) else [[0, 0], [0, 0], [0, 0], [0, 0]]
            x = int(box[0][0])
            x_right = int(box[1][0])
            y = int(box[0][1])
            x_center = (x + x_right) // 2

            # LINE 訊息判斷：靠左半邊 = 對方，靠右半邊 = 自己
            if x_center < mid_x * 0.8:
                sender = "them"
            elif x_center > mid_x * 1.2:
                sender = "me"
            else:
                sender = "unknown"

            # 過濾時間戳和系統訊息
            if len(t) < 12 and (":" in t or "上午" in t or "下午" in t):
                continue
            if "已讀" in t or "收回" in t:
                continue

            raw_items.append({"text": t, "sender": sender, "x": x, "y": y, "conf": conf})

    # 合併相鄰同 sender 文字（y 差距 < 40px）
    if not raw_items:
        return []

    messages = []
    current = {"text": raw_items[0]["text"], "sender": raw_items[0]["sender"],
               "y": raw_items[0]["y"], "conf": raw_items[0]["conf"]}

    for item in raw_items[1:]:
        same_sender = item["sender"] == current["sender"]
        close_y = abs(item["y"] - current["y"]) < 40
        if same_sender and close_y:
            current["text"] += " " + item["text"]
            current["conf"] = min(current["conf"], item["conf"])
        else:
            messages.append(current)
            current = {"text": item["text"], "sender": item["sender"],
                       "y": item["y"], "conf": item["conf"]}
    messages.append(current)

    return messages


def detect_new_messages(previous_msgs, current_msgs):
    """比對新舊訊息，找新增的對方訊息"""
    from difflib import SequenceMatcher

    prev_texts = [m["text"] for m in previous_msgs]
    new_them = []

    for msg in current_msgs:
        if msg["sender"] != "them":
            continue
        is_old = False
        for pt in prev_texts:
            ratio = SequenceMatcher(None, msg["text"], pt).ratio()
            if ratio > 0.7:
                is_old = True
                break
        if not is_old:
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
        f"如果客戶提供了完整資料（姓名+生日+電話），請用電話後五碼算出編號。\n"
        f"只回覆要發送的文字，不要加解釋。多條訊息用 ||| 分隔。"
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
# 搜尋好友 + 進入對話
# ============================================================
def enter_conversation(contact_name, monitor=2):
    """搜尋好友並進入對話視窗（複用 line_send_msg 的邏輯）"""
    from line_locate import (
        locate_line_regions, switch_page, screenshot_line,
        find_line_window, ocr_scan_panel,
    )

    # Step 1: 定位 + 置前
    print(f"[Setup] 定位 LINE 視窗...", flush=True)
    line = find_line_window()
    if not line:
        print("[ERROR] 找不到 LINE 視窗", flush=True)
        return None
    win32gui.SetForegroundWindow(line[0])
    time.sleep(0.5)
    regions = locate_line_regions(monitor)

    if should_stop():
        return None

    # Step 2: 切到好友頁
    if regions["current_page"] != "friend":
        regions = switch_page(regions, "friend", monitor)

    if should_stop():
        return None

    # Step 3: 搜尋好友
    print(f"[Setup] 搜尋好友: {contact_name}", flush=True)
    sx, sy = regions["search_bar"]["center"]
    pyautogui.click(sx, sy)
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.press("delete")
    time.sleep(0.1)
    pyperclip.copy(contact_name)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1.5)

    if should_stop():
        return None

    # Step 4: 從搜尋結果找好友並點擊
    print(f"[Setup] 找搜尋結果...", flush=True)
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
    lp = regions["left_panel"]
    sx_ratio = full_img.size[0] / mon["width"]
    sy_ratio = full_img.size[1] / mon["height"]
    lp_il = int((lp["left"] - mon["left"]) * sx_ratio) - il
    lp_it = int((lp["top"] - mon["top"]) * sy_ratio) - it
    lp_ir = int((lp["right"] - mon["left"]) * sx_ratio) - il
    lp_ib = int((lp["bottom"] - mon["top"]) * sy_ratio) - it
    lw, lh = line_crop.size
    lp_il = max(0, lp_il)
    lp_it = max(0, lp_it)
    lp_ir = min(lw, lp_ir)
    lp_ib = min(lh, lp_ib)
    search_crop = line_crop.crop((lp_il, lp_it, lp_ir, lp_ib))
    sw, sh = search_crop.size

    found = False
    try:
        from difflib import SequenceMatcher
        ocr_items = ocr_scan_panel(search_crop)
        for item in ocr_items:
            ratio = SequenceMatcher(None, contact_name, item["text"]).ratio()
            if contact_name in item["text"] or item["text"] in contact_name or ratio > 0.6:
                # 點擊好友項目的中心（名字下方約 25px 才是項目中心，包含頭像和狀態）
                click_x = int(mon["left"] + (il + lp_il + sw // 2) * (mon["width"] / full_img.size[0]))
                click_y = int(mon["top"] + (it + lp_it + item["y"] + 25) * (mon["height"] / full_img.size[1]))
                pyautogui.click(click_x, click_y)
                time.sleep(0.5)
                # 再點一次確保進入對話（LINE 需要點擊好友項目才能開啟對話）
                pyautogui.click(click_x, click_y)
                time.sleep(1.0)
                print(f"[Setup] 找到 {item['text']}，已點擊進入對話", flush=True)
                found = True
                break
    except Exception as e:
        print(f"[Setup] OCR 搜尋失敗: {e}", flush=True)

    if not found:
        print(f"[Setup] OCR 找不到，啟動 Vision...", flush=True)
        tmp = os.path.join(TMPDIR, "line_search_result.png")
        search_crop.save(tmp, quality=95)
        with open(tmp, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()
        r = client.messages.create(
            model="claude-sonnet-4-6", max_tokens=100,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
                {"type": "text", "text": f"LINE search results ({sw}x{sh}px). Find '{contact_name}'. Raw JSON: {{\"x\":0,\"y\":0}}"}
            ]}]
        )
        resp = r.content[0].text.strip()
        if resp.startswith("```"):
            resp = re.sub(r"^```(?:json)?\s*", "", resp)
            resp = re.sub(r"\s*```$", "", resp)
        match = re.search(r"\{.*?\}", resp, re.DOTALL)
        if match:
            pos = json.loads(match.group())
            px = int(pos.get("x", 0) or 0)
            py = int(pos.get("y", 0) or 0)
            click_x = int(mon["left"] + (il + lp_il + px) * (mon["width"] / full_img.size[0]))
            click_y = int(mon["top"] + (it + lp_it + py) * (mon["height"] / full_img.size[1]))
            pyautogui.click(click_x, click_y)
            time.sleep(1.0)
            found = True

    if not found:
        print("[ERROR] 找不到好友", flush=True)
        return None

    # Step 5: 偵測綠色聊天按鈕（新好友）
    print(f"[Setup] 檢查新好友...", flush=True)
    full_img2, line_crop2, (il2, it2, ir2, ib2), mon2 = screenshot_line(monitor)
    arr = np.array(line_crop2)
    lh2, lw2, _ = arr.shape
    sep = regions.get("separator_x", lw2 // 2)
    green_points = []
    for y in range(lh2 // 4, lh2 * 3 // 4):
        for x in range(sep, lw2, 3):
            r, g, b = int(arr[y, x, 0]), int(arr[y, x, 1]), int(arr[y, x, 2])
            if g > 150 and g > r + 50 and g > b + 30 and r < 100:
                green_points.append((x, y))

    if len(green_points) > 20:
        avg_x = sum(p[0] for p in green_points) // len(green_points)
        avg_y = sum(p[1] for p in green_points) // len(green_points)
        sx_r = mon2["width"] / full_img2.size[0]
        sy_r = mon2["height"] / full_img2.size[1]
        btn_x = int(mon2["left"] + (il2 + avg_x) * sx_r)
        btn_y = int(mon2["top"] + (it2 + avg_y) * sy_r)
        print(f"[Setup] 新好友，點擊綠色聊天按鈕 ({btn_x}, {btn_y})", flush=True)
        pyautogui.click(btn_x, btn_y)
        time.sleep(1.5)

    # 重新定位（進入對話後）
    regions = locate_line_regions(monitor)

    # 驗證 chat_area 和 input_box 座標是否合理
    ca = regions["chat_area"]
    ib = regions["input_box"]
    ca_height = ca["bottom"] - ca["top"]
    ib_height = ib["bottom"] - ib["top"]

    if ca_height < 50 or ib_height < 10:
        # Vision 沒正確找到對話區，用像素分析的分隔線 + 視窗邊界重新計算
        print(f"[Setup] 對話區座標異常（h={ca_height}），用備援方式定位...", flush=True)
        lw_info = regions["line_window"]
        sep = regions.get("separator_x", 200)

        from line_locate import find_line_window, screenshot_line
        _, line_crop, (il, it, ir, ib_img), mon = screenshot_line(monitor)
        lw_px, lh_px = line_crop.size
        sx_r = mon["width"] / (ir - il + lw_px) * (lw_px / (ir - il)) if ir > il else 1
        sy_r = mon["height"] / (ib_img - it + lh_px) * (lh_px / (ib_img - it)) if ib_img > it else 1

        # chat_area: 從分隔線到右邊界，上方 15% 到下方 85%
        regions["chat_area"] = {
            "left": int(mon["left"] + (il + sep) * (mon["width"] / (ir - il + 1 if ir > il else lw_px))),
            "top": int(mon["top"] + (it + int(lh_px * 0.15)) * (mon["height"] / (ib_img - it + 1 if ib_img > it else lh_px))),
            "right": int(mon["left"] + ir * (mon["width"] / (ir - il + 1 if ir > il else lw_px))),
            "bottom": int(mon["top"] + (it + int(lh_px * 0.88)) * (mon["height"] / (ib_img - it + 1 if ib_img > it else lh_px))),
            "center": (0, 0),
        }
        ca = regions["chat_area"]
        ca["center"] = ((ca["left"] + ca["right"]) // 2, (ca["top"] + ca["bottom"]) // 2)

        # input_box: 對話區下方
        regions["input_box"] = {
            "left": ca["left"],
            "top": ca["bottom"],
            "right": ca["right"],
            "bottom": int(mon["top"] + (it + int(lh_px * 0.95)) * (mon["height"] / (ib_img - it + 1 if ib_img > it else lh_px))),
            "center": (0, 0),
        }
        ib = regions["input_box"]
        ib["center"] = ((ib["left"] + ib["right"]) // 2, (ib["top"] + ib["bottom"]) // 2)

        print(f"[Setup] 備援 chat_area: {ca['left']},{ca['top']} → {ca['right']},{ca['bottom']}", flush=True)
        print(f"[Setup] 備援 input_box center: {ib['center']}", flush=True)

    print(f"[Setup] 已進入對話視窗", flush=True)
    return regions


# ============================================================
# 監控 + 自動回覆主迴圈
# ============================================================
COOLDOWN_SECONDS = 15
POLL_INTERVAL = 5
WAIT_COMPLETE = 2
WAIT_COMPLETE_ROUNDS = 2


def monitor_and_reply(regions, stop_time, system_prompt, monitor=2):
    """監控對話區，偵測新訊息後自動回覆"""

    # 初始 OCR 掃描
    chat = grab_chat_area(regions, monitor)
    previous_messages = ocr_extract_messages(chat)
    last_hash = chat_hash(chat)
    cooldown_until = 0
    conversation_history = list(previous_messages)

    print(f"[Monitor] 開始監控 → {stop_time}", flush=True)
    print(f"[Monitor] 冷卻：{COOLDOWN_SECONDS}s / 輪詢：{POLL_INTERVAL}s", flush=True)
    print(f"[Monitor] 初始訊息數：{len(previous_messages)}", flush=True)
    print(f"[Monitor] 停止方式：touch {STOP_FILE}", flush=True)

    while datetime.now().strftime("%H:%M") < stop_time:
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

            # 等對方說完（最多 2 輪 x 2 秒）
            latest_chat = chat
            latest_hash = h
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
            new_them_msgs = detect_new_messages(previous_messages, current_messages)

            if not new_them_msgs:
                last_hash = latest_hash
                previous_messages = current_messages
                conversation_history = list(current_messages)
                continue

            print(f"[{t}] 偵測到 {len(new_them_msgs)} 條新訊息：{new_them_msgs}", flush=True)

            # Claude AI 生成回覆
            reply = generate_reply(system_prompt, conversation_history, new_them_msgs)

            if reply and len(reply) > 1:
                print(f"[{t}] AI 回覆：{reply[:100]}", flush=True)

                # 點輸入框發送
                send_multi_reply(reply, regions)

                cooldown_until = time.time() + COOLDOWN_SECONDS
                print(f"[{t}] 冷卻 {COOLDOWN_SECONDS} 秒", flush=True)

                # 冷卻期間每秒檢查停止
                for _ in range(COOLDOWN_SECONDS):
                    if should_stop():
                        break
                    time.sleep(1)

                # 更新追蹤
                chat = grab_chat_area(regions, monitor)
                last_hash = chat_hash(chat)
                previous_messages = ocr_extract_messages(chat)
                conversation_history = list(previous_messages)
            else:
                print(f"[{t}] 回覆為空，跳過", flush=True)
                last_hash = latest_hash
                previous_messages = current_messages
                conversation_history = list(current_messages)

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

    # 背景預載 OCR
    def _preload_ocr():
        try:
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

    # 開始監控
    print(f"\n[Monitor] 開始自動回覆...", flush=True)
    monitor_and_reply(regions, stop_time, system_prompt, monitor)

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
