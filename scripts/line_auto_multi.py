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
EXCEL_PATH = r"C:\Users\blue_\Desktop\客戶資料.xlsx"

# ============================================================
# Excel 讀寫（客戶資料）
# ============================================================
import openpyxl

def write_customer_to_excel(name, birthday, phone, area, line_name):
    """寫入客戶資料到 Excel（編號先空），回傳行號"""
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    row_num = ws.max_row + 1
    ws.cell(row=row_num, column=1, value="")           # 編號（先空）
    ws.cell(row=row_num, column=2, value=name)          # 姓名
    ws.cell(row=row_num, column=3, value=birthday)      # 出生月日
    ws.cell(row=row_num, column=4, value=phone)         # 電話
    ws.cell(row=row_num, column=5, value=area)          # 地區
    ws.cell(row=row_num, column=6, value=line_name)     # LINE名稱
    ws.cell(row=row_num, column=7, value=datetime.now().strftime("%Y-%m-%d"))  # 報名日期
    wb.save(EXCEL_PATH)
    return row_num

def update_customer_id(row_num, customer_id):
    """補上編號到 Excel"""
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    ws.cell(row=row_num, column=1, value=customer_id)
    wb.save(EXCEL_PATH)

def read_customer_from_excel(customer_id):
    """用編號讀取客戶資料"""
    wb = openpyxl.load_workbook(EXCEL_PATH)
    ws = wb.active
    for row in range(ws.max_row, 1, -1):
        if str(ws.cell(row=row, column=1).value) == str(customer_id):
            return {
                "name": ws.cell(row=row, column=2).value,
                "birthday": ws.cell(row=row, column=3).value,
                "phone": ws.cell(row=row, column=4).value,
                "area": ws.cell(row=row, column=5).value,
                "line_name": ws.cell(row=row, column=6).value,
            }
    return None

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
def find_unread_conversations(monitor=None):
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
def handle_one_customer(conv, regions, system_prompt, sop, all_histories, monitor=None):
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

    # 群組過濾：黑名單（部分匹配） + (數字) 特徵偵測
    SKIP_GROUPS = ["友資群", "友资群", "友資", "友资", "好朋友的群組", "好朋友的群组"]
    if any(g in name or name in g for g in SKIP_GROUPS) or re.search(r"\(\d+\)", name):
        print(f"[Customer] {name} 是群組，跳過", flush=True)
        return regions

    # Step 3: 截圖對話區 + OCR 讀取內容
    chat_img = grab_chat_area(regions, monitor)
    current_messages = ocr_extract_messages(chat_img)
    print(f"[Customer] OCR 讀到 {len(current_messages)} 條訊息", flush=True)

    # 偵測是否為好友（非好友會有 LINE 警告框）
    NOT_FRIEND_KEYWORDS = ["加入好友", "請先確認", "封鎖", "檢舉", "詐騙行為", "用戶可疑"]
    all_text = " ".join(m["text"] for m in current_messages)
    is_friend = not any(kw in all_text for kw in NOT_FRIEND_KEYWORDS)
    print(f"[Customer] 好友狀態: {'是好友' if is_friend else '不是好友'}", flush=True)

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

        # 如果是歡迎詞（包含課程介紹），發送課程圖片
        if "編織" in reply or "歡迎" in reply:
            img_path = sop.get("course_info", {}).get("image", "")
            if img_path:
                if not os.path.isabs(img_path):
                    img_path = str(SCRIPT_DIR.parent / img_path)
                send_image(img_path, regions)
                print(f"[Customer] 已發送課程圖片: {img_path}", flush=True)

        # 回覆加入歷史
        for part in reply.split("|||"):
            part = part.strip()
            if part and not (part.startswith("{") and part.endswith("}")):
                history.append({"text": part, "sender": "me", "y": 0})

        # 偵測到編號 → 提取資料存Excel → 改名 → 分享溫妮 → 讀Excel發友資群
        if "編號" in reply:
            from line_locate import share_contact_card, switch_page, locate_line_regions, rename_friend, CHAT_PAGE_TEMPLATE

            # === Step 2: Claude API 結構化輸出提取報名資料 ===
            # 跟 generate_reply 一樣：送 history + new_them + SOP context
            history_lines = []
            for msg in history[-20:]:
                label = "[客戶]" if msg["sender"] == "them" else "[小編]"
                history_lines.append(f"{label} {msg['text']}")
            history_text = "\n".join(history_lines)
            new_text = "\n".join(f"• {m}" for m in new_them)
            print(f"[Customer] Step2: 新訊息數={len(new_them)}", flush=True)
            extracted = None
            for attempt in range(3):
                try:
                    extract_r = client.messages.create(
                        model="claude-sonnet-4-6",
                        max_tokens=200,
                        messages=[{"role": "user", "content": (
                            f"以下是對話紀錄：\n\n{history_text}\n\n"
                            f"客戶剛發了新訊息：\n{new_text}\n\n"
                            f"從客戶的訊息中提取報名資料。\n\n"
                            f"拆分規則：\n"
                            f"- 中文 = 姓名\n"
                            f"- 4碼數字 = 生日月日（格式 MM/DD，如0910 → 09/10）\n"
                            f"- 09開頭10碼 = 手機號碼\n"
                            f"- 地點相關中文 = 地區\n"
                            f"- 數字可能黏在一起，自己拆分\n"
                            f"- 忽略課程圖片和系統訊息，只看客戶自己打的內容"
                        )}],
                        output_config={
                            "format": {
                                "type": "json_schema",
                                "schema": {
                                    "type": "object",
                                    "properties": {
                                        "customer_name": {"type": "string", "description": "客戶姓名"},
                                        "birthday":      {"type": "string", "description": "出生月日，格式 MM/DD"},
                                        "phone":         {"type": "string", "description": "手機號碼，10碼"},
                                        "area":          {"type": "string", "description": "想參加的地點"}
                                    },
                                    "required": ["customer_name", "birthday", "phone", "area"],
                                    "additionalProperties": False
                                }
                            }
                        }
                    )
                    extracted = json.loads(extract_r.content[0].text)
                    print(f"[Customer] Step2: 資料提取成功（第{attempt+1}次）", flush=True)
                    print(f"[Customer] Step2: {extracted}", flush=True)
                    break
                except Exception as e:
                    print(f"[Customer] Step2: 資料提取失敗（第{attempt+1}次）: {e}", flush=True)
                    if attempt < 2:
                        time.sleep(2)

            # === Step 3: 寫入 Excel（編號先空）===
            if extracted:
                excel_row = write_customer_to_excel(
                    name=extracted.get("customer_name", ""),
                    birthday=extracted.get("birthday", ""),
                    phone=extracted.get("phone", ""),
                    area=extracted.get("area", ""),
                    line_name=name,
                )
                print(f"[Customer] Step3: 已寫入 Excel 第{excel_row}行", flush=True)
            else:
                excel_row = write_customer_to_excel("", "", "", "", name)
                print(f"[Customer] Step3: 提取失敗，寫入空資料到 Excel", flush=True)

            # === Step 4: regex 抓編號 → 補寫 Excel ===
            id_match = re.search(r"編號[：:]\s*\*{0,2}(\d{5})\*{0,2}", reply)
            customer_id = id_match.group(1) if id_match else None
            if customer_id:
                update_customer_id(excel_row, customer_id)
                print(f"[Customer] Step4: 編號 {customer_id} 已寫入 Excel", flush=True)
            else:
                print(f"[Customer] Step4: 抓不到編號", flush=True)

            # === Step 4.5: 如果不是好友，先點「加入好友」===
            if not is_friend:
                from line_locate import ADD_FRIEND_BTN
                ca = regions["chat_area"]
                btn_cx = (ADD_FRIEND_BTN["l"] + ADD_FRIEND_BTN["r"]) // 2
                btn_cy = (ADD_FRIEND_BTN["t"] + ADD_FRIEND_BTN["b"]) // 2
                add_btn_x = ca["left"] + (btn_cx - 371)
                add_btn_y = ca["top"] + (btn_cy - 99)
                pyautogui.click(add_btn_x, add_btn_y)
                print(f"[Customer] Step4.5: 點擊「加入好友」at ({add_btn_x}, {add_btn_y})", flush=True)
                time.sleep(3)

            # === Step 5: rename_friend（把 LINE 名稱改成 日期+編號）===
            # 把客戶當前 LINE 顯示名（name，從 chat_title OCR 抓到）一起傳進去
            # 讓 rename_friend 的窄條 OCR 結果能跟原名比對驗證，避免誤點到狀態訊息（雙重保險）
            rename_name = f"{customer_id} {datetime.now().strftime('%m-%d')}"
            regions = locate_line_regions(monitor)
            rename_friend(regions, rename_name, monitor, current_name=name)
            print(f"[Customer] Step5: 已改名為 {rename_name}", flush=True)
            time.sleep(1)

            search_name = rename_name

            # === Step 6: 分享溫妮好友資訊給客戶 ===
            regions = locate_line_regions(monitor)
            result = share_contact_card(regions, "溫妮", search_name, monitor)
            if result:
                print(f"[Customer] Step6: 已分享溫妮給 {search_name}", flush=True)
            else:
                print(f"[Customer] Step6: 分享溫妮給 {search_name} 失敗", flush=True)

            # === Step 7: 讀 Excel → 組報名資訊 → 發友資群 ===
            print(f"[Customer] Step7: 轉發報名資訊到友資群...", flush=True)
            cust_data = read_customer_from_excel(customer_id)
            if cust_data:
                info = (
                    f"【新報名】\n"
                    f"✏ 姓名：{cust_data['name']}\n"
                    f"✏ 出生月/日：{cust_data['birthday']}\n"
                    f"✏ 聯絡電話：{cust_data['phone']}\n"
                    f"✏ 想參加的地點：{cust_data['area']}\n"
                    f"✏ 編號：{customer_id}"
                )
            else:
                info = f"【新報名】{name}\n編號：{customer_id}"
                print(f"[Customer] Step7: Excel 找不到編號 {customer_id}，使用簡易格式", flush=True)

            # Step 7a: 切到聊天頁（分享溫妮後停在好友頁）
            regions = locate_line_regions(monitor)
            if regions["current_page"] != "chat":
                regions = switch_page(regions, "chat", monitor)

            # Step 7b: 點第一個對話（友資群置頂，永遠在第一個）
            # 用 left_panel 的螢幕座標計算（友資群在 left_panel 頂部）
            # line1_new 圖片座標：友資群中心 x=210, y=140
            # left_panel 圖片座標：l=64, t=111 → 友資群相對 left_panel 頂部偏移 y=140-111=29
            lp = regions["left_panel"]
            first_item_x = (lp["left"] + lp["right"]) // 2
            first_item_y = lp["top"] + 29
            pyautogui.click(first_item_x, first_item_y)
            time.sleep(1.5)
            regions = locate_line_regions(monitor)

            # Step 7d: 發送報名資訊到友資群
            send_reply(info, regions)
            print(f"[Customer] Step7d: 已發送報名資訊到友資群", flush=True)
            time.sleep(1)

            # === Step 8: 分享客戶好友資訊到友資群 ===
            regions = switch_page(regions, "friend", monitor)
            result = share_contact_card(regions, search_name, "友資群", monitor)
            if result:
                print(f"[Customer] Step8: 已分享 {search_name} 的好友資訊到友資群", flush=True)
            else:
                print(f"[Customer] Step8: 分享 {search_name} 的好友資訊到友資群失敗", flush=True)

            # === Step 9: 回到聊天頁 ===
            regions = locate_line_regions(monitor)
            if regions["current_page"] != "chat":
                regions = switch_page(regions, "chat", monitor)
            print(f"[Customer] Step9: 已回到聊天頁", flush=True)

        time.sleep(0.5)

    return regions


# ============================================================
# 主流程
# ============================================================
def main(stop_time, sop_path=DEFAULT_SOP, monitor=None):
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

    # 多開 LINE 設定：要輪流操作的 Sandboxie box 列表（Stage 1 序列化執行）
    # 只列一個 box = 單開模式（向下相容）；多個 box = 序列輪流
    from line_locate import set_active_box, find_line_window
    import win32gui as _w32g
    BOXES = ["blue_1", "blue_2"]  # blue_1=新LINE  blue_2=織夢

    while is_before_stop_time(stop_time):
        if should_stop():
            break

        for box in BOXES:
            if should_stop():
                break
            try:
                # 切到指定 box 的 LINE 視窗
                set_active_box(box)
                line = find_line_window()
                if not line:
                    print(f"[Multi] 找不到 box={box} 的 LINE 視窗，跳過", flush=True)
                    continue
                hwnd = line[0]
                try:
                    _w32g.SetForegroundWindow(hwnd)
                except Exception as e:
                    print(f"[Multi] SetForegroundWindow {box} 失敗: {e}", flush=True)
                time.sleep(1.0)
                print(f"\n[Multi] === 開始處理 box={box} (hwnd={hwnd}) ===", flush=True)

                # Step 1: 切到聊天頁，用定位腳本偵測綠色未讀標記
                regions, unread_list = find_unread_conversations(monitor)

                if not unread_list:
                    print(f"[Multi] box={box} 無未讀，跳下一個", flush=True)
                    continue

                print(f"[Multi] box={box} 偵測到 {len(unread_list)} 個未讀對話", flush=True)

                # Step 2: 逐一處理每個未讀對話
                for conv in unread_list:
                    if should_stop():
                        break

                    regions = handle_one_customer(
                        conv, regions, system_prompt, sop, all_histories, monitor
                    )

                    if should_stop():
                        break

                    # 按 Esc 退出聊天室（不標已讀，對方再回覆會重新出現綠色徽章）
                    pyautogui.press("escape")
                    time.sleep(0.5)

                print(f"[Multi] box={box} 本輪處理完畢", flush=True)

            except Exception as e:
                print(f"[ERR] box={box}: {e}", flush=True)
                import traceback
                traceback.print_exc()
                time.sleep(2)

        # 全部 box 跑完一輪，等待後再從頭
        for _ in range(POLL_INTERVAL):
            if should_stop():
                break
            time.sleep(1)

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
