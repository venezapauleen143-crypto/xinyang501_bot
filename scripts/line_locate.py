"""
LINE 電腦版 UI 定位模組（方案 E：模板比例 + PaddleOCR + Vision 三層架構）

這是所有 LINE 自動化腳本的基礎，負責：
1. 精確定位 LINE 視窗所有 UI 區域
2. 辨識目前在哪個頁面（好友/聊天/加好友）
3. 掃描該頁面的完整內容物和座標
4. 切換頁面後重新掃描新頁面的內容

三層定位架構：
- 第一層（模板比例）：固定元素用預算的相對比例定位（0 秒）
- 第二層（PaddleOCR）：動態文字內容用 GPU OCR 提取（0.4 秒）
- 第三層（Claude Vision）：複雜判斷或 OCR 失敗時的備援（2-4 秒）

使用方式：
    from line_locate import locate_line_regions, switch_page, scan_page_content

    # 定位 + 掃描當前頁面
    regions = locate_line_regions(monitor=2)

    # 框架座標
    regions["sidebar_icons"]["btn_chat"]["center"]  → 聊天頁按鈕
    regions["input_box"]["center"]                  → 輸入框
    regions["chat_area"]                            → 對話區

    # 當前頁面內容
    regions["page_content"]                         → 該頁面的所有項目和座標

    # 切換頁面（自動重新掃描）
    regions = switch_page(regions, "chat", monitor=2)
    regions["page_content"]                         → 聊天頁的所有對話列表
"""

import sys
import io
import os
import re
import json
import time
import base64
import types
import ctypes
import numpy as np

# DPI 縮放設定（必須在所有 GUI 操作之前）
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(0)
except Exception:
    pass

if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")


def _print(msg):
    """安全 print"""
    try:
        print(msg, flush=True)
    except (ValueError, UnicodeEncodeError):
        try:
            sys.stderr.write(msg + "\n")
            sys.stderr.flush()
        except Exception:
            pass


import win32gui
import mss
import pyautogui
from PIL import Image
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"


# ============================================================
# 模板比例常數（從三張參考截圖分析得出）
# 所有數值都是相對於 left_panel 寬度/高度的比例
# ============================================================

# 好友頁（line1.png）— left_panel 內的固定元素相對位置
FRIEND_PAGE_TEMPLATE = {
    # 搜尋欄（相對 left_panel）
    "search_bar": {"y_ratio": 0.04, "h_ratio": 0.03},
    # 用戶個人資料區（搜尋欄下方）
    "user_profile": {"y_ratio": 0.07, "h_ratio": 0.07},
    # 社群區塊標題
    "section_community": {"y_ratio": 0.16, "label": "社群"},
    # 群組區塊標題
    "section_group": {"y_ratio": 0.25, "label": "群組"},
    # 好友區塊標題
    "section_friend": {"y_ratio": 0.36, "label": "好友"},
    # 每個列表項目的行高比例
    "item_height_ratio": 0.065,
}

# 聊天頁（line2.png）— left_panel 內的固定元素相對位置
CHAT_PAGE_TEMPLATE = {
    # 分頁標籤列（全部/群組/社群/篩選）
    "tab_bar": {"y_ratio": 0.0, "h_ratio": 0.025},
    # 分頁標籤的 x 比例位置（相對 left_panel 寬度）
    "tab_all": {"x_ratio": 0.12, "label": "全部"},
    "tab_group": {"x_ratio": 0.32, "label": "群組"},
    "tab_community": {"x_ratio": 0.52, "label": "社群"},
    "tab_filter": {"x_ratio": 0.72, "label": "篩選"},
    # 搜尋欄
    "search_bar": {"y_ratio": 0.04, "h_ratio": 0.03},
    # 聊天列表起始位置
    "list_start_y_ratio": 0.08,
    # 每個對話的行高比例
    "item_height_ratio": 0.075,
}

# 加好友頁（line3.png）— left_panel 內的固定元素相對位置
ADD_FRIEND_PAGE_TEMPLATE = {
    # 三個功能按鈕（相對 left_panel）
    "btn_search_friend": {"y_ratio": 0.04, "h_ratio": 0.045, "label": "搜尋好友"},
    "btn_create_group": {"y_ratio": 0.09, "h_ratio": 0.045, "label": "建立群組"},
    "btn_create_community": {"y_ratio": 0.14, "h_ratio": 0.045, "label": "建立社群"},
    # 官方推薦區塊
    "section_official": {"y_ratio": 0.20, "label": "官方推薦"},
    # 你可能認識的人區塊
    "section_maybe_know": {"y_ratio": 0.35, "label": "您可能認識的人"},
    # 每個推薦項目的行高比例
    "item_height_ratio": 0.06,
}


# ============================================================
# PaddleOCR GPU 全域引擎（只載入一次）
# ============================================================
_ocr_engine = None


def _get_ocr_engine():
    """取得全域 PaddleOCR 引擎"""
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
# 視窗偵測
# ============================================================
def find_line_window():
    """用 win32gui 找到 LINE 主視窗"""
    results = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "LINE" in title:
                rect = win32gui.GetWindowRect(hwnd)
                w = rect[2] - rect[0]
                h = rect[3] - rect[1]
                if w > 300 and h > 300:
                    cls = win32gui.GetClassName(hwnd)
                    results.append((hwnd, title, cls, rect))
    win32gui.EnumWindows(callback, results)
    if not results:
        return None
    return max(results, key=lambda x: (x[3][2] - x[3][0]) * (x[3][3] - x[3][1]))


def screenshot_line(monitor=2):
    """截圖指定螢幕，裁切出 LINE 區域"""
    line = find_line_window()
    if not line:
        raise RuntimeError("找不到 LINE 視窗")

    hwnd, title, cls, (wl, wt, wr, wb) = line

    with mss.mss() as sct:
        mon = sct.monitors[monitor]
        img = sct.grab(mon)
        pil = Image.frombytes("RGB", img.size, img.rgb)
        iw, ih = pil.size

    sx = iw / mon["width"]
    sy = ih / mon["height"]
    il = int((wl - mon["left"]) * sx)
    it = int((wt - mon["top"]) * sy)
    ir = int((wr - mon["left"]) * sx)
    ib = int((wb - mon["top"]) * sy)

    il = max(0, il)
    it = max(0, it)
    ir = min(iw, ir)
    ib = min(ih, ib)

    line_crop = pil.crop((il, it, ir, ib))
    return pil, line_crop, (il, it, ir, ib), mon


# ============================================================
# 像素分析
# ============================================================
def find_sidebar_by_pixel(line_crop):
    """像素分析找最左側功能列的寬度"""
    arr = np.array(line_crop)
    h, w, _ = arr.shape
    for x in range(30, min(80, w)):
        col = arr[h // 4: h * 3 // 4, x, :]
        prev_col = arr[h // 4: h * 3 // 4, x - 1, :]
        if np.mean(col) - np.mean(prev_col) > 30:
            return x
    return 50


def find_separator_by_pixel(line_crop, sidebar_width):
    """像素分析找左側面板和對話區的分隔線"""
    arr = np.array(line_crop)
    h, w, _ = arr.shape
    candidates = []
    for check_y in range(int(h * 0.3), int(h * 0.7), int(h * 0.05)):
        row = arr[check_y, :, :]
        for x in range(sidebar_width + 50, w * 2 // 3):
            brightness = (int(row[x, 0]) + int(row[x, 1]) + int(row[x, 2])) / 3
            prev_brightness = (int(row[x - 1, 0]) + int(row[x - 1, 1]) + int(row[x - 1, 2])) / 3
            if abs(brightness - prev_brightness) > 20 and 180 < brightness < 220:
                candidates.append(x)
                break
    if candidates:
        return sorted(candidates)[len(candidates) // 2]
    return w // 3


# ============================================================
# 第一層-A：像素分析定位 sidebar 按鈕（不靠 Vision）
# ============================================================
# 從 line4.png 參考圖精確量測的比例常數
# sidebar 背景色 ≈ RGB(42,51,74)，圖示比背景亮
# 按鈕間距約 50-55px（在 ~984px 高的視窗中）
SIDEBAR_BUTTONS = {
    # y_ratio = 按鈕中心 y / 視窗高度（從像素分析精確量測）
    "btn_friend":       {"y_ratio": 0.053, "label": "好友"},       # y≈52
    "btn_chat":         {"y_ratio": 0.107, "label": "聊天"},       # y≈105
    "btn_add_friend":   {"y_ratio": 0.165, "label": "加好友"},     # y≈162
    "btn_timeline":     {"y_ratio": 0.219, "label": "動態消息"},   # y≈215
    "btn_call":         {"y_ratio": 0.271, "label": "通話"},       # y≈267
    "btn_keep":         {"y_ratio": 0.335, "label": "Keep"},       # y≈330
    "btn_line_official":{"y_ratio": 0.376, "label": "LINE官方"},   # y≈370
    "btn_more":         {"y_ratio": 0.965, "label": "更多"},       # y≈950
}


def calc_sidebar_icons(sidebar_width, line_height, il, it, mon, full_img_size):
    """用模板比例算出 sidebar 每個按鈕的螢幕絕對座標"""
    sx_ratio = mon["width"] / full_img_size[0]
    sy_ratio = mon["height"] / full_img_size[1]

    x_center = sidebar_width // 2  # 按鈕 x 中心 = sidebar 寬度的一半

    icons = {}
    for btn_name, btn_info in SIDEBAR_BUTTONS.items():
        local_y = int(line_height * btn_info["y_ratio"])
        abs_x = int(mon["left"] + (il + x_center) * sx_ratio)
        abs_y = int(mon["top"] + (it + local_y) * sy_ratio)
        icons[btn_name] = {"center": (abs_x, abs_y), "label": btn_info["label"]}

    return icons


# ============================================================
# 第一層-B：Vision 定位框架（不含 sidebar 按鈕）
# ============================================================
def find_framework_by_vision(line_crop):
    """Claude Vision 定位框架結構 + sidebar 每個按鈕 + 判斷當前頁面"""
    import anthropic
    client = anthropic.Anthropic()

    lw, lh = line_crop.size
    tmp = os.path.join(TMPDIR, "line_vision_locate.png")
    line_crop.save(tmp, quality=95)
    with open(tmp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    prompt = (
        f"Image size: {lw}x{lh} pixels. This is LINE Desktop app.\n\n"
        "Find EXACT pixel coordinates for these regions. DO NOT locate sidebar icons.\n\n"
        "=== MAIN UI REGIONS (bounding boxes l,t,r,b) ===\n"
        "1. search_bar: Search input field in the left panel area. Only the text input.\n"
        "2. left_panel: Panel between the narrow icon sidebar and chat area.\n"
        "3. chat_title: Name text at top of chat area (right side). Just the name.\n"
        "4. chat_area: Message area with conversation bubbles.\n"
        "5. input_box: Text input field at bottom right. Only the text area.\n\n"
        "=== CURRENT PAGE ===\n"
        "Look at the LEFT PANEL content (NOT the sidebar icons):\n"
        '"friend" if it shows categorized lists with headers like 社群/群組/好友\n'
        '"chat" if it shows a conversation list with tabs 全部/群組/社群 at top\n'
        '"add_friend" if it shows buttons like 搜尋好友/建立群組/建立社群\n\n'
        "Return RAW JSON only:\n"
        '{"search_bar":{"l":0,"t":0,"r":0,"b":0},'
        '"left_panel":{"l":0,"t":0,"r":0,"b":0},'
        '"chat_title":{"l":0,"t":0,"r":0,"b":0},'
        '"chat_area":{"l":0,"t":0,"r":0,"b":0},'
        '"input_box":{"l":0,"t":0,"r":0,"b":0},'
        '"current_page":"friend or chat or add_friend"}'
    )

    r = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=800,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
            {"type": "text", "text": prompt},
        ]}],
    )
    resp = r.content[0].text.strip()
    if resp.startswith("```"):
        resp = re.sub(r"^```(?:json)?\s*", "", resp)
        resp = re.sub(r"\s*```$", "", resp)
    return json.loads(resp)


# ============================================================
# 第二層：PaddleOCR 掃描頁面內容
# ============================================================
def ocr_scan_panel(panel_img):
    """
    用 PaddleOCR 掃描 left_panel 區域，提取所有文字和座標。
    回傳: [{"text": "...", "x": int, "y": int, "w": int, "h": int, "conf": float}]
    """
    arr = np.array(panel_img)
    ocr = _get_ocr_engine()
    results = ocr.predict(arr)

    items = []
    if not results:
        return items

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
            y = int(box[0][1])
            w = int(box[1][0] - box[0][0])
            h = int(box[2][1] - box[0][1])
            items.append({"text": t, "x": x, "y": y, "w": w, "h": h, "conf": conf})

    # 依 y 座標排序
    items.sort(key=lambda it: it["y"])
    return items


# ============================================================
# 第二層 + 第一層混合：掃描各頁面的完整內容
# ============================================================
def scan_friend_page(panel_img, panel_rect, mon, full_img_size, il, it):
    """掃描好友頁的完整內容"""
    pw, ph = panel_img.size
    sx_ratio = mon["width"] / full_img_size[0]
    sy_ratio = mon["height"] / full_img_size[1]

    def to_abs(local_x, local_y):
        ax = int(mon["left"] + (il + panel_rect["l"] + local_x) * sx_ratio)
        ay = int(mon["top"] + (it + panel_rect["t"] + local_y) * sy_ratio)
        return (ax, ay)

    # OCR 掃描
    ocr_items = ocr_scan_panel(panel_img)

    # 用模板比例算出固定元素的大致位置
    template = FRIEND_PAGE_TEMPLATE
    section_y = {
        "community": int(ph * template["section_community"]["y_ratio"]),
        "group": int(ph * template["section_group"]["y_ratio"]),
        "friend": int(ph * template["section_friend"]["y_ratio"]),
    }

    # 分類 OCR 結果到各區塊
    communities = []
    groups = []
    friends = []
    user_name = ""

    for item in ocr_items:
        y = item["y"]
        text = item["text"]

        # 用戶名稱（最上方）
        if y < section_y["community"] and not user_name and len(text) > 1:
            if "搜尋" not in text and "Keep" not in text and "姓名" not in text:
                user_name = text

        # 社群區塊的項目
        elif section_y["community"] <= y < section_y["group"]:
            if "社群" not in text and len(text) > 1:
                communities.append({"name": text, "center": to_abs(pw // 2, y + 15)})

        # 群組區塊的項目
        elif section_y["group"] <= y < section_y["friend"]:
            if "群組" not in text and len(text) > 1:
                groups.append({"name": text, "center": to_abs(pw // 2, y + 15)})

        # 好友區塊的項目
        elif y >= section_y["friend"]:
            if "好友" not in text and len(text) > 1:
                friends.append({"name": text, "center": to_abs(pw // 2, y + 15)})

    return {
        "page": "friend",
        "user_name": user_name,
        "communities": communities,
        "groups": groups,
        "friends": friends,
        "fixed_elements": {
            "search_bar": {"center": to_abs(pw // 2, int(ph * template["search_bar"]["y_ratio"]))},
            "user_profile": {"center": to_abs(pw // 2, int(ph * template["user_profile"]["y_ratio"]))},
            "section_community_header": {"center": to_abs(pw // 4, section_y["community"])},
            "section_group_header": {"center": to_abs(pw // 4, section_y["group"])},
            "section_friend_header": {"center": to_abs(pw // 4, section_y["friend"])},
        },
        "all_ocr_items": ocr_items,
    }


def scan_chat_page(panel_img, panel_rect, mon, full_img_size, il, it):
    """掃描聊天頁的完整內容"""
    pw, ph = panel_img.size
    sx_ratio = mon["width"] / full_img_size[0]
    sy_ratio = mon["height"] / full_img_size[1]

    def to_abs(local_x, local_y):
        ax = int(mon["left"] + (il + panel_rect["l"] + local_x) * sx_ratio)
        ay = int(mon["top"] + (it + panel_rect["t"] + local_y) * sy_ratio)
        return (ax, ay)

    template = CHAT_PAGE_TEMPLATE

    # OCR 掃描
    ocr_items = ocr_scan_panel(panel_img)

    # 分頁標籤的絕對座標（用模板比例算）
    tabs = {
        "tab_all": {"label": "全部", "center": to_abs(int(pw * template["tab_all"]["x_ratio"]), int(ph * template["tab_bar"]["y_ratio"] + 12))},
        "tab_group": {"label": "群組", "center": to_abs(int(pw * template["tab_group"]["x_ratio"]), int(ph * template["tab_bar"]["y_ratio"] + 12))},
        "tab_community": {"label": "社群", "center": to_abs(int(pw * template["tab_community"]["x_ratio"]), int(ph * template["tab_bar"]["y_ratio"] + 12))},
        "tab_filter": {"label": "篩選", "center": to_abs(int(pw * template["tab_filter"]["x_ratio"]), int(ph * template["tab_bar"]["y_ratio"] + 12))},
    }

    # 聊天列表（OCR 讀到的對話項目）
    list_start_y = int(ph * template["list_start_y_ratio"])
    conversations = []
    skip_keywords = ["全部", "群組", "社群", "篩選", "搜尋", "聊天", "訊息", "Keep"]

    for item in ocr_items:
        if item["y"] < list_start_y:
            continue
        text = item["text"]
        if any(kw in text for kw in skip_keywords) or len(text) < 2:
            continue

        # 解析未讀數（數字在名稱旁邊）
        unread = 0
        name = text
        # 匹配 "名稱 (數字)" 格式
        m = re.match(r"(.+?)\s*\((\d+)\)\s*$", text)
        if m:
            name = m.group(1)

        # 檢查是否有獨立的數字（未讀數）
        for other in ocr_items:
            if abs(other["y"] - item["y"]) < 15 and other["text"].isdigit():
                unread = int(other["text"])

        conversations.append({
            "name": name,
            "unread": unread,
            "center": to_abs(pw // 2, item["y"] + 15),
            "y": item["y"],
        })

    # 去重（名稱相同的只保留第一個）
    seen = set()
    unique_convs = []
    for c in conversations:
        if c["name"] not in seen:
            seen.add(c["name"])
            unique_convs.append(c)

    return {
        "page": "chat",
        "tabs": tabs,
        "conversations": unique_convs,
        "fixed_elements": {
            "search_bar": {"center": to_abs(pw // 2, int(ph * template["search_bar"]["y_ratio"]))},
        },
        "all_ocr_items": ocr_items,
    }


def scan_add_friend_page(panel_img, panel_rect, mon, full_img_size, il, it):
    """掃描加好友頁的完整內容"""
    pw, ph = panel_img.size
    sx_ratio = mon["width"] / full_img_size[0]
    sy_ratio = mon["height"] / full_img_size[1]

    def to_abs(local_x, local_y):
        ax = int(mon["left"] + (il + panel_rect["l"] + local_x) * sx_ratio)
        ay = int(mon["top"] + (it + panel_rect["t"] + local_y) * sy_ratio)
        return (ax, ay)

    template = ADD_FRIEND_PAGE_TEMPLATE

    # OCR 掃描
    ocr_items = ocr_scan_panel(panel_img)

    # 固定按鈕（用模板比例）
    buttons = {
        "btn_search_friend": {
            "label": "搜尋好友",
            "center": to_abs(pw // 2, int(ph * template["btn_search_friend"]["y_ratio"] + ph * template["btn_search_friend"]["h_ratio"] / 2)),
        },
        "btn_create_group": {
            "label": "建立群組",
            "center": to_abs(pw // 2, int(ph * template["btn_create_group"]["y_ratio"] + ph * template["btn_create_group"]["h_ratio"] / 2)),
        },
        "btn_create_community": {
            "label": "建立社群",
            "center": to_abs(pw // 2, int(ph * template["btn_create_community"]["y_ratio"] + ph * template["btn_create_community"]["h_ratio"] / 2)),
        },
    }

    # 用 OCR 修正按鈕位置（如果 OCR 找到了文字）
    for item in ocr_items:
        for btn_key, btn in buttons.items():
            if btn["label"] in item["text"]:
                buttons[btn_key]["center"] = to_abs(pw // 2, item["y"] + 10)

    # 官方推薦列表
    official_y = int(ph * template["section_official"]["y_ratio"])
    maybe_know_y = int(ph * template["section_maybe_know"]["y_ratio"])

    official_accounts = []
    maybe_know_people = []
    skip_keywords = ["搜尋", "建立", "好友", "群組", "社群", "官方推薦", "您可能認識"]

    for item in ocr_items:
        text = item["text"]
        if any(kw in text for kw in skip_keywords) or len(text) < 2:
            continue

        if official_y <= item["y"] < maybe_know_y:
            official_accounts.append({
                "name": text,
                "center": to_abs(pw // 2, item["y"] + 10),
            })
        elif item["y"] >= maybe_know_y:
            maybe_know_people.append({
                "name": text,
                "center": to_abs(pw // 2, item["y"] + 10),
            })

    return {
        "page": "add_friend",
        "buttons": buttons,
        "official_accounts": official_accounts,
        "maybe_know_people": maybe_know_people,
        "all_ocr_items": ocr_items,
    }


# ============================================================
# 第三層：Vision 備援（OCR 失敗時使用）
# ============================================================
def vision_scan_panel(panel_img, current_page):
    """OCR 失敗時用 Claude Vision 掃描頁面內容"""
    import anthropic
    client = anthropic.Anthropic()

    pw, ph = panel_img.size
    tmp = os.path.join(TMPDIR, "line_panel_scan.png")
    panel_img.save(tmp, quality=95)
    with open(tmp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    if current_page == "chat":
        detail = (
            "List ALL conversations visible. For each: name, unread count (0 if none), "
            "and center Y coordinate in pixels."
        )
    elif current_page == "friend":
        detail = (
            "List the user's name at top, then all items under 社群/群組/好友 sections. "
            "For each item: name, which section it belongs to, and center Y coordinate."
        )
    else:
        detail = (
            "List all buttons (搜尋好友/建立群組/建立社群) and their Y coordinates. "
            "List all recommended accounts and people you may know with their Y coordinates."
        )

    prompt = (
        f"Image size: {pw}x{ph}. This is LINE's left panel showing '{current_page}' page.\n"
        f"{detail}\n"
        "Return RAW JSON array of items: "
        '[{"name":"...","type":"conversation/community/group/friend/button/official/maybe_know",'
        '"y":0,"unread":0}]'
    )

    r = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=600,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
            {"type": "text", "text": prompt},
        ]}],
    )
    resp = r.content[0].text.strip()
    if resp.startswith("```"):
        resp = re.sub(r"^```(?:json)?\s*", "", resp)
        resp = re.sub(r"\s*```$", "", resp)
    try:
        return json.loads(resp)
    except Exception:
        return []


# ============================================================
# 主函數
# ============================================================
def locate_line_regions(monitor=2):
    """
    主函數：定位 LINE 所有 UI 區域 + 掃描當前頁面完整內容

    返回：
    {
        # 框架座標（螢幕絕對座標）
        "sidebar":        {"left", "top", "right", "bottom", "center"},
        "sidebar_icons":  {"btn_friend": {"center"}, "btn_chat": {"center"}, ...},
        "search_bar":     {"left", "top", "right", "bottom", "center"},
        "left_panel":     {"left", "top", "right", "bottom", "center"},
        "chat_title":     {"left", "top", "right", "bottom", "center"},
        "chat_area":      {"left", "top", "right", "bottom", "center"},
        "input_box":      {"left", "top", "right", "bottom", "center"},

        # 頁面資訊
        "current_page":   "friend" | "chat" | "add_friend",
        "page_content":   { ... 該頁面的完整內容和座標 ... },

        # 輔助資訊
        "separator_x":    int,
        "sidebar_width":  int,
        "line_window":    {"left", "top", "right", "bottom"},
    }
    """
    _print("[line_locate] 開始定位...")

    # Step 1: 截圖
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
    lw, lh = line_crop.size
    _print(f"[line_locate] LINE: {lw}x{lh} at ({il},{it})-({ir},{ib})")

    # Step 2: 像素分析
    sidebar_w = find_sidebar_by_pixel(line_crop)
    sep_x = find_separator_by_pixel(line_crop, sidebar_w)
    _print(f"[line_locate] sidebar: {sidebar_w}px, 分隔線: x={sep_x}")

    # Step 3: Vision 定位框架 + sidebar 按鈕
    vision = find_framework_by_vision(line_crop)
    current_page = vision.get("current_page", "unknown")
    _print(f"[line_locate] 當前頁面: {current_page}")

    # 座標轉換函數
    sx_ratio = mon["width"] / full_img.size[0]
    sy_ratio = mon["height"] / full_img.size[1]

    def to_screen_region(v):
        sl = int(mon["left"] + (il + v["l"]) * sx_ratio)
        st = int(mon["top"] + (it + v["t"]) * sy_ratio)
        sr = int(mon["left"] + (il + v["r"]) * sx_ratio)
        sb = int(mon["top"] + (it + v["b"]) * sy_ratio)
        return {"left": sl, "top": st, "right": sr, "bottom": sb, "center": ((sl + sr) // 2, (st + sb) // 2)}

    def to_screen_point(v):
        sx_abs = int(mon["left"] + (il + v["x"]) * sx_ratio)
        sy_abs = int(mon["top"] + (it + v["y"]) * sy_ratio)
        return {"center": (sx_abs, sy_abs)}

    # 組合框架結果
    sidebar_raw = {"l": 0, "t": 0, "r": sidebar_w, "b": lh}
    sidebar = to_screen_region(sidebar_raw)
    search_bar = to_screen_region(vision["search_bar"])
    left_panel = to_screen_region(vision["left_panel"])
    chat_title = to_screen_region(vision["chat_title"])
    input_box = to_screen_region(vision["input_box"])

    chat_area_raw = {
        "l": sep_x, "t": vision["chat_title"]["b"],
        "r": lw, "b": vision["input_box"]["t"],
    }
    chat_area = to_screen_region(chat_area_raw)

    # sidebar 按鈕（用像素分析 + 模板比例，不靠 Vision）
    sidebar_icons = calc_sidebar_icons(sidebar_w, lh, il, it, mon, full_img.size)

    # Step 4: 掃描當前頁面內容（PaddleOCR 主力）
    _print(f"[line_locate] 掃描 {current_page} 頁面內容...")
    panel_rect = vision["left_panel"]
    panel_crop = line_crop.crop((panel_rect["l"], panel_rect["t"], panel_rect["r"], panel_rect["b"]))

    page_content = {}
    try:
        if current_page == "friend":
            page_content = scan_friend_page(panel_crop, panel_rect, mon, full_img.size, il, it)
        elif current_page == "chat":
            page_content = scan_chat_page(panel_crop, panel_rect, mon, full_img.size, il, it)
        elif current_page == "add_friend":
            page_content = scan_add_friend_page(panel_crop, panel_rect, mon, full_img.size, il, it)

        # 檢查 OCR 結果是否足夠（太少項目就啟動 Vision 備援）
        ocr_count = len(page_content.get("all_ocr_items", []))
        if ocr_count < 3:
            _print(f"[line_locate] OCR 只找到 {ocr_count} 項，啟動 Vision 備援...")
            vision_items = vision_scan_panel(panel_crop, current_page)
            page_content["vision_fallback"] = vision_items
    except Exception as e:
        _print(f"[line_locate] 頁面掃描失敗：{e}，啟動 Vision 備援...")
        vision_items = vision_scan_panel(panel_crop, current_page)
        page_content = {"page": current_page, "vision_fallback": vision_items}

    # 組合最終結果
    result = {
        "sidebar": sidebar,
        "sidebar_icons": sidebar_icons,
        "search_bar": search_bar,
        "left_panel": left_panel,
        "chat_title": chat_title,
        "chat_area": chat_area,
        "input_box": input_box,
        "current_page": current_page,
        "page_content": page_content,
        "separator_x": sep_x,
        "sidebar_width": sidebar_w,
        "line_window": {"left": il, "top": it, "right": ir, "bottom": ib},
    }

    # 輸出結果
    _print("[line_locate] 定位完成")
    for name in ["sidebar", "search_bar", "left_panel", "chat_title", "chat_area", "input_box"]:
        r = result[name]
        _print(f"  {name}: ({r['left']},{r['top']})-({r['right']},{r['bottom']}) center={r['center']}")
    _print(f"  current_page: {current_page}")
    _print(f"  sidebar_icons:")
    for btn, info in sidebar_icons.items():
        _print(f"    {btn}: {info['center']}")

    # 輸出頁面內容摘要
    if current_page == "friend":
        _print(f"  page_content: user={page_content.get('user_name','')}, "
               f"communities={len(page_content.get('communities',[]))}, "
               f"groups={len(page_content.get('groups',[]))}, "
               f"friends={len(page_content.get('friends',[]))}")
    elif current_page == "chat":
        convs = page_content.get("conversations", [])
        _print(f"  page_content: {len(convs)} conversations")
        for c in convs[:5]:
            _print(f"    - {c['name']} (unread={c.get('unread',0)})")
    elif current_page == "add_friend":
        _print(f"  page_content: buttons={list(page_content.get('buttons',{}).keys())}, "
               f"official={len(page_content.get('official_accounts',[]))}, "
               f"maybe_know={len(page_content.get('maybe_know_people',[]))}")

    return result


def switch_page(regions, target_page, monitor=2):
    """
    切換 LINE 左側面板到指定頁面，自動重新定位 + 掃描新頁面內容

    參數：
        regions: locate_line_regions() 的返回值
        target_page: "friend" | "chat" | "add_friend"
        monitor: 螢幕編號

    返回：更新後的 regions（含新頁面的完整內容）
    """
    page_to_btn = {
        "friend": "btn_friend",
        "chat": "btn_chat",
        "add_friend": "btn_add_friend",
    }

    if target_page not in page_to_btn:
        raise ValueError(f"未知頁面: {target_page}，支援: friend, chat, add_friend")

    if regions.get("current_page") == target_page:
        _print(f"[line_locate] 已經在 {target_page} 頁面")
        return regions

    btn_name = page_to_btn[target_page]
    btn = regions["sidebar_icons"].get(btn_name)
    if not btn or btn["center"] == (0, 0):
        raise RuntimeError(f"找不到 {btn_name} 按鈕座標")

    cx, cy = btn["center"]
    _print(f"[line_locate] 切換到 {target_page}，點擊 {btn_name} at ({cx}, {cy})")
    pyautogui.click(cx, cy)
    time.sleep(1.0)

    # 重新定位 + 掃描新頁面
    new_regions = locate_line_regions(monitor)

    if new_regions["current_page"] != target_page:
        _print(f"[line_locate] 警告：切換後偵測到 {new_regions['current_page']}，預期 {target_page}")

    return new_regions


def find_conversation(regions, name):
    """
    在聊天列表中找到指定名稱的對話，回傳其座標

    參數：
        regions: locate_line_regions() 的返回值（必須在 chat 頁面）
        name: 要找的對話名稱（模糊匹配）

    返回：
        {"name": "...", "center": (x, y)} 或 None
    """
    if regions.get("current_page") != "chat":
        _print(f"[line_locate] 不在聊天頁，無法搜尋對話")
        return None

    conversations = regions.get("page_content", {}).get("conversations", [])
    for conv in conversations:
        if name in conv["name"] or conv["name"] in name:
            return conv

    return None


def find_friend(regions, name):
    """
    在好友列表中找到指定名稱的好友，回傳其座標

    參數：
        regions: locate_line_regions() 的返回值（必須在 friend 頁面）
        name: 要找的好友名稱（模糊匹配）

    返回：
        {"name": "...", "center": (x, y)} 或 None
    """
    if regions.get("current_page") != "friend":
        _print(f"[line_locate] 不在好友頁，無法搜尋好友")
        return None

    pc = regions.get("page_content", {})
    for section in ["friends", "communities", "groups"]:
        for item in pc.get(section, []):
            if name in item["name"] or item["name"] in name:
                return item

    return None


# === 直接執行時測試 ===
if __name__ == "__main__":
    regions = locate_line_regions(monitor=2)
    print("\n=== 完整結果 ===")
    # 只印重點，不印全部 OCR 原始資料
    summary = {k: v for k, v in regions.items() if k != "page_content"}
    _print(json.dumps(summary, ensure_ascii=False, indent=2, default=str))

    pc = regions.get("page_content", {})
    print(f"\n=== {pc.get('page', '?')} 頁面內容 ===")
    if pc.get("page") == "chat":
        for c in pc.get("conversations", []):
            _print(f"  {c['name']} (unread={c.get('unread',0)}) at {c['center']}")
    elif pc.get("page") == "friend":
        _print(f"  用戶: {pc.get('user_name', '?')}")
        for s in ["communities", "groups", "friends"]:
            _print(f"  {s}: {[i['name'] for i in pc.get(s, [])]}")
    elif pc.get("page") == "add_friend":
        for k, v in pc.get("buttons", {}).items():
            _print(f"  {v['label']} at {v['center']}")
