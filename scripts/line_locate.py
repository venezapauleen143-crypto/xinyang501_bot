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

# chat_title 固定點擊位置（從 line1.png 黃框實測，相對 LINE 視窗）
# 不管名字多長，點擊位置固定不變，不依賴 Vision API
CHAT_TITLE_CLICK = {
    "x_ratio": 0.5417,               # line1 黃框 center x（相對 LINE 視窗寬度）
    "y_ratio": 0.0567,               # line1 黃框 center y（相對 LINE 視窗高度）
}

# 好友頁（line1.png + line8.png）— left_panel 內的固定元素相對位置
FRIEND_PAGE_TEMPLATE = {
    # 搜尋欄（相對 left_panel）
    "search_bar": {"y_ratio": 0.04, "h_ratio": 0.03},
    # 用戶個人資料區（搜尋欄下方）
    "user_profile": {"y_ratio": 0.07, "h_ratio": 0.07},
    # 我的最愛區塊（line8.png，在 user_profile 和 community 之間）
    "section_favorite": {"y_ratio": 0.12, "label": "我的最愛"},
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
    # 未讀徽章 x 位置（line16.png，綠色圓形徽章在項目右側）
    "unread_badge_x_ratio": 0.92,
}

# 未讀訊息偵測（line16.png）
UNREAD_BADGE = {
    # sidebar 聊天按鈕的紅色通知徽章（總未讀數）
    # 紅色徽章 RGB ≈ (230+, <80, <80)，在 sidebar 聊天按鈕右上角
    "sidebar_badge_color": {"r_min": 200, "g_max": 80, "b_max": 80},
    # 聊天列表每個對話的綠色未讀徽章
    # 綠色徽章 RGB ≈ LINE 品牌綠，在項目右側
    "chat_badge_color": {"r_max": 100, "g_min": 150, "b_max": 100},
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


# 搜尋結果頁（line7.png）— 搜尋好友後的 left_panel 佈局
SEARCH_RESULT_TEMPLATE = {
    # 搜尋欄底部（此 y 以上的 OCR 結果是搜尋欄裡打的字，必須跳過）
    "search_bar_bottom_ratio": 0.068,
    # 分類標題（「好友 1」）
    "section_header_ratio": 0.092,
    # 第一個搜尋結果項目起始 y
    "first_item_top_ratio": 0.102,
    # 第一個搜尋結果項目中心 y（點擊這裡）
    "first_item_center_ratio": 0.129,
    # 每個搜尋結果項目的高度
    "item_height_ratio": 0.053,
}


# ============================================================
# PaddleOCR GPU 全域引擎（只載入一次）
# GPU 記憶體保護：限制 VRAM 用量，防止多進程同時跑導致 OOM 當機
# ============================================================
_ocr_engine = None


def _get_ocr_engine():
    """取得全域 PaddleOCR 引擎（singleton，限制 GPU 記憶體）"""
    global _ocr_engine
    if _ocr_engine is None:
        if "modelscope" not in sys.modules:
            fake = types.ModuleType("modelscope")
            fake.__version__ = "0.0.0"
            sys.modules["modelscope"] = fake
            sys.modules["modelscope.utils"] = types.ModuleType("modelscope.utils")
            sys.modules["modelscope.utils.import_utils"] = types.ModuleType("modelscope.utils.import_utils")

        # GPU 記憶體限制（RTX 3060 6GB，限制在 ~900MB 以內）
        os.environ["FLAGS_fraction_of_gpu_memory_to_use"] = "0.15"
        os.environ["FLAGS_initial_gpu_memory_in_mb"] = "512"
        os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"

        from paddleocr import PaddleOCR
        _ocr_engine = PaddleOCR(
            lang="chinese_cht",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
            # gpu_mem 由環境變數 FLAGS_fraction_of_gpu_memory_to_use 控制
        )
        _print("[line_locate] PaddleOCR GPU 已載入（VRAM 限制 15% ≈ 900MB）")
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
    # 確保 DPI 設定正確（每次呼叫都設，防止被其他模組覆蓋）
    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(0)
    except Exception:
        pass

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

    x_center = sidebar_width // 2

    icons = {}
    for btn_name, btn_info in SIDEBAR_BUTTONS.items():
        local_y = int(line_height * btn_info["y_ratio"])
        abs_x = int(mon["left"] + (il + x_center) * sx_ratio)
        abs_y = int(mon["top"] + (it + local_y) * sy_ratio)
        icons[btn_name] = {"center": (abs_x, abs_y), "label": btn_info["label"], "local_y": local_y}

    return icons


def detect_current_page_by_pixel(line_crop, sidebar_width):
    """
    用像素亮度偵測當前頁面（0 秒，不用 API）。
    選中的 sidebar 按鈕會高亮（亮度 > 150），未選中的暗（亮度 < 130）。
    """
    arr = np.array(line_crop)
    x_center = sidebar_width // 2
    lh = arr.shape[0]

    page_buttons = {
        "friend": int(lh * SIDEBAR_BUTTONS["btn_friend"]["y_ratio"]),
        "chat": int(lh * SIDEBAR_BUTTONS["btn_chat"]["y_ratio"]),
        "add_friend": int(lh * SIDEBAR_BUTTONS["btn_add_friend"]["y_ratio"]),
    }

    # 取每個按鈕 10x10 區域的最大亮度
    brightness = {}
    for page, y in page_buttons.items():
        y1 = max(0, y - 5)
        y2 = min(arr.shape[0], y + 6)
        x1 = max(0, x_center - 10)
        x2 = min(arr.shape[1], x_center + 10)
        region = arr[y1:y2, x1:x2, :]
        brightness[page] = float(np.max(np.mean(region, axis=2)))

    # 最亮的就是當前頁面
    max_val = max(brightness.values())
    min_val = min(brightness.values())

    # 如果三個亮度差距太小，無法判斷
    if max_val - min_val < 20:
        _print(f"[line_locate] 像素偵測無法判斷（差距太小）: {brightness}")
        return "unknown"

    current = max(brightness, key=brightness.get)
    _print(f"[line_locate] 像素偵測頁面: {brightness} → {current}")
    return current


def detect_current_page_by_ocr(ocr_items):
    """
    用 OCR 關鍵字偵測當前頁面（備援，從已掃描的 OCR 結果判斷）。
    """
    all_text = " ".join(item["text"] for item in ocr_items)

    # 聊天頁特徵：有「全部」「群組」「社群」分頁標籤
    chat_keywords = ["全部", "搜尋聊天和訊息"]
    # 好友頁特徵：有分類標題
    friend_keywords = ["社群 1", "群組 1", "好友 ", "以姓名搜尋", "Keep筆記"]
    # 加好友頁特徵
    add_friend_keywords = ["搜尋好友", "建立群組", "建立社群", "官方推薦"]

    chat_score = sum(1 for kw in chat_keywords if kw in all_text)
    friend_score = sum(1 for kw in friend_keywords if kw in all_text)
    add_friend_score = sum(1 for kw in add_friend_keywords if kw in all_text)

    scores = {"friend": friend_score, "chat": chat_score, "add_friend": add_friend_score}
    best = max(scores, key=scores.get)

    if scores[best] == 0:
        return "unknown"

    _print(f"[line_locate] OCR 偵測頁面: {scores} → {best}")
    return best


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
        "3. chat_title: Only the friend/group NAME text at top of chat area (right side). Do NOT include member count, icons or buttons. Just the name text itself.\n"
        "4. chat_area: Message area with conversation bubbles.\n"
        "5. input_box: Text input field at bottom right. Only the text area.\n\n"
        "Return RAW JSON only:\n"
        '{"search_bar":{"l":0,"t":0,"r":0,"b":0},'
        '"left_panel":{"l":0,"t":0,"r":0,"b":0},'
        '"chat_title":{"l":0,"t":0,"r":0,"b":0},'
        '"chat_area":{"l":0,"t":0,"r":0,"b":0},'
        '"input_box":{"l":0,"t":0,"r":0,"b":0}}'
    )

    for attempt in range(3):
        try:
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
            data = json.loads(resp)

            # 確保所有座標值都是 int（Vision 有時回傳字串）
            for key in data:
                if isinstance(data[key], dict):
                    for k, v in data[key].items():
                        if isinstance(v, str) and v.isdigit():
                            data[key][k] = int(v)
            return data
        except (json.JSONDecodeError, Exception) as e:
            _print(f"[line_locate] Vision API 回傳異常（第{attempt+1}次）: {e}")
            if attempt < 2:
                time.sleep(2)

    raise RuntimeError("Vision API 連續 3 次回傳不合法 JSON")


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
        "favorite": int(ph * template["section_favorite"]["y_ratio"]),
        "community": int(ph * template["section_community"]["y_ratio"]),
        "group": int(ph * template["section_group"]["y_ratio"]),
        "friend": int(ph * template["section_friend"]["y_ratio"]),
    }

    # 分類 OCR 結果到各區塊
    favorites = []
    communities = []
    groups = []
    friends = []
    user_name = ""

    skip_keywords = ["搜尋", "Keep", "姓名", "我的最愛", "社群", "群組", "好友"]

    for item in ocr_items:
        y = item["y"]
        text = item["text"]

        # 用戶名稱（最上方）
        if y < section_y["favorite"] and not user_name and len(text) > 1:
            if not any(kw in text for kw in skip_keywords):
                user_name = text

        # 我的最愛區塊（line8.png）
        elif section_y["favorite"] <= y < section_y["community"]:
            if len(text) > 1 and not any(kw in text for kw in skip_keywords):
                favorites.append({"name": text, "center": to_abs(pw // 2, y + 15)})

        # 社群區塊的項目
        elif section_y["community"] <= y < section_y["group"]:
            if len(text) > 1 and not any(kw in text for kw in skip_keywords):
                communities.append({"name": text, "center": to_abs(pw // 2, y + 15)})

        # 群組區塊的項目
        elif section_y["group"] <= y < section_y["friend"]:
            if len(text) > 1 and not any(kw in text for kw in skip_keywords):
                groups.append({"name": text, "center": to_abs(pw // 2, y + 15)})

        # 好友區塊的項目
        elif y >= section_y["friend"]:
            if len(text) > 1 and not any(kw in text for kw in skip_keywords):
                friends.append({"name": text, "center": to_abs(pw // 2, y + 15)})

    return {
        "page": "friend",
        "user_name": user_name,
        "favorites": favorites,
        "communities": communities,
        "groups": groups,
        "friends": friends,
        "fixed_elements": {
            "search_bar": {"center": to_abs(pw // 2, int(ph * template["search_bar"]["y_ratio"]))},
            "user_profile": {"center": to_abs(pw // 2, int(ph * template["user_profile"]["y_ratio"]))},
            "section_favorite_header": {"center": to_abs(pw // 4, section_y["favorite"])},
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

    # Step 3: 像素偵測當前頁面（0 秒）
    current_page = detect_current_page_by_pixel(line_crop, sidebar_w)

    # Step 4: 所有框架區域改用固定座標（不再呼叫 Vision API）
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
    # sidebar、search_bar、left_panel 用 Vision
    # chat_area、input_box 用像素比例（不靠 Vision，避免定位錯誤）
    # sidebar：固定位置（從 line1_new.png 白框比對，741x1031 圖片）
    sidebar_raw = {"l": 9, "t": 26, "r": 55, "b": 417}
    sidebar = to_screen_region(sidebar_raw)
    # search_bar：根據頁面使用不同固定位置
    if current_page == "friend":
        # 好友頁（從 line1.png 比對，縮放到 741x1031）
        search_bar_raw = {"l": 68, "t": 37, "r": 322, "b": 64}
    else:
        # 聊天頁（從 line1_new.png 綠框比對，741x1031 圖片）
        search_bar_raw = {"l": 78, "t": 66, "r": 310, "b": 96}
    search_bar = to_screen_region(search_bar_raw)
    # left_panel：根據頁面使用不同固定位置
    if current_page == "friend":
        # 好友頁（從 line1.png 灰框比對，縮放到 741x1031）
        left_panel_raw = {"l": 55, "t": 70, "r": 351, "b": 1021}
    else:
        # 聊天頁（從 line1_new.png 藍框比對，741x1031 圖片）
        left_panel_raw = {"l": 64, "t": 111, "r": 357, "b": 948}
    left_panel = to_screen_region(left_panel_raw)
    # chat_title：固定位置（從 line1_new.png 黃框實測，741x1031 圖片）
    # 黃框 x=[366-412] y=[63-94] center=(389,78)
    # 永遠點這個固定位置，不管客戶名字是什麼
    chat_title_raw = {"l": 366, "t": 63, "r": 412, "b": 94}
    chat_title = to_screen_region(chat_title_raw)

    # chat_area：固定位置（從 line1_new.png 黑框比對，741x1031 圖片）
    chat_area_raw = {"l": 371, "t": 99, "r": 737, "b": 910}
    chat_area = to_screen_region(chat_area_raw)
    # input_box：固定位置（從 line1_new.png 深藍框比對，741x1031 圖片）
    input_box_raw = {"l": 371, "t": 923, "r": 726, "b": 979}
    input_box = to_screen_region(input_box_raw)

    # sidebar 按鈕（用像素分析 + 模板比例，不靠 Vision）
    sidebar_icons = calc_sidebar_icons(sidebar_w, lh, il, it, mon, full_img.size)

    # Step 5: 掃描當前頁面內容（PaddleOCR 主力）
    _print(f"[line_locate] 掃描 {current_page} 頁面內容...")
    panel_rect = left_panel_raw
    panel_crop = line_crop.crop((panel_rect["l"], panel_rect["t"], panel_rect["r"], panel_rect["b"]))

    page_content = {}
    try:
        if current_page == "friend":
            page_content = scan_friend_page(panel_crop, panel_rect, mon, full_img.size, il, it)
        elif current_page == "chat":
            page_content = scan_chat_page(panel_crop, panel_rect, mon, full_img.size, il, it)
        elif current_page == "add_friend":
            page_content = scan_add_friend_page(panel_crop, panel_rect, mon, full_img.size, il, it)

        # OCR 交叉驗證（只在像素偵測不確定時才介入）
        ocr_items = page_content.get("all_ocr_items", [])
        if current_page == "unknown":
            ocr_page = detect_current_page_by_ocr(ocr_items)
            if ocr_page != "unknown":
                _print(f"[line_locate] 像素無法判斷，OCR 判斷為 {ocr_page}")
                current_page = ocr_page
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


def screenshot_line_window(monitor=2):
    """
    截取 LINE 視窗的完整截圖（只裁 LINE 視窗，跟 line16.png 一樣的方式）。
    用 win32gui 找視窗位置，裁切出 LINE 視窗部分。

    回傳：
        PIL Image（LINE 視窗截圖，約 742x1032）
    """
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
    return line_crop


def find_unread_badges(monitor=2):
    """
    用 LINE16 的固定位置逐行檢查像素顏色，找出有綠色未讀徽章的對話（line16.png 方式）。
    screenshot_line_window 截圖 → 在徽章固定 x 位置逐行檢查綠色 → 回傳有未讀的座標。

    LINE16 實測數據：
    - 徽章 x 位置：視窗寬度 * 0.455
    - 第一個徽章 y：視窗高度 * 0.151
    - 每個項目間隔：視窗高度 * 0.069

    參數：
        monitor: 螢幕編號

    回傳：
        [{"y": int, "center": (screen_x, screen_y)}]
    """
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
    arr = np.array(line_crop)
    lh, lw, _ = arr.shape
    sx_r = mon["width"] / full_img.size[0]
    sy_r = mon["height"] / full_img.size[1]

    # LINE16 實測的徽章位置
    badge_x = int(lw * 0.455)           # 徽章 x 位置
    first_badge_y = int(lh * 0.151)     # 第一個徽章 y
    item_interval = int(lh * 0.069)     # 每個項目間隔
    color = UNREAD_BADGE["chat_badge_color"]

    result = []
    y = first_badge_y

    while y < lh - 50:
        # 檢查這個位置的像素是不是綠色
        # 掃描徽章 x 附近幾個像素，增加可靠性
        green_count = 0
        for dx in range(-8, 9):
            for dy in range(-5, 6):
                sx = min(lw - 1, max(0, badge_x + dx))
                sy = min(lh - 1, max(0, y + dy))
                r, g, b = int(arr[sy, sx, 0]), int(arr[sy, sx, 1]), int(arr[sy, sx, 2])
                if r < color["r_max"] and g > color["g_min"] and b < color["b_max"]:
                    green_count += 1

        if green_count > 15:
            # 有綠色徽章，記錄點擊座標（點在項目中間偏左，名字的位置）
            click_x = int(mon["left"] + (il + int(lw * 0.25)) * sx_r)
            click_y = int(mon["top"] + (it + y) * sy_r)
            result.append({"y": y, "center": (click_x, click_y)})

        y += item_interval

    _print(f"[line_locate] 綠色未讀徽章（LINE16 方式）: {len(result)} 個")
    return result


def detect_unread(monitor=2):
    """
    截取 LINE 視窗截圖（line16.png 方式），存檔供讀取判斷。
    用 screenshot_line_window 裁切完整 LINE 視窗（742x1032），
    截圖清晰，可直接看到所有未讀徽章數字。

    回傳：
        截圖檔案路徑
    """
    line_img = screenshot_line_window(monitor)
    save_path = os.path.join(TMPDIR, "line_unread_check.png")
    line_img.save(save_path, quality=95)
    _print(f"[line_locate] LINE 視窗截圖已存: {save_path} ({line_img.size[0]}x{line_img.size[1]})")
    return save_path


def find_friend(regions, name):
    """
    在好友列表中找到指定名稱的好友，回傳其座標（支援繁簡體模糊匹配）

    參數：
        regions: locate_line_regions() 的返回值（必須在 friend 頁面）
        name: 要找的好友名稱（模糊匹配）

    返回：
        {"name": "...", "center": (x, y)} 或 None
    """
    from difflib import SequenceMatcher

    if regions.get("current_page") != "friend":
        _print(f"[line_locate] 不在好友頁，無法搜尋好友")
        return None

    pc = regions.get("page_content", {})
    for section in ["favorites", "friends", "communities", "groups"]:
        for item in pc.get(section, []):
            # 完全匹配或包含
            if name in item["name"] or item["name"] in name:
                return item
            # 模糊匹配（處理繁簡體差異、OCR 微小辨識差異）
            ratio = SequenceMatcher(None, name, item["name"]).ratio()
            if ratio > 0.5:
                return item

    return None


# ============================================================
# 分享好友資訊（line8.png + line9.png + line10.png 流程）
# ============================================================
# 「選擇傳送對象」面板位置（line9.png 黃色框）
# 面板出現在 LINE 視窗的【左邊外側】，不在 LINE 視窗裡面
# 從 click_positions_v2 于晏哥標記的正確位置反算：
# 面板右邊界 = LINE 左邊界 - 129px（不是緊貼 LINE）
# 面板 top = LINE top + 298px（不是從 LINE 頂部開始）
SHARE_PANEL_POSITION = {
    "width": 353,                    # 面板寬度（固定像素）
    "height": 555,                   # 面板高度（固定像素）
    "right_offset": 129,             # 面板右邊界距 LINE 左邊界的距離
    "top_offset": 298,               # 面板 top 距 LINE top 的距離
}

# 面板內部元素偏移量（相對面板左上角的固定像素偏移，從參考圖實測，不用比例）
SHARE_DIALOG_OFFSETS = {
    "search_bar_x": 176,             # 搜尋欄 x（line10 黑色框）
    "search_bar_y": 108,             # 搜尋欄 y（line10 黑色框）
    "circle_x": 326,                 # 圓圈 x（line14 紅點）
    "circle_y": 188,                 # 圓圈 y（line14 紅點）
    "share_btn_x": 120,              # 分享按鈕 x（line12 紅色框）
    "share_btn_y": 563,              # 分享按鈕 y（line12 紅色框）
}


# 個人資料小卡片 LINE17 偏移量（從 line19 放大圖實測，相對面板左上角）
# line19 面板尺寸 287x508
PROFILE_CARD_OFFSETS = {
    "close_x": 259,                  # X關閉按鈕 x（line19 綠框）
    "close_y": 22,                   # X關閉按鈕 y（line19 綠框）
    "name_x": 146,                   # 名字文字 x（line19 紅框，水平置中）
    "name_y": 322,                   # 名字文字 y（line19 紅框，頭像下方）
}

# 個人資料編輯畫面 LINE18 偏移量（從 line20 放大圖實測，相對面板左上角）
# line20 面板尺寸 283x509
PROFILE_EDIT_OFFSETS = {
    "name_x": 140,                   # 名字文字 x（line20 紅框，水平置中）
    "name_y": 325,                   # 名字文字 y（line20 紅框，頭像下方）
    "save_x": 105,                   # 儲存按鈕 x（line20 黃框，綠色按鈕）
    "save_y": 474,                   # 儲存按鈕 y（line20 黃框）
    "cancel_x": 177,                 # 取消按鈕 x（白色按鈕，儲存右邊）
    "cancel_y": 473,                 # 取消按鈕 y
}


def rename_friend(regions, new_name, monitor=2):
    """
    把當前聊天對象的好友名稱改成 new_name（LINE17 → LINE18 流程）。

    LINE17/LINE18 是同一個獨立視窗（288x512, class=Qt663QWindowIcon），內容切換。
    偏移量從 line19（小卡片放大圖）和 line20（編輯畫面放大圖）實測。

    流程：
    1. 點擊 chat_title（LINE17 黃框）→ 個人資料小卡片彈出
    2. win32gui 找到小卡片視窗（Qt663QWindowIcon）
    3. 點擊名字（line19 紅框偏移）→ 視窗切換為編輯畫面
    4. win32gui 重新抓視窗座標
    5. 三次點擊名字（line20 紅框偏移）→ 全選
    6. 輸入 new_name → 覆蓋原名
    7. 點擊儲存（line20 黃框偏移）→ 視窗切回小卡片
    8. 點 X 關閉（line19 綠框偏移）

    參數：
        regions: locate_line_regions() 的返回值（需要 chat_title）
        new_name: 要改成的新名稱（通常是編號 = 電話後五碼）
        monitor: 螢幕編號
    """
    import pyautogui
    import pyperclip

    _print(f"[rename] 開始改名: {new_name}")

    # Step 1: 點擊 chat_title 開啟個人資料小卡片（用定位腳本的 regions）
    ct = regions.get("chat_title", {})
    ct_cx, ct_cy = ct.get("center", (0, 0))
    if ct_cx == 0:
        _print("[rename] 錯誤: 找不到 chat_title")
        return False

    _print(f"[rename] Step1: 點擊 chat_title ({ct_cx}, {ct_cy})")
    pyautogui.click(ct_cx, ct_cy)
    time.sleep(1.5)

    # Step 2: 用 win32gui 找到個人資料小卡片（LINE17 獨立子視窗）
    profile_card = _find_line_popup("profile_card")
    if not profile_card:
        _print("[rename] 錯誤: 找不到個人資料小卡片視窗")
        pyautogui.press("escape")
        return False

    card_hwnd, card_left, card_top, card_right, card_bottom = profile_card
    _print(f"[rename] Step2: 找到小卡片 hwnd={card_hwnd} ({card_left},{card_top})-({card_right},{card_bottom})")

    # Step 3: 點擊小卡片的名字（LINE17 紅框）→ 開啟編輯畫面
    name_x = card_left + PROFILE_CARD_OFFSETS["name_x"]
    name_y = card_top + PROFILE_CARD_OFFSETS["name_y"]
    _print(f"[rename] Step3: 點擊名字 ({name_x}, {name_y})")
    pyautogui.click(name_x, name_y)
    time.sleep(1.5)

    # Step 4: LINE18 是同一個視窗（內容切換），重新抓座標
    edit_popup = _find_line_popup("profile_edit")
    if not edit_popup:
        _print("[rename] 錯誤: 找不到編輯畫面視窗")
        pyautogui.press("escape")
        return False

    edit_hwnd, edit_left, edit_top, edit_right, edit_bottom = edit_popup
    _print(f"[rename] Step4: 編輯畫面 hwnd={edit_hwnd} ({edit_left},{edit_top})-({edit_right},{edit_bottom})")

    # Step 5: 三次點擊名字（LINE18 紅框）→ 全選
    edit_name_x = edit_left + PROFILE_EDIT_OFFSETS["name_x"]
    edit_name_y = edit_top + PROFILE_EDIT_OFFSETS["name_y"]
    _print(f"[rename] Step5: 三次點擊名字 ({edit_name_x}, {edit_name_y})")
    pyautogui.click(edit_name_x, edit_name_y, clicks=3)
    time.sleep(0.5)

    # Step 6: 輸入新名稱（用剪貼簿貼上，避免中文輸入問題）
    _print(f"[rename] Step6: 輸入新名稱 '{new_name}'")
    pyperclip.copy(new_name)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.5)

    # Step 7: 點擊儲存（LINE18 黃框 = 綠色儲存按鈕）
    save_x = edit_left + PROFILE_EDIT_OFFSETS["save_x"]
    save_y = edit_top + PROFILE_EDIT_OFFSETS["save_y"]
    _print(f"[rename] Step7: 點擊儲存 ({save_x}, {save_y})")
    pyautogui.click(save_x, save_y)
    time.sleep(1.5)

    # Step 8: 儲存後回到小卡片（同一個視窗切回 LINE17）→ 點 X 關閉（LINE19 綠框）
    profile_card2 = _find_line_popup("profile_card")
    if profile_card2:
        _, card2_left, card2_top, _, _ = profile_card2
        close_x = card2_left + PROFILE_CARD_OFFSETS["close_x"]
        close_y = card2_top + PROFILE_CARD_OFFSETS["close_y"]
        _print(f"[rename] Step8: 點擊 X 關閉 ({close_x}, {close_y})")
        pyautogui.click(close_x, close_y)
    else:
        _print("[rename] 小卡片已關閉，按 Esc")
        pyautogui.press("escape")
    time.sleep(0.5)

    _print(f"[rename] 改名完成: {new_name}")
    return True


def _find_line_popup(popup_type="profile_card"):
    """
    用 win32gui 找 LINE 的獨立彈窗視窗。

    popup_type:
        "profile_card" — LINE17 個人資料小卡片（288x512，class=Qt663QWindowIcon）
        "profile_edit" — LINE18 個人資料編輯畫面（同一個視窗，內容切換）
        "share_dialog" — LINE9 選擇傳送對象面板

    回傳：(hwnd, left, top, right, bottom) 或 None
    """
    import win32gui

    # 找 LINE 主視窗的 hwnd，用來排除
    line = find_line_window()
    line_hwnd = line[0] if line else None

    results = []

    def _enum_callback(hwnd, _):
        if not win32gui.IsWindowVisible(hwnd):
            return
        if hwnd == line_hwnd:
            return
        try:
            cls = win32gui.GetClassName(hwnd)
        except Exception:
            return

        # LINE 彈窗 class：Qt663QWindowIcon 或 EVA_Window
        if "Qt6" not in cls and "EVA_Window" not in cls:
            return

        # 排除 Note/Glow/shadow（通知氣泡）
        if "Note" in cls or "Glow" in cls or "shadow" in cls:
            return

        rect = win32gui.GetWindowRect(hwnd)
        left, top, right, bottom = rect
        w = right - left
        h = bottom - top

        # 排除太大的（主視窗）和太小的
        if w > 600 or h > 800:
            return
        if w < 50 or h < 50:
            return

        results.append((hwnd, left, top, right, bottom, w, h))

    win32gui.EnumWindows(_enum_callback, None)

    if not results:
        return None

    # 根據 popup_type 篩選
    for hwnd, left, top, right, bottom, w, h in results:
        if popup_type in ("profile_card", "profile_edit"):
            # LINE17/LINE18：同一個視窗 288x512（實測），class=Qt663QWindowIcon
            if 200 < w < 400 and 400 < h < 600:
                return (hwnd, left, top, right, bottom)
        elif popup_type == "share_dialog":
            # LINE9 分享面板：寬約 300-400，高約 500-650
            if 250 < w < 450 and 450 < h < 700:
                return (hwnd, left, top, right, bottom)

    # fallback: 回傳最小的那個
    results.sort(key=lambda x: x[5] * x[6])
    if results:
        r = results[0]
        return (r[0], r[1], r[2], r[3], r[4])

    return None


def share_contact_card(regions, share_who, share_to, monitor=2):
    """
    分享好友資訊卡給指定的人（line8 → line9 → line10 流程）。

    流程：
    1. 切到好友頁 → 搜尋 share_who → 找到並右鍵點擊（line8）
    2. 右鍵選單 → OCR 找「分享好友資訊」→ 點擊
    3. 「選擇傳送對象」面板出現在 LINE 視窗左邊外側（line9 黃色框）
    4. 在外側面板搜尋 share_to（line9 黑色框）
    5. 點圓圈（line9 紅色框）
    6. 點「分享」（line10 黃色框）

    參數：
        regions: locate_line_regions() 的返回值
        share_who: 要分享的好友名稱（如「溫妮」）
        share_to: 分享給誰（如「仁輝 JAMES」）
        monitor: 螢幕編號

    返回：
        True（成功）或 False（失敗）
    """
    import pyautogui
    import pyperclip

    # === Step 1: 切到好友頁 → 搜尋 share_who → 找到 ===
    _print(f"[share] 搜尋好友: {share_who}")
    if regions.get("current_page") != "friend":
        regions = switch_page(regions, "friend", monitor)

    # 用搜尋功能找 share_who
    friend_pos = search_friend_and_scan(regions, share_who, monitor)
    if friend_pos is None:
        _print(f"[share] 搜尋找不到 {share_who}")
        return False

    # === Step 2: 右鍵點擊 → OCR 找「分享好友資訊」→ 點擊（line8）===
    fx, fy = friend_pos["center"]
    _print(f"[share] 右鍵點擊 {share_who} at ({fx}, {fy})")
    pyautogui.rightClick(fx, fy)
    time.sleep(0.8)

    # 截圖 OCR 找「分享好友資訊」
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
    sx_r = mon["width"] / full_img.size[0]
    sy_r = mon["height"] / full_img.size[1]

    menu_items = ocr_scan_panel(line_crop)
    menu_target = None
    for item in menu_items:
        if "分享" in item["text"] and "好友" in item["text"]:
            menu_target = item
            break
        if "分享好友" in item["text"]:
            menu_target = item
            break

    if menu_target is None:
        _print(f"[share] OCR 找不到「分享好友資訊」選項")
        pyautogui.press("escape")
        time.sleep(0.3)
        return False

    menu_x = int(mon["left"] + (il + menu_target["x"] + 50) * sx_r)
    menu_y = int(mon["top"] + (it + menu_target["y"] + 10) * sy_r)
    _print(f"[share] 點擊「分享好友資訊」at ({menu_x}, {menu_y})")
    pyautogui.click(menu_x, menu_y)
    time.sleep(1.0)

    # === Step 3-6: 「選擇傳送對象」面板是獨立的 LINE 子視窗（line9 黃色框）===
    # 用 win32gui 直接找到面板視窗的螢幕座標，不用算
    time.sleep(0.5)
    panel_hwnd = None
    panel_rect = None
    all_line = []
    def _find_share_panel(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if "LINE" in title:
                rect = win32gui.GetWindowRect(hwnd)
                w = rect[2] - rect[0]
                h = rect[3] - rect[1]
                all_line.append((hwnd, rect, w, h))
    win32gui.EnumWindows(_find_share_panel, None)

    # 面板是較小的那個 LINE 視窗（約 350x550）
    for hwnd, rect, w, h in all_line:
        if 300 < w < 400 and 500 < h < 600:
            panel_hwnd = hwnd
            panel_rect = rect
            break

    if panel_rect is None:
        _print(f"[share] 找不到「選擇傳送對象」面板視窗")
        return False

    panel_left = panel_rect[0]
    panel_top = panel_rect[1]
    panel_w = panel_rect[2] - panel_rect[0]
    panel_h = panel_rect[3] - panel_rect[1]
    panel_cx = (panel_left + panel_rect[2]) // 2

    _print(f"[share] 外側面板: ({panel_left},{panel_top}) 寬={panel_w} 高={panel_h}")

    # Step 4: 點搜尋欄（line10 黑色框）
    search_x = panel_left + SHARE_DIALOG_OFFSETS["search_bar_x"]
    search_y = panel_top + SHARE_DIALOG_OFFSETS["search_bar_y"]
    _print(f"[share] 點搜尋欄 at ({search_x}, {search_y})")
    pyautogui.click(search_x, search_y)
    time.sleep(0.3)
    pyperclip.copy(share_to)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1.5)

    # Step 5: 點圓圈（line14 紅點）
    circle_x = panel_left + SHARE_DIALOG_OFFSETS["circle_x"]
    circle_y = panel_top + SHARE_DIALOG_OFFSETS["circle_y"]
    _print(f"[share] 點擊圓圈 at ({circle_x}, {circle_y})")
    pyautogui.click(circle_x, circle_y)
    time.sleep(0.5)

    # Step 6: 點「分享」按鈕（line12 紅色框）
    share_btn_x = panel_left + SHARE_DIALOG_OFFSETS["share_btn_x"]
    share_btn_y = panel_top + SHARE_DIALOG_OFFSETS["share_btn_y"]
    _print(f"[share] 點擊「分享」at ({share_btn_x}, {share_btn_y})")
    pyautogui.click(share_btn_x, share_btn_y)
    time.sleep(1.0)

    _print(f"[share] 已分享 {share_who} 的好友資訊給 {share_to}")
    return True


# ============================================================
# 對話區截圖 + OCR（純像素定位，不用 Vision）
# ============================================================
def screenshot_chat_area(regions, monitor=2):
    """
    截取 LINE 對話區圖片。
    用 regions 裡的 chat_area 座標（像素比例算出，不是 Vision）。

    參數：
        regions: locate_line_regions() 的返回值
        monitor: 螢幕編號

    返回：
        PIL Image（對話區截圖）
    """
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
    lw, lh = line_crop.size

    ca = regions["chat_area"]
    sx_ratio = full_img.size[0] / mon["width"]
    sy_ratio = full_img.size[1] / mon["height"]

    ca_il = max(0, int((ca["left"] - mon["left"]) * sx_ratio) - il)
    ca_it = max(0, int((ca["top"] - mon["top"]) * sy_ratio) - it)
    ca_ir = min(lw, int((ca["right"] - mon["left"]) * sx_ratio) - il)
    ca_ib = min(lh, int((ca["bottom"] - mon["top"]) * sy_ratio) - it)

    # 安全檢查
    if ca_ir <= ca_il or ca_ib <= ca_it:
        sep = regions.get("separator_x", lw // 3)
        ca_il = sep
        ca_it = int(lh * 0.04)
        ca_ir = lw
        ca_ib = int(lh * 0.92)

    return line_crop.crop((ca_il, ca_it, ca_ir, ca_ib))


def _fix_ocr_spacing(text):
    """修復 OCR 吃掉空格的問題：中文和數字之間自動加空格，超長數字自動拆分"""
    # 中文後面接數字 → 加空格
    text = re.sub(r'([\u4e00-\u9fff])(\d)', r'\1 \2', text)
    # 數字後面接中文 → 加空格
    text = re.sub(r'(\d)([\u4e00-\u9fff])', r'\1 \2', text)

    # 連續超過 10 碼數字 → 嘗試拆分（前4碼生日 + 後10碼電話）
    def _split_long_digits(m):
        digits = m.group()
        if len(digits) >= 14:
            # 前4碼可能是生日（MMDD），後面是電話
            mm = digits[:2]
            if mm in ['01','02','03','04','05','06','07','08','09','10','11','12']:
                return digits[:4] + ' ' + digits[4:]
        return digits

    text = re.sub(r'\d{14,}', _split_long_digits, text)
    return text


def ocr_scan_chat(chat_img):
    """
    PaddleOCR 掃描對話區，提取所有訊息文字 + 用氣泡顏色判斷 sender。
    LINE 綠色氣泡 = 我方（me），白色氣泡 = 對方（them）。

    參數：
        chat_img: screenshot_chat_area() 的返回值

    返回：
        [{"text": "...", "sender": "them/me/unknown", "y": int, "conf": float}]
    """
    arr = np.array(chat_img)
    ch, cw = arr.shape[:2]

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

            # OCR 後處理：修復空格問題
            t = _fix_ocr_spacing(t)

            box = boxes[i] if i < len(boxes) else [[0, 0], [0, 0], [0, 0], [0, 0]]
            x = int(box[0][0])
            x_right = int(box[1][0])
            y = int(box[0][1])
            y_center = int((box[0][1] + box[2][1]) / 2)

            # 過濾時間戳和系統訊息
            if len(t) < 12 and (":" in t or "上午" in t or "下午" in t):
                continue
            if "已讀" in t or "已误" in t or "收回" in t:
                continue
            if "輸入訊息" in t:
                continue

            # 用氣泡顏色判斷 sender：掃描該行所有像素，數綠色像素數量
            # LINE 綠色氣泡 RGB ≈ (195,246,157)，從實際截圖取得
            scan_y = min(ch - 1, y_center)
            row = arr[scan_y, :, :]
            green_count = 0
            for px in range(cw):
                rv, gv, bv = int(row[px, 0]), int(row[px, 1]), int(row[px, 2])
                if 150 < rv < 210 and gv > 220 and 130 < bv < 180:
                    green_count += 1

            if green_count > 30:
                sender = "me"      # 該行有綠色氣泡 = 我方
            else:
                sender = "them"    # 該行沒有綠色 = 對方

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


# ============================================================
# 搜尋好友 + 掃描搜尋結果（line7.png 流程）
# ============================================================
def search_friend_and_scan(regions, friend_name, monitor=2):
    """
    在搜尋欄輸入好友名稱，掃描搜尋結果，回傳好友項目的螢幕座標。

    流程（對應 line7.png）：
    1. 點搜尋欄 → 輸入好友名稱
    2. OCR 掃描 left_panel（跳過搜尋欄文字）
    3. 在搜尋結果中找到好友 → 回傳座標

    參數：
        regions: locate_line_regions() 的返回值
        friend_name: 要搜尋的好友名稱
        monitor: 螢幕編號

    返回：
        {"name": "...", "center": (x, y)} 或 None
    """
    import pyautogui
    import pyperclip
    from difflib import SequenceMatcher

    # Step 1: 點搜尋欄 → 輸入好友名稱
    sx, sy = regions["search_bar"]["center"]
    pyautogui.click(sx, sy)
    time.sleep(0.5)
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.press("delete")
    time.sleep(0.1)
    pyperclip.copy(friend_name)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1.5)
    _print(f"[line_locate] 搜尋: {friend_name}")

    # Step 2: 截圖 + OCR 掃描 left_panel
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
    lw_px, lh_px = line_crop.size

    lp = regions["left_panel"]
    sx_ratio = full_img.size[0] / mon["width"]
    sy_ratio = full_img.size[1] / mon["height"]
    lp_il = max(0, int((lp["left"] - mon["left"]) * sx_ratio) - il)
    lp_it = max(0, int((lp["top"] - mon["top"]) * sy_ratio) - it)
    lp_ir = min(lw_px, int((lp["right"] - mon["left"]) * sx_ratio) - il)
    lp_ib = min(lh_px, int((lp["bottom"] - mon["top"]) * sy_ratio) - it)

    panel_crop = line_crop.crop((lp_il, lp_it, lp_ir, lp_ib))
    pw, ph = panel_crop.size

    # Step 3: OCR 掃描，跳過搜尋欄區域（line7.png 的關鍵）
    # 搜尋欄底部 y = 視窗高度 * search_bar_bottom_ratio
    skip_y = int(lh_px * SEARCH_RESULT_TEMPLATE["search_bar_bottom_ratio"])
    # 轉換成 panel_crop 內的 y（減去 panel 的 top offset）
    skip_y_in_panel = max(skip_y - lp_it, 30)
    _print(f"[line_locate] 搜尋結果：跳過 y < {skip_y_in_panel} 的 OCR 結果（搜尋欄文字）")

    ocr_items = ocr_scan_panel(panel_crop)

    # Step 4: 從搜尋欄下方找好友名（line7.png 黑框項目）
    for item in ocr_items:
        if item["y"] < skip_y_in_panel:
            continue  # 跳過搜尋欄裡的文字
        if "好友" in item["text"] and len(item["text"]) < 6:
            continue  # 跳過「好友 1」分類標題
        ratio = SequenceMatcher(None, friend_name, item["text"]).ratio()
        if friend_name in item["text"] or item["text"] in friend_name or ratio > 0.6:
            # 計算螢幕絕對座標（點擊 panel 中間 x，好友項目 y + 偏移）
            abs_x = int(mon["left"] + (il + lp_il + pw // 2) * (mon["width"] / full_img.size[0]))
            abs_y = int(mon["top"] + (it + lp_it + item["y"] + 25) * (mon["height"] / full_img.size[1]))
            _print(f"[line_locate] 搜尋結果找到: {item['text']} (y={item['y']}), 螢幕座標=({abs_x}, {abs_y})")
            return {"name": item["text"], "center": (abs_x, abs_y)}

    _print(f"[line_locate] OCR 找不到 {friend_name}")
    return None


def enter_chat_from_search(friend_pos, regions, monitor=2):
    """
    點擊搜尋結果中的好友項目（line7.png 黑框），進入聊天視窗。
    自動判斷：有聊天記錄 → 直接進入；沒有 → 點綠色聊天按鈕。

    參數：
        friend_pos: search_friend_and_scan() 的返回值 {"name", "center"}
        regions: 當前的 regions
        monitor: 螢幕編號

    返回：
        更新後的 regions（已進入聊天），或 None（失敗）
    """
    import pyautogui

    click_x, click_y = friend_pos["center"]

    for attempt in range(3):
        _print(f"[line_locate] 第 {attempt+1} 次點擊好友: ({click_x}, {click_y})")
        pyautogui.click(click_x, click_y)
        time.sleep(0.8)
        pyautogui.click(click_x, click_y)
        time.sleep(1.0)

        # 截圖判斷右側：有沒有 input_box / 綠色聊天按鈕
        full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
        arr = np.array(line_crop)
        lh2, lw2, _ = arr.shape
        sep = regions.get("separator_x", lw2 // 3)

        # 偵測綠色聊天按鈕（沒聊過天的新好友）
        green_points = []
        for y in range(lh2 // 4, lh2 * 3 // 4):
            for x in range(sep, lw2, 3):
                rv, gv, bv = int(arr[y, x, 0]), int(arr[y, x, 1]), int(arr[y, x, 2])
                if gv > 150 and gv > rv + 50 and gv > bv + 30 and rv < 100:
                    green_points.append((x, y))

        if len(green_points) > 20:
            # 沒聊過天 → 點綠色聊天按鈕
            avg_x = sum(p[0] for p in green_points) // len(green_points)
            avg_y = sum(p[1] for p in green_points) // len(green_points)
            sx_r = mon["width"] / full_img.size[0]
            sy_r = mon["height"] / full_img.size[1]
            btn_x = int(mon["left"] + (il + avg_x) * sx_r)
            btn_y = int(mon["top"] + (it + avg_y) * sy_r)
            _print(f"[line_locate] 新好友，點綠色聊天按鈕 ({btn_x}, {btn_y})")
            pyautogui.click(btn_x, btn_y)
            time.sleep(1.5)

        # 重新定位，驗證 chat_area
        new_regions = locate_line_regions(monitor)
        ca = new_regions["chat_area"]
        ca_height = ca["bottom"] - ca["top"]

        if ca_height >= 50:
            _print(f"[line_locate] 已進入聊天（chat_area 高度={ca_height}）")
            return new_regions
        else:
            _print(f"[line_locate] 尚未進入聊天（chat_area 高度={ca_height}），重試...")
            if attempt == 1:
                # 第三次嘗試：雙擊
                pyautogui.doubleClick(click_x, click_y)
                time.sleep(1.0)

    _print(f"[line_locate] 3 次嘗試都無法進入聊天")
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
