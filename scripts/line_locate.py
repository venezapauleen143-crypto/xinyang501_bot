"""
LINE 電腦版 UI 定位模組 — 每次運行都用 Claude Vision 精確找到所有區域
不使用固定座標，每次都重新偵測

LINE 有三個左側面板模式（依功能列圖示切換）：
- 好友頁：分類列表（社群/群組/好友）
- 聊天頁：最近對話列表
- 加好友頁：搜尋好友/建立群組/推薦

定位的 7 個區域：
1. sidebar        — 最左側功能列（好友/聊天/加好友等圖示按鈕）
2. search_bar     — 搜尋輸入框
3. left_panel     — 左側面板（聊天列表/好友列表/加好友功能）
4. chat_title     — 聊天標題（群組名/好友名）
5. chat_area      — 對話內容區
6. input_box      — 輸入框
7. current_page   — 目前在哪個頁面（friend/chat/add_friend）

使用方式：
    from line_locate import locate_line_regions
    regions = locate_line_regions(monitor=2)
    # regions["search_bar"]["center"]  → 搜尋欄座標
    # regions["chat_area"]             → 對話區座標
    # regions["input_box"]["center"]   → 輸入框座標
"""

import sys
import io
import os
import re
import json
import base64
import numpy as np

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
from PIL import Image
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"


def find_line_window():
    """用 win32gui 找到 LINE 主視窗"""
    results = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            cls = win32gui.GetClassName(hwnd)
            # LINE 視窗的 class 通常是 Qt 相關或含 LINE
            if "LINE" in title and win32gui.IsWindowVisible(hwnd):
                rect = win32gui.GetWindowRect(hwnd)
                w = rect[2] - rect[0]
                h = rect[3] - rect[1]
                # 過濾太小的視窗（通知等）
                if w > 300 and h > 300:
                    results.append((hwnd, title, cls, rect))
    win32gui.EnumWindows(callback, results)

    if not results:
        return None

    # 取最大的 LINE 視窗
    main = max(results, key=lambda x: (x[3][2] - x[3][0]) * (x[3][3] - x[3][1]))
    return main


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

    # 座標轉換
    sx = iw / mon["width"]
    sy = ih / mon["height"]
    il = int((wl - mon["left"]) * sx)
    it = int((wt - mon["top"]) * sy)
    ir = int((wr - mon["left"]) * sx)
    ib = int((wb - mon["top"]) * sy)

    # 確保在圖片範圍內
    il = max(0, il)
    it = max(0, it)
    ir = min(iw, ir)
    ib = min(ih, ib)

    line_crop = pil.crop((il, it, ir, ib))

    return pil, line_crop, (il, it, ir, ib), mon


def find_sidebar_by_pixel(line_crop):
    """像素分析找最左側功能列的寬度（通常是深色窄條）"""
    arr = np.array(line_crop)
    h, w, _ = arr.shape

    # 功能列通常在最左邊 30-60px，背景偏深色或灰色
    # 找左側第一條明顯的垂直分隔線
    for x in range(30, min(80, w)):
        col = arr[h // 4: h * 3 // 4, x, :]
        avg_brightness = np.mean(col)
        # 如果突然變亮，代表功能列結束
        prev_col = arr[h // 4: h * 3 // 4, x - 1, :]
        prev_brightness = np.mean(prev_col)
        if avg_brightness - prev_brightness > 30:
            return x

    return 50  # 預設值


def find_separator_by_pixel(line_crop, sidebar_width):
    """像素分析找左側面板和對話區的分隔線"""
    arr = np.array(line_crop)
    h, w, _ = arr.shape

    # 從 sidebar 之後開始找垂直分隔線
    candidates = []
    for check_y in range(int(h * 0.3), int(h * 0.7), int(h * 0.05)):
        row = arr[check_y, :, :]
        for x in range(sidebar_width + 50, w * 2 // 3):
            r, g, b = int(row[x, 0]), int(row[x, 1]), int(row[x, 2])
            prev_r, prev_g, prev_b = int(row[x - 1, 0]), int(row[x - 1, 1]), int(row[x - 1, 2])
            # 分隔線通常是細的灰線
            brightness = (r + g + b) / 3
            prev_brightness = (prev_r + prev_g + prev_b) / 3
            if abs(brightness - prev_brightness) > 20 and 180 < brightness < 220:
                candidates.append(x)
                break

    if candidates:
        return sorted(candidates)[len(candidates) // 2]
    return w // 3


def find_regions_by_vision(line_crop):
    """Claude Vision 精確定位所有 LINE UI 元素"""
    import anthropic
    client = anthropic.Anthropic()

    lw, lh = line_crop.size
    tmp = os.path.join(TMPDIR, "line_vision_locate.png")
    line_crop.save(tmp, quality=95)
    with open(tmp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    prompt = (
        f"Image size: {lw}x{lh} pixels. This is LINE Desktop app.\n\n"
        "Find EXACT pixel coordinates for these 6 elements. BE EXTREMELY PRECISE:\n\n"
        "1. sidebar: The narrow vertical icon bar on the far LEFT side of the window. "
        "Contains icons for Friends, Chat, Add Friends, Timeline, Calls, Keep, etc. "
        "Usually about 40-60px wide with a different background color.\n\n"
        "2. search_bar: The search input field in the left panel area. "
        "Only the white/light text input area where user types to search.\n\n"
        "3. left_panel: The panel between the sidebar and the chat area. "
        "Contains either friend list, chat list, or add-friend options depending on current page. "
        "Starts from sidebar right edge, ends at the separator line before chat area.\n\n"
        "4. chat_title: The name/title at the top of the chat area (right side). "
        "Shows the group name or friend name of the current conversation. "
        "Just the name text area, not the entire header bar.\n\n"
        "5. chat_area: The main message area with conversation bubbles. "
        "Right side, between the chat title header and the input box.\n\n"
        "6. input_box: The text input field at the bottom right where user types messages. "
        "Only the white text area, not the attachment/emoji buttons.\n\n"
        "Also determine which page the left panel is currently showing:\n"
        '- "friend" if showing categorized lists (社群/群組/好友)\n'
        '- "chat" if showing recent conversation list\n'
        '- "add_friend" if showing search/create group/recommendations\n\n'
        "Return RAW JSON only. No markdown. No explanation:\n"
        '{"sidebar":{"l":0,"t":0,"r":0,"b":0},'
        '"search_bar":{"l":0,"t":0,"r":0,"b":0},'
        '"left_panel":{"l":0,"t":0,"r":0,"b":0},'
        '"chat_title":{"l":0,"t":0,"r":0,"b":0},'
        '"chat_area":{"l":0,"t":0,"r":0,"b":0},'
        '"input_box":{"l":0,"t":0,"r":0,"b":0},'
        '"current_page":"friend or chat or add_friend"}'
    )

    r = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=500,
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


def locate_line_regions(monitor=2):
    """
    主函數：找到 LINE 所有 UI 區域的精確座標

    流程：
    1. 截圖 + 找到 LINE 視窗
    2. 像素分析找 sidebar 寬度和分隔線
    3. Claude Vision 找所有元素的精確位置
    4. 交叉驗證 + 用已確認的邊界修正

    返回：
    {
        "sidebar":      {"left", "top", "right", "bottom", "center"},
        "search_bar":   {"left", "top", "right", "bottom", "center"},
        "left_panel":   {"left", "top", "right", "bottom", "center"},
        "chat_title":   {"left", "top", "right", "bottom", "center"},
        "chat_area":    {"left", "top", "right", "bottom", "center"},
        "input_box":    {"left", "top", "right", "bottom", "center"},
        "current_page": "friend" | "chat" | "add_friend",
        "separator_x":  int,
        "sidebar_width": int,
        "line_window":  {"left", "top", "right", "bottom"},
    }
    所有座標都是螢幕絕對座標（可以直接用 pyautogui.click）
    """
    _print("[line_locate] 開始定位...")

    # Step 1: 截圖
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
    lw, lh = line_crop.size
    _print(f"[line_locate] LINE: {lw}x{lh} at ({il},{it})-({ir},{ib})")

    # Step 2: 像素分析
    sidebar_w = find_sidebar_by_pixel(line_crop)
    sep_x = find_separator_by_pixel(line_crop, sidebar_w)
    _print(f"[line_locate] sidebar 寬度: {sidebar_w}, 分隔線: x={sep_x}")

    # Step 3: Claude Vision 定位
    vision = find_regions_by_vision(line_crop)
    current_page = vision.get("current_page", "unknown")
    _print(f"[line_locate] Vision: {json.dumps(vision, ensure_ascii=False)}")
    _print(f"[line_locate] 當前頁面: {current_page}")

    # Step 4: 組合結果
    def extract(key):
        v = vision[key]
        return {"left": v["l"], "top": v["t"], "right": v["r"], "bottom": v["b"]}

    sidebar = extract("sidebar")
    search_bar = extract("search_bar")
    left_panel = extract("left_panel")
    chat_title = extract("chat_title")
    input_box = extract("input_box")

    # 對話區用確認的邊界修正
    chat_area = {
        "left": sep_x,
        "top": chat_title.get("bottom", vision["chat_area"]["t"]),
        "right": lw,
        "bottom": input_box.get("top", vision["chat_area"]["b"]),
    }

    # 轉換為螢幕絕對座標
    sx_ratio = mon["width"] / full_img.size[0]
    sy_ratio = mon["height"] / full_img.size[1]

    def to_screen(region):
        sl = int(mon["left"] + (il + region["left"]) * sx_ratio)
        st = int(mon["top"] + (it + region["top"]) * sy_ratio)
        sr = int(mon["left"] + (il + region["right"]) * sx_ratio)
        sb = int(mon["top"] + (it + region["bottom"]) * sy_ratio)
        cx = (sl + sr) // 2
        cy = (st + sb) // 2
        return {"left": sl, "top": st, "right": sr, "bottom": sb, "center": (cx, cy)}

    result = {
        "sidebar": to_screen(sidebar),
        "search_bar": to_screen(search_bar),
        "left_panel": to_screen(left_panel),
        "chat_title": to_screen(chat_title),
        "chat_area": to_screen(chat_area),
        "input_box": to_screen(input_box),
        "current_page": current_page,
        "separator_x": sep_x,
        "sidebar_width": sidebar_w,
        "line_window": {"left": il, "top": it, "right": ir, "bottom": ib},
    }

    _print("[line_locate] 定位完成")
    for name in ["sidebar", "search_bar", "left_panel", "chat_title", "chat_area", "input_box"]:
        r = result[name]
        _print(f"  {name}: ({r['left']},{r['top']})-({r['right']},{r['bottom']}) center={r['center']}")
    _print(f"  current_page: {current_page}")

    return result


# === 直接執行時測試 ===
if __name__ == "__main__":
    regions = locate_line_regions(monitor=2)
    print("\n=== 結果 ===")
    _print(json.dumps(regions, ensure_ascii=False, indent=2, default=str))
