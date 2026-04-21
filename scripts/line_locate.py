"""
LINE 電腦版 UI 定位模組 — 每次運行都用 Claude Vision 精確找到所有區域
不使用固定座標，每次都重新偵測

LINE 有三個主要左側面板模式（依 sidebar 圖示切換）：
- 好友頁（friend）：用戶名稱 + 社群/群組/好友分類列表
- 聊天頁（chat）：最近對話列表（全部/群組/社群分頁）
- 加好友頁（add_friend）：搜尋好友/建立群組/建立社群/官方推薦/你可能認識的人

定位區域：
1. sidebar          — 最左側功能列整體區域
2. sidebar_icons    — sidebar 每個圖示按鈕的個別座標（用於切換頁面）
   - btn_friend       好友頁按鈕（人形圖示）
   - btn_chat         聊天頁按鈕（對話框圖示）
   - btn_add_friend   加好友頁按鈕（人形+加號圖示）
   - btn_timeline     動態消息按鈕（時鐘/圓形圖示）
   - btn_call         通話按鈕（電話圖示）
   - btn_keep         Keep 按鈕（書籤圖示）
   - btn_more         更多按鈕（三點圖示）
3. search_bar       — 搜尋輸入框
4. left_panel       — 左側面板（內容依頁面切換）
5. chat_title       — 聊天標題（群組名/好友名）
6. chat_area        — 對話內容區
7. input_box        — 輸入框
8. current_page     — 目前在哪個頁面

使用方式：
    from line_locate import locate_line_regions, switch_page
    regions = locate_line_regions(monitor=2)

    # 取得任何區域的座標
    regions["search_bar"]["center"]       → 搜尋欄中心座標
    regions["input_box"]["center"]        → 輸入框中心座標
    regions["chat_area"]                  → 對話區完整座標

    # 取得 sidebar 按鈕座標（用於切換頁面）
    regions["sidebar_icons"]["btn_friend"]["center"]      → 好友頁按鈕
    regions["sidebar_icons"]["btn_chat"]["center"]        → 聊天頁按鈕
    regions["sidebar_icons"]["btn_add_friend"]["center"]  → 加好友頁按鈕

    # 切換頁面
    switch_page(regions, "chat")        → 切到聊天頁
    switch_page(regions, "friend")      → 切到好友頁
    switch_page(regions, "add_friend")  → 切到加好友頁
"""

import sys
import io
import os
import re
import json
import time
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
import pyautogui
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
            if "LINE" in title and win32gui.IsWindowVisible(hwnd):
                rect = win32gui.GetWindowRect(hwnd)
                w = rect[2] - rect[0]
                h = rect[3] - rect[1]
                if w > 300 and h > 300:
                    results.append((hwnd, title, cls, rect))
    win32gui.EnumWindows(callback, results)
    if not results:
        return None
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


def find_sidebar_by_pixel(line_crop):
    """像素分析找最左側功能列的寬度"""
    arr = np.array(line_crop)
    h, w, _ = arr.shape

    for x in range(30, min(80, w)):
        col = arr[h // 4: h * 3 // 4, x, :]
        avg_brightness = np.mean(col)
        prev_col = arr[h // 4: h * 3 // 4, x - 1, :]
        prev_brightness = np.mean(prev_col)
        if avg_brightness - prev_brightness > 30:
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
            r, g, b = int(row[x, 0]), int(row[x, 1]), int(row[x, 2])
            prev_r, prev_g, prev_b = int(row[x - 1, 0]), int(row[x - 1, 1]), int(row[x - 1, 2])
            brightness = (r + g + b) / 3
            prev_brightness = (prev_r + prev_g + prev_b) / 3
            if abs(brightness - prev_brightness) > 20 and 180 < brightness < 220:
                candidates.append(x)
                break

    if candidates:
        return sorted(candidates)[len(candidates) // 2]
    return w // 3


def find_regions_by_vision(line_crop):
    """Claude Vision 精確定位所有 LINE UI 元素 + sidebar 每個圖示按鈕"""
    import anthropic
    client = anthropic.Anthropic()

    lw, lh = line_crop.size
    tmp = os.path.join(TMPDIR, "line_vision_locate.png")
    line_crop.save(tmp, quality=95)
    with open(tmp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    prompt = (
        f"Image size: {lw}x{lh} pixels. This is LINE Desktop app.\n\n"
        "Find EXACT pixel coordinates for ALL of the following. BE EXTREMELY PRECISE.\n\n"
        "=== SIDEBAR ICONS (the narrow vertical bar on the far left) ===\n"
        "The sidebar contains clickable icon buttons arranged vertically. "
        "Find the CENTER POINT (x, y) of each icon button:\n\n"
        "- btn_friend: Friends icon (person silhouette, usually 1st from top)\n"
        "- btn_chat: Chat icon (speech bubble, usually 2nd from top)\n"
        "- btn_add_friend: Add Friend icon (person with + sign, usually 3rd from top)\n"
        "- btn_timeline: Timeline/News Feed icon (usually 4th)\n"
        "- btn_call: Call icon (phone shape, usually 5th)\n"
        "- btn_keep: Keep/Bookmark icon (flag or bookmark shape)\n"
        "- btn_line_official: LINE official icon (LINE logo, usually near bottom)\n"
        "- btn_more: More options (three dots '...' at very bottom)\n\n"
        "=== MAIN UI REGIONS (bounding boxes) ===\n\n"
        "1. sidebar: The entire sidebar bar area (left=0 to its right edge)\n\n"
        "2. search_bar: The search input field in the left panel. "
        "Only the text input area where user types.\n\n"
        "3. left_panel: The panel between sidebar and chat area. "
        "Contains friend list, chat list, or add-friend options.\n\n"
        "4. chat_title: The name/title text at top of chat area (right side). "
        "Just the name text, not entire header.\n\n"
        "5. chat_area: Main message area with conversation. "
        "Between chat title and input box.\n\n"
        "6. input_box: Text input field at bottom right. "
        "Only the white text area.\n\n"
        "=== CURRENT PAGE ===\n"
        "Which page is the left panel showing?\n"
        '- "friend" if showing categorized lists (社群/群組/好友)\n'
        '- "chat" if showing recent conversation list with tabs (全部/群組/社群)\n'
        '- "add_friend" if showing search friend/create group/recommendations\n\n'
        "Return RAW JSON only. No markdown. No explanation:\n"
        '{"sidebar":{"l":0,"t":0,"r":0,"b":0},'
        '"btn_friend":{"x":0,"y":0},'
        '"btn_chat":{"x":0,"y":0},'
        '"btn_add_friend":{"x":0,"y":0},'
        '"btn_timeline":{"x":0,"y":0},'
        '"btn_call":{"x":0,"y":0},'
        '"btn_keep":{"x":0,"y":0},'
        '"btn_line_official":{"x":0,"y":0},'
        '"btn_more":{"x":0,"y":0},'
        '"search_bar":{"l":0,"t":0,"r":0,"b":0},'
        '"left_panel":{"l":0,"t":0,"r":0,"b":0},'
        '"chat_title":{"l":0,"t":0,"r":0,"b":0},'
        '"chat_area":{"l":0,"t":0,"r":0,"b":0},'
        '"input_box":{"l":0,"t":0,"r":0,"b":0},'
        '"current_page":"friend or chat or add_friend"}'
    )

    r = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=800,
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
    3. Claude Vision 找所有元素 + sidebar 每個按鈕的精確位置
    4. 交叉驗證 + 用已確認的邊界修正

    返回：
    {
        "sidebar":        {"left", "top", "right", "bottom", "center"},
        "sidebar_icons":  {
            "btn_friend":       {"center": (x, y)},
            "btn_chat":         {"center": (x, y)},
            "btn_add_friend":   {"center": (x, y)},
            "btn_timeline":     {"center": (x, y)},
            "btn_call":         {"center": (x, y)},
            "btn_keep":         {"center": (x, y)},
            "btn_line_official": {"center": (x, y)},
            "btn_more":         {"center": (x, y)},
        },
        "search_bar":     {"left", "top", "right", "bottom", "center"},
        "left_panel":     {"left", "top", "right", "bottom", "center"},
        "chat_title":     {"left", "top", "right", "bottom", "center"},
        "chat_area":      {"left", "top", "right", "bottom", "center"},
        "input_box":      {"left", "top", "right", "bottom", "center"},
        "current_page":   "friend" | "chat" | "add_friend",
        "separator_x":    int,
        "sidebar_width":  int,
        "line_window":    {"left", "top", "right", "bottom"},
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
    # 螢幕座標轉換
    sx_ratio = mon["width"] / full_img.size[0]
    sy_ratio = mon["height"] / full_img.size[1]

    def to_screen_region(v):
        """bounding box → 螢幕絕對座標"""
        sl = int(mon["left"] + (il + v["l"]) * sx_ratio)
        st = int(mon["top"] + (it + v["t"]) * sy_ratio)
        sr = int(mon["left"] + (il + v["r"]) * sx_ratio)
        sb = int(mon["top"] + (it + v["b"]) * sy_ratio)
        cx = (sl + sr) // 2
        cy = (st + sb) // 2
        return {"left": sl, "top": st, "right": sr, "bottom": sb, "center": (cx, cy)}

    def to_screen_point(v):
        """center point → 螢幕絕對座標"""
        sx_abs = int(mon["left"] + (il + v["x"]) * sx_ratio)
        sy_abs = int(mon["top"] + (it + v["y"]) * sy_ratio)
        return {"center": (sx_abs, sy_abs)}

    # 主要區域
    sidebar = to_screen_region(vision["sidebar"])
    search_bar = to_screen_region(vision["search_bar"])
    left_panel = to_screen_region(vision["left_panel"])
    chat_title = to_screen_region(vision["chat_title"])
    input_box = to_screen_region(vision["input_box"])

    # 對話區用確認的邊界修正
    chat_area_raw = {
        "l": sep_x,
        "t": vision["chat_title"]["b"],
        "r": lw,
        "b": vision["input_box"]["t"],
    }
    chat_area = to_screen_region(chat_area_raw)

    # Sidebar 每個按鈕的座標
    btn_names = [
        "btn_friend", "btn_chat", "btn_add_friend", "btn_timeline",
        "btn_call", "btn_keep", "btn_line_official", "btn_more"
    ]
    sidebar_icons = {}
    for btn in btn_names:
        if btn in vision and "x" in vision[btn]:
            sidebar_icons[btn] = to_screen_point(vision[btn])
        else:
            sidebar_icons[btn] = {"center": (0, 0)}

    result = {
        "sidebar": sidebar,
        "sidebar_icons": sidebar_icons,
        "search_bar": search_bar,
        "left_panel": left_panel,
        "chat_title": chat_title,
        "chat_area": chat_area,
        "input_box": input_box,
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
    _print(f"  sidebar_icons:")
    for btn, info in sidebar_icons.items():
        _print(f"    {btn}: {info['center']}")

    return result


def switch_page(regions, target_page, monitor=2):
    """
    切換 LINE 左側面板到指定頁面

    參數：
        regions: locate_line_regions() 的返回值
        target_page: "friend" | "chat" | "add_friend"
        monitor: 螢幕編號

    返回：
        更新後的 regions（重新定位）
    """
    page_to_btn = {
        "friend": "btn_friend",
        "chat": "btn_chat",
        "add_friend": "btn_add_friend",
    }

    if target_page not in page_to_btn:
        raise ValueError(f"未知頁面: {target_page}，支援: friend, chat, add_friend")

    # 如果已經在目標頁面，不需要切換
    if regions.get("current_page") == target_page:
        _print(f"[line_locate] 已經在 {target_page} 頁面，不需切換")
        return regions

    btn_name = page_to_btn[target_page]
    btn = regions["sidebar_icons"].get(btn_name)

    if not btn or btn["center"] == (0, 0):
        raise RuntimeError(f"找不到 {btn_name} 按鈕座標，無法切換")

    cx, cy = btn["center"]
    _print(f"[line_locate] 切換到 {target_page}，點擊 {btn_name} at ({cx}, {cy})")
    pyautogui.click(cx, cy)
    time.sleep(1.0)

    # 重新定位（頁面已切換，左側面板內容不同）
    new_regions = locate_line_regions(monitor)
    _print(f"[line_locate] 切換完成，當前頁面: {new_regions['current_page']}")
    return new_regions


# === 直接執行時測試 ===
if __name__ == "__main__":
    regions = locate_line_regions(monitor=2)
    print("\n=== 完整結果 ===")
    _print(json.dumps(regions, ensure_ascii=False, indent=2, default=str))
