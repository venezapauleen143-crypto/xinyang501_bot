"""
Telegram UI 定位模組 — 每次運行都用 OCR + Claude Vision 精確找到所有區域
不使用固定座標，不偷懶，每次都重新偵測

定位的 5 個區域：
1. search_bar   — 搜尋輸入框（不含漢堡選單和圖示）
2. friend_name  — 好友名稱文字（只有名字，不含按鈕圖示）
3. chat_area    — 對話區（從好友名稱下方到輸入框上方）
4. input_box    — 輸入框（只有白色文字輸入區，不含按鈕）
5. contact_list — 聯絡人清單（左側欄）

使用方式：
    from tg_locate import locate_telegram_regions
    regions = locate_telegram_regions(monitor=2)
    # regions["search_bar"] = {"left": x, "top": y, "right": x, "bottom": y, "center": (cx, cy)}
"""

import sys
import io
import os
import re
import json
import base64
import numpy as np

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import win32gui
import mss
import pytesseract
from PIL import Image, ImageDraw
from dotenv import load_dotenv
from pathlib import Path

# Setup
load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
pytesseract.pytesseract.tesseract_cmd = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"


def find_telegram_window():
    """用 win32gui 找到 Telegram 主視窗"""
    results = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            cls = win32gui.GetClassName(hwnd)
            if "Qt" in cls and "QWindow" in cls:
                rect = win32gui.GetWindowRect(hwnd)
                results.append((hwnd, win32gui.GetWindowText(hwnd), rect))
    win32gui.EnumWindows(callback, results)
    if not results:
        return None
    # 取最大的 Qt 視窗
    main = max(results, key=lambda x: (x[2][2] - x[2][0]) * (x[2][3] - x[2][1]))
    return main


def screenshot_telegram(monitor=2):
    """截圖指定螢幕，裁切出 Telegram 區域"""
    tg = find_telegram_window()
    if not tg:
        raise RuntimeError("找不到 Telegram 視窗")

    hwnd, title, (wl, wt, wr, wb) = tg

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

    tg_crop = pil.crop((il, it, ir, ib))

    return pil, tg_crop, (il, it, ir, ib), mon


def find_separator_by_pixel(tg_crop):
    """像素分析找分隔線（聯絡人清單和對話區的分界）"""
    arr = np.array(tg_crop)
    h, w, _ = arr.shape
    candidates = []
    for check_y in range(int(h * 0.3), int(h * 0.8), int(h * 0.05)):
        row = arr[check_y, :, :]
        for x in range(w // 5, w * 3 // 4):
            r, g, b = int(row[x, 0]), int(row[x, 1]), int(row[x, 2])
            if g > r + 5 and g > 150 and r < 220 and b < g:
                candidates.append(x)
                break
    if candidates:
        return sorted(candidates)[len(candidates) // 2]
    return w // 2


def find_search_by_ocr(tg_crop):
    """OCR 找「搜尋」文字的精確位置"""
    ocr_data = pytesseract.image_to_data(
        tg_crop, lang="chi_tra+eng", output_type=pytesseract.Output.DICT
    )
    for i, txt in enumerate(ocr_data["text"]):
        t = txt.strip()
        conf = ocr_data["conf"][i]
        if conf > 25 and t:
            if "搜尋" in t or "搜索" in t or "search" in t.lower():
                return {
                    "left": ocr_data["left"][i],
                    "top": ocr_data["top"][i],
                    "width": ocr_data["width"][i],
                    "height": ocr_data["height"][i],
                    "text": t,
                }
    return None


def find_regions_by_vision(tg_crop):
    """Claude Vision 精確定位所有 UI 元素"""
    import anthropic
    client = anthropic.Anthropic()

    tg_w, tg_h = tg_crop.size
    tmp = os.path.join(TMPDIR, "tg_vision_locate.png")
    tg_crop.save(tmp, quality=95)
    with open(tmp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    prompt = (
        f"Image size: {tg_w}x{tg_h} pixels. This is Telegram Desktop.\n\n"
        "Find EXACT pixel coordinates for these 5 elements. BE EXTREMELY PRECISE:\n\n"
        "1. search_bar: ONLY the white search input field (where user types to search). "
        "Do NOT include the hamburger menu (3 horizontal lines) on the left. "
        "Do NOT include icons on the right. Just the text input area.\n\n"
        "2. friend_name: ONLY the name text of the chat partner at the top of conversation. "
        "Just the rectangle tightly around the NAME TEXT itself. "
        "Do NOT include the entire header bar, icons, buttons, or phone/video call buttons. "
        "It should be a small narrow rectangle.\n\n"
        "3. chat_area: The message area with chat bubbles. "
        "Left edge starts at the separator line between contact list and chat. "
        "Right edge goes to the window right edge. "
        "Top starts immediately below the header/friend name area. "
        "Bottom goes to immediately above the input box area.\n\n"
        "4. input_box: ONLY the white text input field at bottom where user types messages. "
        "Do NOT include emoji button on left, attachment button, or send button on right. "
        "Just the white text area.\n\n"
        "5. contact_list: Left sidebar with all chats. Below search bar to bottom.\n\n"
        "Return RAW JSON only. No markdown. No explanation. No code blocks:\n"
        '{"search_bar":{"l":0,"t":0,"r":0,"b":0},'
        '"friend_name":{"l":0,"t":0,"r":0,"b":0},'
        '"chat_area":{"l":0,"t":0,"r":0,"b":0},'
        '"input_box":{"l":0,"t":0,"r":0,"b":0},'
        '"contact_list":{"l":0,"t":0,"r":0,"b":0}}'
    )

    r = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=400,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
            {"type": "text", "text": prompt},
        ]}],
    )
    resp = r.content[0].text.strip()
    # Clean markdown if present
    if resp.startswith("```"):
        resp = re.sub(r"^```(?:json)?\s*", "", resp)
        resp = re.sub(r"\s*```$", "", resp)
    return json.loads(resp)


def locate_telegram_regions(monitor=2, debug=False):
    """
    主函數：找到 Telegram 所有 UI 區域的精確座標

    流程：
    1. 截圖 + 找到 Telegram 視窗
    2. 像素分析找分隔線
    3. OCR 找「搜尋」文字（方案 B）
    4. Claude Vision 找所有元素的精確位置（方案 C）
    5. 交叉驗證 + 用已確認的邊界修正 chat_area

    返回：
    {
        "search_bar":   {"left", "top", "right", "bottom", "center"},
        "friend_name":  {"left", "top", "right", "bottom", "center"},
        "chat_area":    {"left", "top", "right", "bottom", "center"},
        "input_box":    {"left", "top", "right", "bottom", "center"},
        "contact_list": {"left", "top", "right", "bottom", "center"},
        "separator_x":  int,
        "tg_window":    {"left", "top", "right", "bottom"},
    }
    所有座標都是螢幕絕對座標（可以直接用 pyautogui.click）
    """
    print("[tg_locate] 開始定位...", flush=True)

    # Step 1: 截圖
    full_img, tg_crop, (il, it, ir, ib), mon = screenshot_telegram(monitor)
    tg_w, tg_h = tg_crop.size
    print(f"[tg_locate] Telegram: {tg_w}x{tg_h} at ({il},{it})-({ir},{ib})", flush=True)

    # Step 2: 像素分析找分隔線
    sep_x = find_separator_by_pixel(tg_crop)
    print(f"[tg_locate] 分隔線: x={sep_x}", flush=True)

    # Step 3: OCR 找搜尋欄
    ocr_search = find_search_by_ocr(tg_crop)
    if ocr_search:
        print(f"[tg_locate] OCR 找到「{ocr_search['text']}」at ({ocr_search['left']},{ocr_search['top']})", flush=True)

    # Step 4: Claude Vision 找所有元素
    vision = find_regions_by_vision(tg_crop)
    print(f"[tg_locate] Vision: {json.dumps(vision, ensure_ascii=False)}", flush=True)

    # Step 5: 組合結果，交叉驗證
    # 搜尋欄：優先用 OCR（更精確），Vision 當備援
    if ocr_search:
        sb = {
            "left": ocr_search["left"],
            "top": ocr_search["top"] - 3,
            "right": ocr_search["left"] + ocr_search["width"] + 10,
            "bottom": ocr_search["top"] + ocr_search["height"] + 3,
        }
    else:
        v = vision["search_bar"]
        sb = {"left": v["l"], "top": v["t"], "right": v["r"], "bottom": v["b"]}

    # 好友名稱：用 Vision（OCR 不一定能找到名字）
    v_fn = vision["friend_name"]
    fn = {"left": v_fn["l"], "top": v_fn["t"], "right": v_fn["r"], "bottom": v_fn["b"]}
    # 驗證：名字框不應該太大
    if (fn["bottom"] - fn["top"]) > 40 or (fn["right"] - fn["left"]) > 400:
        print("[tg_locate] WARNING: friend_name 框太大，可能不準確", flush=True)

    # 輸入框：用 Vision
    v_ib = vision["input_box"]
    ib_box = {"left": v_ib["l"], "top": v_ib["t"], "right": v_ib["r"], "bottom": v_ib["b"]}

    # 對話區：不信任 Vision 的值，用確認過的邊界計算
    # top = friend_name 的 bottom（好友名稱下方）
    # bottom = input_box 的 top（輸入框上方）
    # left = 分隔線
    # right = Telegram 視窗右邊界
    ca = {
        "left": sep_x,
        "top": fn["bottom"],
        "right": tg_w,
        "bottom": ib_box["top"],
    }

    # 聯絡人清單：用分隔線和 Vision
    v_cl = vision["contact_list"]
    cl = {"left": 0, "top": v_cl["t"], "right": sep_x, "bottom": tg_h}

    # 轉換為螢幕絕對座標（可直接用 pyautogui.click）
    # tg_crop 的座標 + Telegram 在螢幕中的偏移
    sx_ratio = mon["width"] / (full_img.size[0])  # image px → screen px
    sy_ratio = mon["height"] / (full_img.size[1])

    def to_screen(region):
        """tg_crop 內部座標 → 螢幕絕對座標"""
        sl = int(mon["left"] + (il + region["left"]) * sx_ratio)
        st = int(mon["top"] + (it + region["top"]) * sy_ratio)
        sr = int(mon["left"] + (il + region["right"]) * sx_ratio)
        sb = int(mon["top"] + (it + region["bottom"]) * sy_ratio)
        cx = (sl + sr) // 2
        cy = (st + sb) // 2
        return {"left": sl, "top": st, "right": sr, "bottom": sb, "center": (cx, cy)}

    result = {
        "search_bar": to_screen(sb),
        "friend_name": to_screen(fn),
        "chat_area": to_screen(ca),
        "input_box": to_screen(ib_box),
        "contact_list": to_screen(cl),
        "separator_x": sep_x,
        "tg_window": {"left": il, "top": it, "right": ir, "bottom": ib},
    }

    print("[tg_locate] 定位完成", flush=True)
    for name in ["search_bar", "friend_name", "chat_area", "input_box", "contact_list"]:
        r = result[name]
        print(f"  {name}: ({r['left']},{r['top']})-({r['right']},{r['bottom']}) center={r['center']}", flush=True)

    # Debug: 畫出來存檔
    if debug:
        draw = ImageDraw.Draw(full_img)
        styles = {
            "search_bar": "orange",
            "contact_list": "yellow",
            "friend_name": "black",
            "chat_area": "red",
            "input_box": "cyan",
        }
        for name, color in styles.items():
            r = result[name]
            # 轉回 image 座標畫圖
            draw.rectangle(
                [(il + sb["left"] if name == "search_bar" else il + locals().get(name, ca)["left"],
                  it + sb["top"] if name == "search_bar" else it + locals().get(name, ca)["top"]),
                 (il + sb["right"] if name == "search_bar" else il + locals().get(name, ca)["right"],
                  it + sb["bottom"] if name == "search_bar" else it + locals().get(name, ca)["bottom"])],
                outline=color, width=3
            )
        out = os.path.join(TMPDIR, "tg_locate_debug.png")
        full_img.save(out)
        print(f"  debug 截圖: {out}", flush=True)

    return result


# === 直接執行時測試 ===
if __name__ == "__main__":
    regions = locate_telegram_regions(monitor=2, debug=False)
    print("\n=== 結果 ===")
    print(json.dumps(regions, ensure_ascii=False, indent=2, default=str))
