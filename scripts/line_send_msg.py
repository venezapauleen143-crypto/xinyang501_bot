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


def send_message(contact_name, message, monitor=2):
    """完整流程：定位 LINE → 搜尋好友 → 點進對話 → 發送訊息"""

    from line_locate import (
        locate_line_regions, switch_page, screenshot_line,
        find_line_window,
    )

    # ============================================================
    # Step 1: 定位 LINE 視窗 + 置前
    # ============================================================
    print(f"[Step 1] 定位 LINE 視窗...", flush=True)
    line = find_line_window()
    if not line:
        print("[ERROR] 找不到 LINE 視窗", flush=True)
        return False

    win32gui.SetForegroundWindow(line[0])
    time.sleep(0.5)

    regions = locate_line_regions(monitor)
    print(f"[Step 1] 當前頁面: {regions['current_page']}", flush=True)

    # ============================================================
    # Step 2: 切到好友頁
    # ============================================================
    if regions["current_page"] != "friend":
        print(f"[Step 2] 切到好友頁...", flush=True)
        regions = switch_page(regions, "friend", monitor)
    else:
        print(f"[Step 2] 已在好友頁", flush=True)

    # ============================================================
    # Step 3: 點搜尋欄 → 輸入好友名稱
    # ============================================================
    print(f"[Step 3] 搜尋好友: {contact_name}", flush=True)
    sx, sy = regions["search_bar"]["center"]
    pyautogui.click(sx, sy)
    time.sleep(0.5)

    # 清空搜尋欄
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.press("delete")
    time.sleep(0.1)

    # 輸入好友名稱
    pyperclip.copy(contact_name)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1.5)
    print(f"[Step 3] 已輸入: {contact_name}", flush=True)

    # ============================================================
    # Step 4: 從搜尋結果找到好友並點擊
    # ============================================================
    print(f"[Step 4] 找搜尋結果...", flush=True)

    # 重新截圖（搜尋結果已出現）
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)

    # 裁切左側面板（搜尋結果在這裡）
    lp = regions["left_panel"]
    # 轉回圖片內座標
    sx_ratio = full_img.size[0] / mon["width"]
    sy_ratio = full_img.size[1] / mon["height"]
    lp_il = int((lp["left"] - mon["left"]) * sx_ratio) - il
    lp_it = int((lp["top"] - mon["top"]) * sy_ratio) - it
    lp_ir = int((lp["right"] - mon["left"]) * sx_ratio) - il
    lp_ib = int((lp["bottom"] - mon["top"]) * sy_ratio) - it

    # 確保範圍合理
    lw, lh = line_crop.size
    lp_il = max(0, lp_il)
    lp_it = max(0, lp_it)
    lp_ir = min(lw, lp_ir)
    lp_ib = min(lh, lp_ib)

    search_result_crop = line_crop.crop((lp_il, lp_it, lp_ir, lp_ib))
    sw, sh = search_result_crop.size

    # 先用 PaddleOCR 找好友名稱
    found = False
    try:
        from line_locate import ocr_scan_panel
        ocr_items = ocr_scan_panel(search_result_crop)
        for item in ocr_items:
            # 模糊比對（處理簡繁體差異、OCR 辨識微小差異）
            from difflib import SequenceMatcher
            ratio = SequenceMatcher(None, contact_name, item["text"]).ratio()
            if contact_name in item["text"] or item["text"] in contact_name or ratio > 0.6:
                # 找到了，計算螢幕座標
                click_x = int(mon["left"] + (il + lp_il + sw // 2) * (mon["width"] / full_img.size[0]))
                click_y = int(mon["top"] + (it + lp_it + item["y"] + 15) * (mon["height"] / full_img.size[1]))
                pyautogui.click(click_x, click_y)
                time.sleep(1.0)
                print(f"[Step 4] OCR 找到 {item['text']}，點擊 ({click_x}, {click_y})", flush=True)
                found = True
                break
    except Exception as e:
        print(f"[Step 4] OCR 失敗: {e}", flush=True)

    # OCR 找不到就用 Vision
    if not found:
        print(f"[Step 4] OCR 找不到，啟動 Vision...", flush=True)
        import anthropic
        client = anthropic.Anthropic()

        tmp = os.path.join(TMPDIR, "line_search_result.png")
        search_result_crop.save(tmp, quality=95)
        with open(tmp, "rb") as f:
            b64 = base64.b64encode(f.read()).decode()

        r = client.messages.create(
            model="claude-sonnet-4-6", max_tokens=100,
            messages=[{"role": "user", "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
                {"type": "text", "text": (
                    f"This is LINE search results ({sw}x{sh}px). "
                    f"Find the contact named '{contact_name}'. "
                    f"Return the CENTER coordinates of that row. "
                    f"Raw JSON only: {{\"x\":0,\"y\":0}}"
                )}
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
            print(f"[Step 4] Vision 找到，點擊 ({click_x}, {click_y})", flush=True)
            found = True
        else:
            # 最後備援：點搜尋結果第一個項目
            first_y = int(mon["top"] + (it + lp_it + 80) * (mon["height"] / full_img.size[1]))
            first_x = int(mon["left"] + (il + lp_il + sw // 2) * (mon["width"] / full_img.size[0]))
            pyautogui.click(first_x, first_y)
            time.sleep(1.0)
            print(f"[Step 4] 備援：點第一個搜尋結果 ({first_x}, {first_y})", flush=True)
            found = True

    if not found:
        print("[ERROR] 找不到好友", flush=True)
        return False

    # ============================================================
    # Step 5: 確認進入對話視窗
    # ============================================================
    print(f"[Step 5] 確認對話視窗...", flush=True)
    time.sleep(0.5)

    # 重新定位，確認 chat_title 是不是目標好友
    regions = locate_line_regions(monitor)
    # 用 Vision 讀取 chat_title 的名字
    full_img2, line_crop2, (il2, it2, ir2, ib2), mon2 = screenshot_line(monitor)
    ct = regions["chat_title"]
    ct_il = int((ct["left"] - mon2["left"]) * (full_img2.size[0] / mon2["width"])) - il2
    ct_it = int((ct["top"] - mon2["top"]) * (full_img2.size[1] / mon2["height"])) - it2
    ct_ir = int((ct["right"] - mon2["left"]) * (full_img2.size[0] / mon2["width"])) - il2
    ct_ib = int((ct["bottom"] - mon2["top"]) * (full_img2.size[1] / mon2["height"])) - it2

    ct_il = max(0, ct_il)
    ct_it = max(0, ct_it)
    ct_ir = min(line_crop2.size[0], ct_ir)
    ct_ib = min(line_crop2.size[1], ct_ib)

    title_crop = line_crop2.crop((ct_il, ct_it, ct_ir, ct_ib))

    # OCR 讀 chat_title
    try:
        from line_locate import ocr_scan_panel
        title_items = ocr_scan_panel(title_crop)
        detected_name = " ".join(item["text"] for item in title_items) if title_items else ""
    except Exception:
        detected_name = ""

    if contact_name in detected_name or detected_name in contact_name:
        print(f"[Step 5] 確認正確: {detected_name}", flush=True)
    else:
        print(f"[Step 5] 聊天標題: {detected_name}（可能不完全匹配，繼續發送）", flush=True)

    # ============================================================
    # Step 5.5: 偵測是否為新好友（沒有聊天記錄，需先點「聊天」按鈕）
    # ============================================================
    # 新好友的對話區不會有輸入框，只有一個綠色「聊天」按鈕
    # 偵測方式：直接掃描對話區有沒有綠色按鈕（最可靠）
    print(f"[Step 5.5] 檢查是否需要點「聊天」按鈕...", flush=True)

    full_img3, line_crop3, (il3, it3, ir3, ib3), mon3 = screenshot_line(monitor)
    import numpy as np
    arr = np.array(line_crop3)
    lh3, lw3, _ = arr.shape
    sep = regions.get("separator_x", lw3 // 2)

    # 掃描右半部（對話區）找綠色像素
    green_points = []
    for y in range(lh3 // 4, lh3 * 3 // 4):
        for x in range(sep, lw3, 3):
            r, g, b = int(arr[y, x, 0]), int(arr[y, x, 1]), int(arr[y, x, 2])
            if g > 150 and g > r + 50 and g > b + 30 and r < 100:
                green_points.append((x, y))

    has_green_button = len(green_points) > 20  # 綠色按鈕至少有 20+ 像素

    if has_green_button:
        print(f"[Step 5.5] 偵測到綠色「聊天」按鈕（{len(green_points)} 像素），新好友！", flush=True)

        if green_points:
            # 取綠色區塊的中心
            avg_x = sum(p[0] for p in green_points) // len(green_points)
            avg_y = sum(p[1] for p in green_points) // len(green_points)

            sx_r = mon3["width"] / full_img3.size[0]
            sy_r = mon3["height"] / full_img3.size[1]
            btn_x = int(mon3["left"] + (il3 + avg_x) * sx_r)
            btn_y = int(mon3["top"] + (it3 + avg_y) * sy_r)

            print(f"[Step 5.5] 找到綠色「聊天」按鈕 at ({btn_x}, {btn_y})，點擊...", flush=True)
            pyautogui.click(btn_x, btn_y)
            time.sleep(1.5)

            # 重新定位（現在應該有輸入框了）
            regions = locate_line_regions(monitor)
            print(f"[Step 5.5] 對話視窗已開啟", flush=True)
        else:
            print(f"[Step 5.5] 有綠色像素但找不到按鈕中心，嘗試直接發送...", flush=True)
    else:
        print(f"[Step 5.5] 不是新好友（沒有綠色按鈕），直接發送", flush=True)

    # ============================================================
    # Step 6: 點輸入框 → 輸入訊息 → Enter 發送
    # ============================================================
    print(f"[Step 6] 發送訊息...", flush=True)
    ix, iy = regions["input_box"]["center"]
    pyautogui.click(ix, iy)
    time.sleep(0.3)
    pyperclip.copy(message)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1)

    print(f"[Step 6] 已送出: {message}", flush=True)

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
