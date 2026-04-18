"""
Telegram 傳訊息腳本
用法：python tg_send_msg.py <聯絡人名稱> <訊息內容>
範例：python tg_send_msg.py 宇欣 你好啊
"""
import sys, io, ctypes
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(0)
except Exception:
    pass

import win32gui, mss, os, time, base64, numpy as np
import pyautogui, pyperclip
from PIL import Image
import anthropic
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
client = anthropic.Anthropic()
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"


def find_telegram():
    """用 win32gui 找 Telegram 窗口"""
    results = []
    def cb(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            cls = win32gui.GetClassName(hwnd)
            if "Qt" in cls and "QWindow" in cls:
                results.append((hwnd, win32gui.GetWindowText(hwnd), win32gui.GetWindowRect(hwnd)))
    win32gui.EnumWindows(cb, results)
    if not results:
        return None
    return max(results, key=lambda x: (x[2][2]-x[2][0]) * (x[2][3]-x[2][1]))


def get_tg_regions(tg_info):
    """取得 Telegram 各區域座標"""
    hwnd, title, (wl, wt, wr, wb) = tg_info

    with mss.mss() as sct:
        mon = sct.monitors[2]
        img = sct.grab(mon)
        pil = Image.frombytes("RGB", img.size, img.rgb)
        iw, ih = pil.size

    sx = iw / mon["width"]
    sy = ih / mon["height"]
    il = int((wl - mon["left"]) * sx)
    it = int((wt - mon["top"]) * sy)
    ir = int((wr - mon["left"]) * sx)
    ib = int((wb - mon["top"]) * sy)

    # 找分隔線
    arr = np.array(pil.crop((il, it, ir, ib)))
    th, tw, _ = arr.shape
    candidates = []
    for check_y in range(int(th*0.3), int(th*0.8), int(th*0.05)):
        row = arr[check_y, :, :]
        for x in range(tw // 5, tw * 3 // 4):
            r, g, b = int(row[x, 0]), int(row[x, 1]), int(row[x, 2])
            if g > r + 5 and g > 150 and r < 220 and b < g:
                candidates.append(x)
                break
    split_x = sorted(candidates)[len(candidates)//2] if candidates else tw // 2

    return {
        "window": (wl, wt, wr, wb),
        "img_region": (il, it, ir, ib),
        "split_x": split_x,
        "search_x": wl + int((wr - wl) * 0.15),
        "search_y": wt + 40,
        "input_x": (wl + wr) // 2 + int((wr - wl) * 0.15),
        "input_y": wb - 25,
    }


def screenshot_conv(regions):
    """截取對話區"""
    il, it, ir, ib = regions["img_region"]
    split_x = regions["split_x"]
    with mss.mss() as sct:
        img = sct.grab(sct.monitors[2])
        pil = Image.frombytes("RGB", img.size, img.rgb)
    return pil.crop((il + split_x, it, ir, ib))


def send_message(contact_name, message):
    """完整流程：找 Telegram → 搜尋聯絡人 → 點進去 → 打字 → 送出"""

    # 1. 找 Telegram
    tg = find_telegram()
    if not tg:
        print("❌ 找不到 Telegram 窗口")
        return False
    print(f"✅ 找到 Telegram: {tg[1]}")

    regions = get_tg_regions(tg)
    wl, wt, wr, wb = regions["window"]

    # 2. 點搜尋框
    pyautogui.click(regions["search_x"], regions["search_y"])
    time.sleep(0.5)
    print("✅ 點擊搜尋框")

    # 3. 清空 + 輸入聯絡人名稱
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.press("delete")
    time.sleep(0.1)
    pyperclip.copy(contact_name)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1.5)
    print(f"✅ 搜尋: {contact_name}")

    # 4. 截圖讓 Claude 找搜尋結果中的聯絡人
    with mss.mss() as sct:
        img = sct.grab(sct.monitors[2])
        pil = Image.frombytes("RGB", img.size, img.rgb)
    il, it, ir, ib = regions["img_region"]
    search_area = pil.crop((il, it + 50, il + regions["split_x"], ib))
    search_path = os.path.join(TMPDIR, "tg_search.png")
    search_area.save(search_path)

    with open(search_path, "rb") as f:
        d = base64.b64encode(f.read()).decode()

    resp = client.messages.create(model="claude-sonnet-4-6", max_tokens=100,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": d}},
            {"type": "text", "text": f"這是 Telegram 搜尋結果列表截圖({search_area.width}x{search_area.height}px)。找到名稱為「{contact_name}」的聯絡人，回傳它的中心座標。只回 JSON: {{\"x\":數字,\"y\":數字}}"}
        ]}])

    import json, re
    match = re.search(r'\{.*?\}', resp.content[0].text, re.DOTALL)
    if match:
        pos = json.loads(match.group())
        # 座標是相對於 search_area 的，要轉成螢幕絕對座標
        abs_x = wl + int(pos["x"] * (wr - wl) / search_area.width * 0.4)
        abs_y = wt + 50 + int(pos["y"] * (wb - wt) / search_area.height)
        pyautogui.click(abs_x, abs_y)
        time.sleep(0.8)
        print(f"✅ 點擊聯絡人 ({abs_x}, {abs_y})")
    else:
        # 備用：直接點搜尋結果第一個
        pyautogui.click(regions["search_x"], regions["search_y"] + 60)
        time.sleep(0.8)
        print("⚠️ Vision 找不到，點第一個搜尋結果")

    # 5. 點輸入框
    # 重新取得區域（因為可能切換了聯絡人）
    tg = find_telegram()
    regions = get_tg_regions(tg)
    pyautogui.click(regions["input_x"], regions["input_y"])
    time.sleep(0.3)
    print("✅ 點擊輸入框")

    # 6. 輸入訊息並送出
    pyperclip.copy(message)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1)
    print(f"✅ 已送出: {message}")

    # 7. 截圖確認
    conv = screenshot_conv(regions)
    conv.save(os.path.join(TMPDIR, "tg_send_confirm.png"))
    print(f"✅ 確認截圖: {TMPDIR}/tg_send_confirm.png")

    return True


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        sys.exit(0)

    contact = sys.argv[1]
    msg = " ".join(sys.argv[2:])
    send_message(contact, msg)
