import sys, io, ctypes
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

# 強制 DPI awareness，讓 mss 和 win32gui 座標一致
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(0)
except Exception:
    pass

import win32gui, mss, os, base64, time, pyautogui, pyperclip, numpy as np, hashlib
from PIL import Image
from datetime import datetime
import anthropic
from dotenv import load_dotenv
from pathlib import Path

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
client = anthropic.Anthropic()
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"
STOP_TIME = "14:00"

PERSONA = """你是小牛馬，一個嘴賤幽默的男生，說話像真人一樣自然。絕對不能暴露自己是AI、機器人、程式。
風格：
- 台灣口語：然後、就是、對啊、沒有啦、還好啦
- 嘴賤互嗆但不惡意，兄弟之間的感覺
- 會吐槽、自嘲、幽默
- 最多5句不超過120字
- 不用emoji
- 說話像真人朋友聊天，有情緒、有個性、會生氣會開玩笑
- 絕對禁止說出「我是AI」「我是機器人」「我是程式」「作為AI」等任何暴露身份的話"""

def find_telegram():
    results = []
    def callback(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            cls = win32gui.GetClassName(hwnd)
            if "Qt" in cls and "QWindow" in cls:
                title = win32gui.GetWindowText(hwnd)
                rect = win32gui.GetWindowRect(hwnd)
                results.append((hwnd, title, rect))
    win32gui.EnumWindows(callback, results)
    return results

def init_regions():
    """啟動時找一次窗口+像素找分隔線"""
    tg_list = find_telegram()
    if not tg_list:
        return None
    main_tg = max(tg_list, key=lambda x: (x[2][2]-x[2][0]) * (x[2][3]-x[2][1]))
    hwnd, title, (wl, wt, wr, wb) = main_tg
    print(f"Telegram: \"{title}\" at ({wl},{wt})-({wr},{wb})", flush=True)

    with mss.mss() as sct:
        mon = sct.monitors[2]
        img = sct.grab(mon)
        pil = Image.frombytes("RGB", img.size, img.rgb)
        iw, ih = pil.size
        print(f"mss: {iw}x{ih}", flush=True)

    sx = iw / mon["width"]
    sy = ih / mon["height"]
    il = int((wl - mon["left"]) * sx)
    it = int((wt - mon["top"]) * sy)
    ir = int((wr - mon["left"]) * sx)
    ib = int((wb - mon["top"]) * sy)

    # 裁切 Telegram 窗口，逐列找最長暗色垂直線 = 分隔線
    arr = np.array(pil)
    tg_crop = arr[it:ib, il:ir, :]
    h, w, _ = tg_crop.shape

    # 多行掃描找白色→綠色壁紙交界（排除藍色高亮）
    candidates = []
    for check_y in range(int(h*0.3), int(h*0.8), int(h*0.05)):
        row = tg_crop[check_y, :, :]
        for x in range(w // 5, w * 3 // 4):
            r, g, b = int(row[x, 0]), int(row[x, 1]), int(row[x, 2])
            if g > r + 5 and g > 150 and r < 220 and b < g:
                candidates.append(x)
                break
    best_x = sorted(candidates)[len(candidates)//2] if candidates else w // 2

    # 對話區 = 分隔線右邊，包含標題列和輸入框
    chat_region = (il + best_x, it, ir, ib)

    # 輸入框：像素分析找底部白色區域（分隔線右邊）
    input_y1 = h - 1
    for y in range(h - 1, h - 60, -1):
        row = tg_crop[y, best_x:w, :]
        white = np.sum((row[:, 0] > 240) & (row[:, 1] > 240) & (row[:, 2] > 240))
        if white > (w - best_x) * 0.5:
            input_y1 = y
        elif input_y1 < h - 1:
            break
    i_x1, i_x2 = best_x, w
    mid_iy = (input_y1 + h - 1) // 2
    row = tg_crop[mid_iy, best_x:w, :]
    for x in range(0, w - best_x):
        if row[x, 0] > 240 and row[x, 1] > 240 and row[x, 2] > 240:
            i_x1 = best_x + x; break
    for x in range(w - best_x - 1, 0, -1):
        if row[x, 0] > 240 and row[x, 1] > 240 and row[x, 2] > 240:
            i_x2 = best_x + x; break
    input_x = int(mon["left"] + (il + (i_x1 + i_x2) // 2) / sx)
    input_y = int(mon["top"] + (it + (input_y1 + h - 1) // 2) / sy)
    input_pos = (input_x, input_y)

    print(f"分隔線: x={best_x}/{w}, 對話區: {chat_region}, 輸入框: {input_pos}", flush=True)
    return chat_region, input_pos

def grab_conv(region):
    """只截圖+裁切對話區"""
    cl, ct, cr, cb = region
    with mss.mss() as sct:
        img = sct.grab(sct.monitors[2])
        pil = Image.frombytes("RGB", img.size, img.rgb)
    return pil.crop((cl, ct, cr, cb))

def conv_hash(conv_img):
    small = conv_img.resize((80, 40))
    return hashlib.md5(small.tobytes()).hexdigest()

def gen_reply(conv_img):
    """送 Claude 分析對話並回覆"""
    p = os.path.join(TMPDIR, "tg_auto.png")
    conv_img.save(p)
    with open(p, "rb") as f:
        d = base64.b64encode(f.read()).decode()
    r = client.messages.create(model="claude-sonnet-4-6", max_tokens=200, system=PERSONA,
        messages=[{"role":"user","content":[
        {"type":"image","source":{"type":"base64","media_type":"image/png","data":d}},
        {"type":"text","text":"""這是 Telegram 對話截圖，請依照以下步驟分析：

1. 辨識對話結構：
   - 綠色氣泡（靠右）= 我（小牛馬）發的
   - 白色氣泡（靠左）= 對方發的

2. 分析上下文：從上到下讀完整段對話，理解主題和情緒

3. 判斷是否回覆：
   - 最底部是白色氣泡 → 需要回覆
   - 最底部是綠色氣泡 → 只回一個字「等」

4. 如果需要回覆：根據整段上下文回應，不是只看最後一條。對方問問題就回答，嗆人就幽默化解，聊天就接話。如果對方連發多條，整體理解後一次回覆。

只回覆要發送的文字，不要加任何格式或解釋。"""}
    ]}])
    return r.content[0].text.strip()

def send(msg, input_pos):
    ix, iy = input_pos
    pyautogui.click(ix, iy)
    time.sleep(0.3)
    pyperclip.copy(msg)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1)

# === 啟動 ===
result = init_regions()
if not result:
    print("找不到 Telegram")
    sys.exit(1)

chat_region, input_pos = result

# 先截一張確認對話區正確
conv = grab_conv(chat_region)
last_hash = conv_hash(conv)
sent_hashes = set()
sent_hashes.add(last_hash)

cooldown_until = 0  # 冷卻時間戳

print(f"[{datetime.now().strftime('%H:%M:%S')}] 監控啟動 → {STOP_TIME}", flush=True)

while datetime.now().strftime("%H:%M") < STOP_TIME:
    time.sleep(6)
    try:
        # 冷卻期間跳過
        if time.time() < cooldown_until:
            continue

        conv = grab_conv(chat_region)
        h = conv_hash(conv)

        if h != last_hash:
            t = datetime.now().strftime("%H:%M:%S")
            print(f"[{t}] 偵測到變化", flush=True)

            reply = gen_reply(conv)
            reply = reply.encode("cp950", errors="ignore").decode("cp950")

            if reply and reply != "等" and len(reply) > 2:
                print(f"[{t}] → {reply}", flush=True)
                send(reply, input_pos)
                # 10 秒冷卻
                cooldown_until = time.time() + 10
                print(f"[{t}] 冷卻 10 秒", flush=True)
                # 冷卻結束後重新截圖當基準
                time.sleep(10)
                conv = grab_conv(chat_region)
                last_hash = conv_hash(conv)
            else:
                last_hash = h
        else:
            last_hash = h
    except Exception as e:
        print(f"[ERR] {e}", flush=True)
        time.sleep(5)

print("DONE", flush=True)
