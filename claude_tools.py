"""
Claude Code 工具腳本 - 整合小牛馬所有技能
用法：python claude_tools.py <tool> [參數...]

tools:
  weather <城市>                查詢天氣
  stock <代號> [期間]           查詢股票（期間: 1d 5d 1mo 3mo 6mo 1y，預設 1mo）
  image <prompt> [文字]         生成圖片（prompt 用英文，文字為疊加中文）
  screenshot                    截圖存到桌面
  click <x> <y>                 點擊
  double_click <x> <y>          雙擊
  right_click <x> <y>           右鍵點擊
  move <x> <y>                  移動滑鼠
  type <文字>                   輸入文字
  press <按鍵>                  按下按鍵
  open <程式>                   開啟程式
  scroll <up|down> [格數]       滾動
  pos                           取得目前滑鼠座標
  bot_status                    查看 bot 是否在執行
  bot_restart                   重啟 bot
  schedule_list                 列出所有排程任務
  schedule_add <名稱> <時間HH:MM> <腳本路徑>   新增每日排程（並設定可喚醒）
  schedule_del <名稱>           刪除排程任務
  memory_save <chat_id> <內容>  儲存長期記憶
  memory_list <chat_id>         列出長期記憶
  memory_del <chat_id> <id>     刪除指定記憶
  vision [問題]                 截圖並讓 Claude 分析畫面內容
  find_image <圖片路徑>         在螢幕上找到指定圖片的位置
  browser <動作> [參數]         瀏覽器自動化（open/click/type/get_text/screenshot/goto/close）
  window_list                   列出所有視窗
  window_focus <標題關鍵字>     切換到指定視窗
  window_close <標題關鍵字>     關閉指定視窗
  window_min <標題關鍵字>       最小化視窗
  window_max <標題關鍵字>       最大化視窗
  hotkey <按鍵組合>             執行組合鍵（如 ctrl+c, alt+tab）
  clipboard_get                 讀取剪貼簿內容
  clipboard_set <內容>          寫入剪貼簿
  file_list <路徑>              列出資料夾內容
  file_read <路徑>              讀取文字檔案
  file_write <路徑> <內容>      寫入文字檔案
  file_delete <路徑>            刪除檔案或資料夾
  file_copy <來源> <目標>       複製檔案
  file_move <來源> <目標>       移動/重新命名檔案
  file_search <資料夾> <關鍵字> 搜尋檔案
  sysinfo                       查看 CPU/記憶體/磁碟使用率
  process_list                  列出所有執行中程序
  process_kill <名稱或PID>      結束程序
  notify <標題> <訊息>          發送 Windows 桌面通知
  tts <文字>                    文字轉語音朗讀
  record_start                  開始錄製滑鼠鍵盤動作
  record_stop                   停止錄製並儲存
  record_play <檔案>            重播錄製的動作
  email <收件人> <主旨> <內容>  發送 Email
"""

import sys
import os
import io
import time
import sqlite3
import subprocess
import requests
import pyautogui
from pathlib import Path
from datetime import datetime
from dotenv import load_dotenv

# 強制 stdout 使用 UTF-8
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))

pyautogui.FAILSAFE = True
SCREENSHOT_DIR = Path.home() / "Desktop"

# ── 天氣 ────────────────────────────────────────────

def get_weather(city: str):
    try:
        res = requests.get(
            f"https://wttr.in/{city}?format=j1",
            headers={"User-Agent": "curl/7.68.0"},
            timeout=10
        )
        data = res.json()
        current = data["current_condition"][0]
        area = data["nearest_area"][0]
        location = area["areaName"][0]["value"] + ", " + area["country"][0]["value"]
        desc = current["lang_zh-tw"][0]["value"] if "lang_zh-tw" in current else current["weatherDesc"][0]["value"]
        temp = current["temp_C"]
        feels = current["FeelsLikeC"]
        humidity = current["humidity"]
        wind = current["windspeedKmph"]
        wind_dir = current["winddir16Point"]
        print(
            f"📍 {location}\n"
            f"🌤 {desc}\n"
            f"🌡 氣溫：{temp}°C（體感 {feels}°C）\n"
            f"💧 濕度：{humidity}%\n"
            f"💨 風速：{wind} km/h（{wind_dir}）"
        )
    except Exception as e:
        print(f"查不到「{city}」的天氣：{e}")


# ── 股票 ────────────────────────────────────────────

def calc_rsi(closes, period=14):
    deltas = closes.diff().dropna()
    gains = deltas.clip(lower=0)
    losses = -deltas.clip(upper=0)
    avg_gain = gains.rolling(period).mean().iloc[-1]
    avg_loss = losses.rolling(period).mean().iloc[-1]
    if avg_loss == 0:
        return 100.0
    rs = avg_gain / avg_loss
    return round(100 - (100 / (1 + rs)), 1)

def get_stock(symbol: str, period: str = "1mo"):
    try:
        import yfinance as yf
        ticker = yf.Ticker(symbol)
        info = ticker.info
        hist = ticker.history(period="3mo")

        if hist.empty:
            print(f"找不到「{symbol}」的股票數據，請確認代號是否正確。")
            return

        name = info.get("longName") or info.get("shortName") or symbol
        currency = info.get("currency", "")
        current = hist["Close"].iloc[-1]
        prev = hist["Close"].iloc[-2] if len(hist) > 1 else current
        change = current - prev
        change_pct = (change / prev) * 100 if prev else 0
        arrow = "▲" if change >= 0 else "▼"
        volume = hist["Volume"].iloc[-1]
        avg_volume = hist["Volume"].tail(20).mean()
        volume_ratio = volume / avg_volume if avg_volume else 1

        ma5 = hist["Close"].tail(5).mean()
        ma20 = hist["Close"].tail(20).mean()
        ma60 = hist["Close"].tail(60).mean() if len(hist) >= 60 else hist["Close"].mean()
        rsi = calc_rsi(hist["Close"]) if len(hist) >= 15 else None

        if ma5 > ma20 > ma60:
            trend = "強勢多頭（MA5>MA20>MA60）📈"
        elif ma5 < ma20 < ma60:
            trend = "強勢空頭（MA5<MA20<MA60）📉"
        elif ma5 > ma20:
            trend = "短線偏多（MA5>MA20）🔼"
        else:
            trend = "短線偏空（MA5<MA20）🔽"

        rsi_note = ""
        if rsi is not None:
            if rsi >= 80:
                rsi_note = "（嚴重超買 ⚠️）"
            elif rsi >= 70:
                rsi_note = "（超買區間）"
            elif rsi <= 20:
                rsi_note = "（嚴重超賣 💡）"
            elif rsi <= 30:
                rsi_note = "（超賣區間）"
            else:
                rsi_note = "（中性）"

        market_cap = info.get("marketCap")
        pe_ratio = info.get("trailingPE")
        week52_high = info.get("fiftyTwoWeekHigh")
        week52_low = info.get("fiftyTwoWeekLow")
        high_period = hist["High"].tail(20).max()
        low_period = hist["Low"].tail(20).min()
        pct_from_high = ((current - high_period) / high_period * 100) if high_period else 0
        pct_from_low = ((current - low_period) / low_period * 100) if low_period else 0

        result = (
            f"📊 {name} ({symbol})\n"
            f"💰 現價：{current:.2f} {currency}  {arrow} {abs(change):.2f} ({change_pct:+.2f}%)\n"
            f"📦 成交量：{volume:,}（均量 {volume_ratio:.1f}x）\n"
            f"\n── 技術指標 ──\n"
            f"MA5：{ma5:.2f}　MA20：{ma20:.2f}　MA60：{ma60:.2f}\n"
            f"趨勢：{trend}\n"
        )
        if rsi is not None:
            result += f"RSI(14)：{rsi}{rsi_note}\n"
        result += (
            f"近20日高點：{high_period:.2f}（距高 {pct_from_high:.1f}%）\n"
            f"近20日低點：{low_period:.2f}（距低 +{pct_from_low:.1f}%）\n"
        )
        if week52_high and week52_low:
            result += f"52週高低：{week52_low:.2f} ~ {week52_high:.2f}\n"
        result += "\n── 基本面 ──\n"
        if market_cap:
            mc_str = f"{market_cap/1e12:.2f}兆" if market_cap >= 1e12 else f"{market_cap/1e8:.0f}億"
            result += f"市值：{mc_str} {currency}\n"
        if pe_ratio:
            result += f"本益比：{pe_ratio:.1f}\n"

        print(result.strip())

    except Exception as e:
        print(f"查詢「{symbol}」失敗：{e}")


# ── 圖片生成 ────────────────────────────────────────

def add_text_overlay(image_bytes: bytes, text: str) -> bytes:
    try:
        from PIL import Image, ImageDraw, ImageFont
        img = Image.open(io.BytesIO(image_bytes)).convert("RGBA")
        w, h = img.size
        overlay = Image.new("RGBA", img.size, (0, 0, 0, 0))
        draw = ImageDraw.Draw(overlay)
        font_size = max(36, w // 12)
        font_path = "C:/Windows/Fonts/msjhbd.ttc"
        try:
            font = ImageFont.truetype(font_path, font_size)
        except Exception:
            font = ImageFont.load_default()
        bbox = draw.textbbox((0, 0), text, font=font)
        tw, th = bbox[2] - bbox[0], bbox[3] - bbox[1]
        x = (w - tw) // 2
        y = h - th - int(h * 0.08)
        for dx, dy in [(-2,-2),(2,-2),(-2,2),(2,2)]:
            draw.text((x+dx, y+dy), text, font=font, fill=(0, 0, 0, 180))
        draw.text((x, y), text, font=font, fill=(255, 255, 255, 255))
        result = Image.alpha_composite(img, overlay).convert("RGB")
        buf = io.BytesIO()
        result.save(buf, format="JPEG", quality=92)
        return buf.getvalue()
    except Exception:
        return image_bytes

def generate_image(prompt: str, overlay_text: str = ""):
    hf_token = os.getenv("HF_TOKEN")
    if not hf_token:
        print("未設定 HF_TOKEN，請在 .env 加入 HF_TOKEN=your_token")
        return
    headers = {"Authorization": f"Bearer {hf_token}"}
    payload = {"inputs": prompt}
    models = [
        "black-forest-labs/FLUX.1-schnell",
        "stabilityai/stable-diffusion-xl-base-1.0",
    ]
    image_bytes = None
    for model in models:
        for attempt in range(2):
            try:
                print(f"嘗試模型：{model}...")
                res = requests.post(
                    f"https://router.huggingface.co/hf-inference/models/{model}",
                    headers=headers,
                    json=payload,
                    timeout=120
                )
                if res.status_code == 200 and res.headers.get("content-type", "").startswith("image"):
                    image_bytes = res.content
                    break
                if res.status_code == 503:
                    time.sleep(10)
                    continue
            except Exception:
                pass
        if image_bytes:
            break

    if not image_bytes:
        print("圖片生成失敗")
        return

    if overlay_text:
        image_bytes = add_text_overlay(image_bytes, overlay_text)

    filename = SCREENSHOT_DIR / f"generated_{datetime.now().strftime('%Y%m%d_%H%M%S')}.jpg"
    with open(filename, "wb") as f:
        f.write(image_bytes)
    print(f"圖片已儲存：{filename}")


# ── 桌面控制 ────────────────────────────────────────

def screenshot():
    img = pyautogui.screenshot()
    filename = SCREENSHOT_DIR / f"screenshot_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
    img.save(filename)
    print(f"截圖已儲存：{filename}")
    return str(filename)

def click(x, y):
    pyautogui.click(int(x), int(y))
    print(f"已點擊 ({x}, {y})")

def double_click(x, y):
    pyautogui.doubleClick(int(x), int(y))
    print(f"已雙擊 ({x}, {y})")

def right_click(x, y):
    pyautogui.rightClick(int(x), int(y))
    print(f"已右鍵點擊 ({x}, {y})")

def move(x, y):
    pyautogui.moveTo(int(x), int(y), duration=0.3)
    print(f"滑鼠已移到 ({x}, {y})")

def type_text(text):
    pyautogui.write(text, interval=0.05)
    print(f"已輸入：{text}")

def press_key(key):
    pyautogui.press(key)
    print(f"已按下：{key}")

def open_app(app):
    subprocess.Popen(app, shell=True)
    print(f"已開啟：{app}")

def scroll(direction, amount=3):
    amount = int(amount)
    pyautogui.scroll(amount if direction == "up" else -amount)
    print(f"已向{direction}滾動 {amount} 格")

def pos():
    x, y = pyautogui.position()
    print(f"目前滑鼠位置：({x}, {y})")


# ── 長期記憶 ────────────────────────────────────────

MEMORY_DB = Path("C:/Users/blue_/claude-telegram-bot/memory.db")

def memory_save(chat_id: int, content: str):
    conn = sqlite3.connect(MEMORY_DB)
    conn.execute("CREATE TABLE IF NOT EXISTS long_term_memory (id INTEGER PRIMARY KEY AUTOINCREMENT, chat_id INTEGER NOT NULL, content TEXT NOT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)")
    conn.execute("INSERT INTO long_term_memory (chat_id, content) VALUES (?, ?)", (chat_id, content))
    conn.commit()
    conn.close()
    print(f"✅ 已儲存記憶：{content}")

def memory_list(chat_id: int):
    conn = sqlite3.connect(MEMORY_DB)
    conn.execute("CREATE TABLE IF NOT EXISTS long_term_memory (id INTEGER PRIMARY KEY AUTOINCREMENT, chat_id INTEGER NOT NULL, content TEXT NOT NULL, created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP)")
    rows = conn.execute("SELECT id, content, created_at FROM long_term_memory WHERE chat_id=? ORDER BY id DESC", (chat_id,)).fetchall()
    conn.close()
    if not rows:
        print("目前沒有長期記憶。")
        return
    print(f"📝 長期記憶（chat_id={chat_id}）：")
    for r in rows:
        print(f"  [{r[0]}] {r[1]}  ({r[2]})")

def memory_del(chat_id: int, memory_id: int):
    conn = sqlite3.connect(MEMORY_DB)
    conn.execute("DELETE FROM long_term_memory WHERE id=? AND chat_id=?", (memory_id, chat_id))
    conn.commit()
    conn.close()
    print(f"✅ 已刪除記憶 ID {memory_id}")


# ── Vision 截圖理解 ──────────────────────────────────

def vision(question: str = "請描述這個畫面上有什麼，以及目前電腦在做什麼事。"):
    from anthropic import Anthropic
    import base64
    client = Anthropic()
    img = pyautogui.screenshot()
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    img_b64 = base64.standard_b64encode(buf.getvalue()).decode("utf-8")
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        messages=[{
            "role": "user",
            "content": [
                {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": img_b64}},
                {"type": "text", "text": question}
            ]
        }]
    )
    print(response.content[0].text)


# ── 圖像定位 ─────────────────────────────────────────

def find_image(template_path: str, confidence: float = 0.8):
    try:
        location = pyautogui.locateOnScreen(template_path, confidence=confidence)
        if location:
            cx, cy = pyautogui.center(location)
            print(f"✅ 找到圖片：中心座標 ({cx}, {cy})，區域 {location}")
            return cx, cy
        else:
            print("❌ 畫面上找不到該圖片")
            return None
    except Exception as e:
        print(f"❌ 搜尋失敗：{e}")
        return None


# ── 瀏覽器自動化 ──────────────────────────────────────

_browser_context = {}

def browser(action: str, *args):
    from playwright.sync_api import sync_playwright

    if action == "open":
        url = args[0] if args else "https://www.google.com"
        def _open():
            pw = sync_playwright().start()
            b = pw.chromium.launch(headless=False)
            page = b.new_page()
            page.goto(url)
            _browser_context["pw"] = pw
            _browser_context["browser"] = b
            _browser_context["page"] = page
            print(f"✅ 已開啟：{url}")
        _open()

    elif action == "click":
        page = _browser_context.get("page")
        if not page:
            print("❌ 瀏覽器未開啟，請先執行 browser open <url>")
            return
        selector = args[0]
        page.click(selector)
        print(f"✅ 已點擊：{selector}")

    elif action == "type":
        page = _browser_context.get("page")
        if not page:
            print("❌ 瀏覽器未開啟")
            return
        selector, text = args[0], " ".join(args[1:])
        page.fill(selector, text)
        print(f"✅ 已輸入到 {selector}：{text}")

    elif action == "get_text":
        page = _browser_context.get("page")
        if not page:
            print("❌ 瀏覽器未開啟")
            return
        selector = args[0] if args else "body"
        text = page.inner_text(selector)
        print(text[:2000])

    elif action == "screenshot":
        page = _browser_context.get("page")
        if not page:
            print("❌ 瀏覽器未開啟")
            return
        filename = SCREENSHOT_DIR / f"browser_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png"
        page.screenshot(path=str(filename))
        print(f"✅ 截圖已儲存：{filename}")

    elif action == "goto":
        page = _browser_context.get("page")
        if not page:
            print("❌ 瀏覽器未開啟")
            return
        page.goto(args[0])
        print(f"✅ 已前往：{args[0]}")

    elif action == "close":
        if "browser" in _browser_context:
            _browser_context["browser"].close()
            _browser_context["pw"].stop()
            _browser_context.clear()
            print("✅ 瀏覽器已關閉")
    else:
        print(f"❌ 未知動作：{action}。可用：open / click / type / get_text / screenshot / goto / close")


# ── 視窗管理 ─────────────────────────────────────────

def window_list():
    import pygetwindow as gw
    wins = [w for w in gw.getAllWindows() if w.title.strip()]
    for w in wins:
        print(f"  [{w._hWnd}] {w.title}")

def _find_window(keyword):
    import pygetwindow as gw
    wins = [w for w in gw.getAllWindows() if keyword.lower() in w.title.lower()]
    return wins[0] if wins else None

def window_focus(keyword):
    w = _find_window(keyword)
    if w:
        w.activate()
        print(f"✅ 已切換到：{w.title}")
    else:
        print(f"❌ 找不到視窗：{keyword}")

def window_close(keyword):
    w = _find_window(keyword)
    if w:
        w.close()
        print(f"✅ 已關閉：{w.title}")
    else:
        print(f"❌ 找不到視窗：{keyword}")

def window_min(keyword):
    w = _find_window(keyword)
    if w:
        w.minimize()
        print(f"✅ 已最小化：{w.title}")
    else:
        print(f"❌ 找不到視窗：{keyword}")

def window_max(keyword):
    w = _find_window(keyword)
    if w:
        w.maximize()
        print(f"✅ 已最大化：{w.title}")
    else:
        print(f"❌ 找不到視窗：{keyword}")


# ── 組合鍵 / 剪貼簿 ──────────────────────────────────

def hotkey(*keys):
    combo = "+".join(keys)
    pyautogui.hotkey(*keys)
    print(f"✅ 已執行組合鍵：{combo}")

def clipboard_get():
    import pyperclip
    text = pyperclip.paste()
    print(text)

def clipboard_set(text):
    import pyperclip
    pyperclip.copy(text)
    print(f"✅ 已寫入剪貼簿：{text}")


# ── 檔案系統 ─────────────────────────────────────────

def file_list(path="."):
    p = Path(path)
    if not p.exists():
        print(f"❌ 路徑不存在：{path}")
        return
    for item in sorted(p.iterdir()):
        tag = "📁" if item.is_dir() else "📄"
        print(f"  {tag} {item.name}")

def file_read(path):
    p = Path(path)
    if not p.exists():
        print(f"❌ 檔案不存在：{path}")
        return
    print(p.read_text(encoding="utf-8", errors="replace"))

def file_write(path, content):
    Path(path).write_text(content, encoding="utf-8")
    print(f"✅ 已寫入：{path}")

def file_delete(path):
    import shutil
    p = Path(path)
    if not p.exists():
        print(f"❌ 不存在：{path}")
        return
    if p.is_dir():
        shutil.rmtree(p)
    else:
        p.unlink()
    print(f"✅ 已刪除：{path}")

def file_copy(src, dst):
    import shutil
    shutil.copy2(src, dst)
    print(f"✅ 已複製：{src} → {dst}")

def file_move(src, dst):
    import shutil
    shutil.move(src, dst)
    print(f"✅ 已移動：{src} → {dst}")

def file_search(folder, keyword):
    results = list(Path(folder).rglob(f"*{keyword}*"))
    if not results:
        print(f"找不到包含「{keyword}」的檔案")
        return
    for r in results:
        print(f"  {r}")


# ── 系統監控 ─────────────────────────────────────────

def sysinfo():
    import psutil
    cpu = psutil.cpu_percent(interval=1)
    mem = psutil.virtual_memory()
    disk = psutil.disk_usage("C:/")
    print(
        f"💻 系統狀態\n"
        f"CPU：{cpu}%\n"
        f"記憶體：{mem.percent}%（已用 {mem.used//1024//1024}MB / 共 {mem.total//1024//1024}MB）\n"
        f"磁碟 C：{disk.percent}%（已用 {disk.used//1024//1024//1024}GB / 共 {disk.total//1024//1024//1024}GB）"
    )

def process_list():
    import psutil
    print(f"{'PID':<8} {'CPU%':<7} {'記憶體MB':<10} 程序名稱")
    print("-" * 45)
    procs = sorted(psutil.process_iter(["pid","name","cpu_percent","memory_info"]),
                   key=lambda p: p.info["memory_info"].rss if p.info["memory_info"] else 0, reverse=True)
    for p in procs[:30]:
        mem_mb = p.info["memory_info"].rss // 1024 // 1024 if p.info["memory_info"] else 0
        print(f"{p.info['pid']:<8} {p.info['cpu_percent']:<7} {mem_mb:<10} {p.info['name']}")

def process_kill(name_or_pid):
    import psutil
    try:
        pid = int(name_or_pid)
        psutil.Process(pid).kill()
        print(f"✅ 已結束 PID {pid}")
    except ValueError:
        killed = 0
        for p in psutil.process_iter(["pid","name"]):
            if name_or_pid.lower() in p.info["name"].lower():
                p.kill()
                killed += 1
        print(f"✅ 已結束 {killed} 個「{name_or_pid}」程序")


# ── Windows 通知 ──────────────────────────────────────

def notify(title, message):
    try:
        from win10toast import ToastNotifier
        ToastNotifier().show_toast(title, message, duration=5, threaded=True)
        print(f"✅ 通知已送出：{title} - {message}")
    except Exception as e:
        print(f"❌ 通知失敗：{e}")


# ── 語音合成 TTS ──────────────────────────────────────

def tts(text):
    import pyttsx3
    engine = pyttsx3.init()
    engine.setProperty("rate", 180)
    engine.say(text)
    engine.runAndWait()
    print(f"✅ 已朗讀：{text}")


# ── 動作錄製/重播 ─────────────────────────────────────

_record_events = []
_record_listener = None
RECORD_DIR = Path("C:/Users/blue_/recordings")
RECORD_DIR.mkdir(exist_ok=True)

def record_start():
    from pynput import mouse, keyboard
    global _record_events, _record_listener
    _record_events = []
    start_time = time.time()

    def on_move(x, y):
        _record_events.append({"t": time.time()-start_time, "type": "move", "x": x, "y": y})
    def on_click(x, y, button, pressed):
        _record_events.append({"t": time.time()-start_time, "type": "click", "x": x, "y": y, "button": str(button), "pressed": pressed})
    def on_key(key):
        try:
            _record_events.append({"t": time.time()-start_time, "type": "key", "key": key.char})
        except AttributeError:
            _record_events.append({"t": time.time()-start_time, "type": "key", "key": str(key)})

    ml = mouse.Listener(on_move=on_move, on_click=on_click)
    kl = keyboard.Listener(on_press=on_key)
    ml.start(); kl.start()
    _record_listener = (ml, kl)
    print("✅ 開始錄製，執行 record_stop 停止")

def record_stop():
    global _record_listener
    if _record_listener:
        for l in _record_listener:
            l.stop()
        _record_listener = None
    filename = RECORD_DIR / f"rec_{datetime.now().strftime('%Y%m%d_%H%M%S')}.json"
    import json
    filename.write_text(json.dumps(_record_events, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅ 錄製完成，已儲存：{filename}（共 {len(_record_events)} 個事件）")

def record_play(filepath):
    import json
    events = json.loads(Path(filepath).read_text(encoding="utf-8"))
    prev_t = 0
    for e in events:
        delay = e["t"] - prev_t
        if delay > 0:
            time.sleep(min(delay, 2))
        prev_t = e["t"]
        if e["type"] == "move":
            pyautogui.moveTo(e["x"], e["y"])
        elif e["type"] == "click" and e["pressed"]:
            pyautogui.click(e["x"], e["y"])
        elif e["type"] == "key":
            try:
                pyautogui.press(e["key"])
            except Exception:
                pass
    print(f"✅ 重播完成")


# ── Email ─────────────────────────────────────────────

def send_email(to: str, subject: str, body: str):
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    smtp_host = os.getenv("SMTP_HOST", "smtp.gmail.com")
    smtp_port = int(os.getenv("SMTP_PORT", "587"))
    smtp_user = os.getenv("SMTP_USER", "")
    smtp_pass = os.getenv("SMTP_PASS", "")
    if not smtp_user or not smtp_pass:
        print("❌ 請在 .env 設定 SMTP_USER 和 SMTP_PASS")
        return
    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = to
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    with smtplib.SMTP(smtp_host, smtp_port) as s:
        s.starttls()
        s.login(smtp_user, smtp_pass)
        s.send_message(msg)
    print(f"✅ Email 已寄出到：{to}")


# ── 排程管理 ────────────────────────────────────────

SCHTASKS = "C:\\Windows\\System32\\schtasks.exe"
BOT_SCRIPT = r"C:\Users\blue_\claude-telegram-bot\bot.py"

def bot_status():
    result = subprocess.run(
        ["powershell.exe", "-Command",
         "Get-Process pythonw -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count"],
        capture_output=True, text=True
    )
    count = int(result.stdout.strip()) if result.stdout.strip().isdigit() else 0
    if count > 0:
        print(f"✅ Bot 執行中（{count} 個 pythonw 程序）")
    else:
        print("❌ Bot 未執行")

def bot_restart():
    # 停掉所有 pythonw
    subprocess.run(["powershell.exe", "-Command", "Stop-Process -Name pythonw -Force -ErrorAction SilentlyContinue"])
    import time
    time.sleep(1)
    subprocess.Popen(["pythonw", BOT_SCRIPT], cwd=str(Path(BOT_SCRIPT).parent))
    print("✅ Bot 已重啟")

def schedule_list():
    result = subprocess.run(
        [SCHTASKS, "/Query", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, encoding="cp950", errors="replace"
    )
    lines = [l for l in result.stdout.strip().splitlines() if l.strip()]
    print(f"{'任務名稱':<35} {'下次執行':<25} {'狀態'}")
    print("-" * 75)
    for line in lines:
        parts = line.strip('"').split('","')
        if len(parts) >= 3:
            name = parts[0].replace("\\", "").strip()
            next_run = parts[1].strip()
            status = parts[2].strip()
            print(f"{name:<35} {next_run:<25} {status}")

def schedule_add(name: str, time_hhmm: str, script_path: str):
    # 建立排程
    subprocess.run([SCHTASKS, "/Create", "/TN", name,
                    "/TR", f"pythonw {script_path}",
                    "/SC", "DAILY", "/ST", time_hhmm, "/F"],
                   capture_output=True)
    # 設定喚醒
    ps = (
        f"$t = Get-ScheduledTask -TaskName '{name}';"
        f"$t.Settings.WakeToRun = $true;"
        f"$t.Settings.DisallowStartIfOnBatteries = $false;"
        f"$t.Settings.StopIfGoingOnBatteries = $false;"
        f"Set-ScheduledTask -TaskName '{name}' -Settings $t.Settings | Out-Null;"
        f"Write-Host '已建立'"
    )
    subprocess.run(["powershell.exe", "-Command", ps], capture_output=True, text=True)
    print(f"✅ 排程 [{name}] 已建立，每天 {time_hhmm} 執行，可喚醒電腦")

def schedule_del(name: str):
    result = subprocess.run([SCHTASKS, "/Delete", "/TN", name, "/F"],
                            capture_output=True, text=True, encoding="cp950", errors="replace")
    if result.returncode == 0:
        print(f"✅ 排程 [{name}] 已刪除")
    else:
        print(f"❌ 刪除失敗：{result.stderr.strip()}")


# ── 主程式 ──────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(0)

    tool = sys.argv[1]
    args = sys.argv[2:]

    tools = {
        "weather":       lambda: get_weather(args[0]),
        "stock":         lambda: get_stock(args[0], args[1] if len(args) > 1 else "1mo"),
        "image":         lambda: generate_image(args[0], args[1] if len(args) > 1 else ""),
        "screenshot":    lambda: screenshot(),
        "click":         lambda: click(*args),
        "double_click":  lambda: double_click(*args),
        "right_click":   lambda: right_click(*args),
        "move":          lambda: move(*args),
        "type":          lambda: type_text(" ".join(args)),
        "press":         lambda: press_key(args[0]),
        "open":          lambda: open_app(" ".join(args)),
        "scroll":        lambda: scroll(*args),
        "pos":           lambda: pos(),
        "bot_status":    lambda: bot_status(),
        "bot_restart":   lambda: bot_restart(),
        "schedule_list": lambda: schedule_list(),
        "schedule_add":  lambda: schedule_add(args[0], args[1], args[2]),
        "schedule_del":  lambda: schedule_del(args[0]),
        "memory_save":   lambda: memory_save(int(args[0]), " ".join(args[1:])),
        "memory_list":   lambda: memory_list(int(args[0])),
        "memory_del":    lambda: memory_del(int(args[0]), int(args[1])),
        "vision":        lambda: vision(" ".join(args) if args else "請描述這個畫面上有什麼，以及目前電腦在做什麼事。"),
        "find_image":    lambda: find_image(args[0], float(args[1]) if len(args) > 1 else 0.8),
        "browser":       lambda: browser(args[0], *args[1:]),
        "window_list":   lambda: window_list(),
        "window_focus":  lambda: window_focus(" ".join(args)),
        "window_close":  lambda: window_close(" ".join(args)),
        "window_min":    lambda: window_min(" ".join(args)),
        "window_max":    lambda: window_max(" ".join(args)),
        "hotkey":        lambda: hotkey(*args),
        "clipboard_get": lambda: clipboard_get(),
        "clipboard_set": lambda: clipboard_set(" ".join(args)),
        "file_list":     lambda: file_list(args[0] if args else "."),
        "file_read":     lambda: file_read(args[0]),
        "file_write":    lambda: file_write(args[0], " ".join(args[1:])),
        "file_delete":   lambda: file_delete(args[0]),
        "file_copy":     lambda: file_copy(args[0], args[1]),
        "file_move":     lambda: file_move(args[0], args[1]),
        "file_search":   lambda: file_search(args[0], args[1]),
        "sysinfo":       lambda: sysinfo(),
        "process_list":  lambda: process_list(),
        "process_kill":  lambda: process_kill(args[0]),
        "notify":        lambda: notify(args[0], " ".join(args[1:])),
        "tts":           lambda: tts(" ".join(args)),
        "record_start":  lambda: record_start(),
        "record_stop":   lambda: record_stop(),
        "record_play":   lambda: record_play(args[0]),
        "email":         lambda: send_email(args[0], args[1], " ".join(args[2:])),
    }

    if tool not in tools:
        print(f"未知工具：{tool}")
        print(__doc__)
        sys.exit(1)

    try:
        tools[tool]()
    except Exception as e:
        print(f"執行失敗：{e}")
        sys.exit(1)
