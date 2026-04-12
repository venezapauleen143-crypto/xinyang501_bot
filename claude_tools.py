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
  stt                           語音辨識（麥克風錄音轉文字）
  ocr [圖片路徑]                OCR 截圖或指定圖片，辨識文字
  workflow_run <json檔>         執行自動化工作流程
  workflow_save <名稱> <json>   儲存工作流程
  screen_watch <圖片> <指令>    監控螢幕出現指定圖片時執行指令
  monitors                      列出所有螢幕資訊
  zip <來源> <目標zip>          壓縮檔案或資料夾
  unzip <zip檔> <目標資料夾>    解壓縮
  download <網址> [儲存路徑]    下載檔案
  print_file <檔案路徑>         列印文件
  wifi_list                     列出可用 WiFi 網路
  wifi_connect <SSID> <密碼>    連線 WiFi
  screen_stream <秒數>          截圖串流（每秒截圖存桌面，持續N秒）
  wake_listen <關鍵字>          等待語音說出關鍵字後執行
  drag <x1> <y1> <x2> <y2>     拖曳從(x1,y1)到(x2,y2)
  right_menu <x> <y> <項目>     右鍵選單並選擇項目
  ai_plan <目標>                AI 自動規劃並執行多步驟任務
  clipboard_history             顯示剪貼簿歷史紀錄
  vdesktop <left|right|new>     切換虛擬桌面
  power <sleep|restart|shutdown> 電源管理
  bt_scan                       掃描藍牙裝置
  bt_connect <MAC位址>          連線藍牙裝置
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


# ── 語音辨識 STT ─────────────────────────────────────

def stt(duration=5):
    import sounddevice as sd
    import soundfile as sf
    import speech_recognition as sr
    import tempfile
    print(f"🎤 錄音 {duration} 秒，請說話...")
    sample_rate = 16000
    recording = sd.rec(int(duration * sample_rate), samplerate=sample_rate, channels=1, dtype="int16")
    sd.wait()
    with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
        tmp_path = f.name
    sf.write(tmp_path, recording, sample_rate)
    r = sr.Recognizer()
    with sr.AudioFile(tmp_path) as source:
        audio = r.record(source)
    Path(tmp_path).unlink(missing_ok=True)
    try:
        text = r.recognize_google(audio, language="zh-TW")
        print(f"✅ 辨識結果：{text}")
        return text
    except sr.UnknownValueError:
        print("❌ 無法辨識語音")
    except sr.RequestError as e:
        print(f"❌ STT 服務錯誤：{e}")


# ── OCR 文字辨識 ──────────────────────────────────────

def ocr(image_path: str = ""):
    import easyocr
    reader = easyocr.Reader(["ch_tra", "en"], gpu=False)
    if image_path:
        source = image_path
    else:
        img = pyautogui.screenshot()
        source = str(SCREENSHOT_DIR / f"ocr_tmp_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
        img.save(source)
    results = reader.readtext(source)
    texts = [r[1] for r in results]
    output = "\n".join(texts)
    print(output)
    return output


# ── 自動化工作流程 ────────────────────────────────────

WORKFLOW_DIR = Path("C:/Users/blue_/workflows")
WORKFLOW_DIR.mkdir(exist_ok=True)

def workflow_save(name: str, json_content: str):
    import json
    path = WORKFLOW_DIR / f"{name}.json"
    data = json.loads(json_content)
    path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"✅ 工作流程已儲存：{path}")

def workflow_run(json_path: str):
    import json
    path = Path(json_path)
    if not path.exists():
        path = WORKFLOW_DIR / f"{json_path}.json"
    steps = json.loads(path.read_text(encoding="utf-8"))
    print(f"▶ 執行工作流程：{path.name}（共 {len(steps)} 步）")
    for i, step in enumerate(steps, 1):
        tool = step.get("tool")
        args = step.get("args", [])
        delay = step.get("delay", 0)
        print(f"  [{i}] {tool} {args}")
        if delay:
            time.sleep(delay)
        try:
            # 動態呼叫已有工具
            tool_map = {
                "click": lambda a: pyautogui.click(int(a[0]), int(a[1])),
                "type": lambda a: pyautogui.write(" ".join(a), interval=0.05),
                "press": lambda a: pyautogui.press(a[0]),
                "hotkey": lambda a: pyautogui.hotkey(*a),
                "open": lambda a: subprocess.Popen(" ".join(a), shell=True),
                "screenshot": lambda a: screenshot(),
                "move": lambda a: pyautogui.moveTo(int(a[0]), int(a[1]), duration=0.3),
                "scroll": lambda a: pyautogui.scroll(int(a[1]) if a[0]=="up" else -int(a[1])),
                "wait": lambda a: time.sleep(float(a[0])),
                "notify": lambda a: notify(a[0], " ".join(a[1:])),
                "browser": lambda a: browser(a[0], *a[1:]),
                "file_write": lambda a: file_write(a[0], " ".join(a[1:])),
            }
            if tool in tool_map:
                tool_map[tool](args)
            else:
                print(f"    ⚠️ 未知工具：{tool}")
        except Exception as e:
            print(f"    ❌ 步驟失敗：{e}")
    print("✅ 工作流程執行完畢")


# ── 螢幕監控觸發 ──────────────────────────────────────

def screen_watch(template_path: str, command: str, interval: float = 2.0, timeout: float = 60.0):
    print(f"👁 監控中：等待 [{template_path}] 出現，超時 {timeout}s...")
    start = time.time()
    while time.time() - start < timeout:
        try:
            loc = pyautogui.locateOnScreen(template_path, confidence=0.8)
            if loc:
                print(f"✅ 偵測到目標！執行：{command}")
                subprocess.run(command, shell=True)
                return
        except Exception:
            pass
        time.sleep(interval)
    print("⏰ 監控逾時，未偵測到目標")


# ── 多螢幕支援 ───────────────────────────────────────

def monitors():
    try:
        from screeninfo import get_monitors
        for i, m in enumerate(get_monitors()):
            print(f"螢幕 {i}: {m.width}x{m.height} 位置({m.x},{m.y}) {'主螢幕' if m.is_primary else ''}")
    except Exception as e:
        print(f"❌ {e}")


# ── ZIP 壓縮 ─────────────────────────────────────────

def zip_files(source: str, dest: str):
    import zipfile
    src = Path(source)
    with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zf:
        if src.is_dir():
            for f in src.rglob("*"):
                zf.write(f, f.relative_to(src.parent))
        else:
            zf.write(src, src.name)
    print(f"✅ 已壓縮：{source} → {dest}")

def unzip(zip_path: str, dest: str):
    import zipfile
    with zipfile.ZipFile(zip_path, "r") as zf:
        zf.extractall(dest)
    print(f"✅ 已解壓縮：{zip_path} → {dest}")


# ── 網路下載 ─────────────────────────────────────────

def download(url: str, save_path: str = ""):
    if not save_path:
        save_path = str(SCREENSHOT_DIR / url.split("/")[-1].split("?")[0])
    r = requests.get(url, stream=True, timeout=60)
    r.raise_for_status()
    with open(save_path, "wb") as f:
        for chunk in r.iter_content(chunk_size=8192):
            f.write(chunk)
    print(f"✅ 已下載：{save_path}")


# ── 印表機 ───────────────────────────────────────────

def print_file(path: str):
    try:
        import win32api
        win32api.ShellExecute(0, "print", path, None, ".", 0)
        print(f"✅ 已送印：{path}")
    except Exception as e:
        print(f"❌ 列印失敗：{e}")


# ── WiFi 管理 ────────────────────────────────────────

def wifi_list():
    result = subprocess.run(
        ["powershell.exe", "-Command", "netsh wlan show networks mode=bssid"],
        capture_output=True, text=True, encoding="cp950", errors="replace"
    )
    print(result.stdout[:3000])

def wifi_connect(ssid: str, password: str):
    profile = f"""<?xml version="1.0"?>
<WLANProfile xmlns="http://www.microsoft.com/networking/WLAN/profile/v1">
    <name>{ssid}</name>
    <SSIDConfig><SSID><name>{ssid}</name></SSID></SSIDConfig>
    <connectionType>ESS</connectionType>
    <connectionMode>auto</connectionMode>
    <MSM><security><authEncryption>
        <authentication>WPA2PSK</authentication>
        <encryption>AES</encryption>
    </authEncryption>
    <sharedKey><keyType>passPhrase</keyType><protected>false</protected>
        <keyMaterial>{password}</keyMaterial>
    </sharedKey></security></MSM>
</WLANProfile>"""
    profile_path = Path("C:/Users/blue_/wifi_tmp.xml")
    profile_path.write_text(profile, encoding="utf-8")
    subprocess.run(["powershell.exe", "-Command",
                    f"netsh wlan add profile filename='{profile_path}'; netsh wlan connect name='{ssid}'"],
                   capture_output=True)
    profile_path.unlink(missing_ok=True)
    print(f"✅ 已嘗試連線：{ssid}")


# ── 螢幕串流 ─────────────────────────────────────────

def screen_stream(duration: int = 10, interval: float = 1.0):
    print(f"📹 開始串流 {duration} 秒（每 {interval} 秒截圖）...")
    end = time.time() + duration
    count = 0
    while time.time() < end:
        img = pyautogui.screenshot()
        filename = SCREENSHOT_DIR / f"stream_{datetime.now().strftime('%H%M%S')}_{count:03d}.png"
        img.save(filename)
        count += 1
        time.sleep(interval)
    print(f"✅ 串流完成，共 {count} 張截圖存至桌面")


# ── 語音喚醒 ─────────────────────────────────────────

def wake_listen(keyword: str = "小牛馬", duration: int = 5):
    import sounddevice as sd
    import soundfile as sf
    import speech_recognition as sr
    import tempfile
    print(f"👂 監聽中，等待說出「{keyword}」...")
    sample_rate = 16000
    while True:
        recording = sd.rec(int(duration * sample_rate), samplerate=sample_rate, channels=1, dtype="int16")
        sd.wait()
        with tempfile.NamedTemporaryFile(suffix=".wav", delete=False) as f:
            tmp_path = f.name
        sf.write(tmp_path, recording, sample_rate)
        r = sr.Recognizer()
        try:
            with sr.AudioFile(tmp_path) as source:
                audio = r.record(source)
            text = r.recognize_google(audio, language="zh-TW")
            Path(tmp_path).unlink(missing_ok=True)
            if keyword in text:
                print(f"✅ 偵測到喚醒詞：{text}")
                return text
        except Exception:
            Path(tmp_path).unlink(missing_ok=True)


# ── 拖曳 / 右鍵選單 ──────────────────────────────────

def drag(x1, y1, x2, y2, duration=0.5):
    pyautogui.moveTo(int(x1), int(y1))
    pyautogui.dragTo(int(x2), int(y2), duration=float(duration), button="left")
    print(f"✅ 已拖曳 ({x1},{y1}) → ({x2},{y2})")

def right_menu(x, y, item: str = ""):
    pyautogui.rightClick(int(x), int(y))
    time.sleep(0.3)
    if item:
        import pygetwindow as gw
        pyautogui.write(item, interval=0.05)
        time.sleep(0.2)
        pyautogui.press("enter")
        print(f"✅ 已右鍵點擊並選擇：{item}")
    else:
        print(f"✅ 已右鍵點擊 ({x},{y})")


# ── 多步驟 AI 規劃 ────────────────────────────────────

def ai_plan(goal: str):
    from anthropic import Anthropic
    client = Anthropic()
    print(f"🧠 AI 規劃目標：{goal}")
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system="""你是一個電腦自動化規劃師。用戶給你一個目標，你要把它拆解成可執行的步驟。
每個步驟使用以下工具之一：click/type/press/hotkey/open/screenshot/browser/file_write/wait/notify
以 JSON 陣列格式回傳，例如：
[
  {"tool": "open", "args": ["notepad"], "delay": 1},
  {"tool": "type", "args": ["Hello World"]},
  {"tool": "hotkey", "args": ["ctrl","s"]}
]
只回傳 JSON，不要其他文字。""",
        messages=[{"role": "user", "content": f"目標：{goal}"}]
    )
    import json
    plan_text = response.content[0].text.strip()
    try:
        steps = json.loads(plan_text)
        print(f"📋 規劃完成，共 {len(steps)} 步：")
        for i, s in enumerate(steps, 1):
            print(f"  [{i}] {s.get('tool')} {s.get('args', [])}")
        print("\n▶ 開始執行...")
        tool_map = {
            "click": lambda a: pyautogui.click(int(a[0]), int(a[1])),
            "type": lambda a: pyautogui.write(" ".join(str(x) for x in a), interval=0.05),
            "press": lambda a: pyautogui.press(a[0]),
            "hotkey": lambda a: pyautogui.hotkey(*a),
            "open": lambda a: subprocess.Popen(" ".join(str(x) for x in a), shell=True),
            "screenshot": lambda a: screenshot(),
            "wait": lambda a: time.sleep(float(a[0])),
            "notify": lambda a: notify(a[0], a[1] if len(a) > 1 else ""),
            "move": lambda a: pyautogui.moveTo(int(a[0]), int(a[1]), duration=0.3),
            "scroll": lambda a: pyautogui.scroll(int(a[1]) if a[0] == "up" else -int(a[1])),
            "file_write": lambda a: file_write(a[0], " ".join(str(x) for x in a[1:])),
            "browser": lambda a: browser(a[0], *a[1:]),
        }
        for i, step in enumerate(steps, 1):
            t = step.get("tool")
            a = step.get("args", [])
            d = step.get("delay", 0)
            if d: time.sleep(d)
            if t in tool_map:
                tool_map[t](a)
                print(f"  ✅ 步驟 {i} 完成")
            else:
                print(f"  ⚠️ 未知工具：{t}")
        print("✅ 全部執行完畢")
    except json.JSONDecodeError:
        print(f"規劃結果：\n{plan_text}")


# ── 剪貼簿歷史 ───────────────────────────────────────

_clipboard_history = []

def clipboard_history():
    import pyperclip
    current = pyperclip.paste()
    if current and (not _clipboard_history or _clipboard_history[-1] != current):
        _clipboard_history.append(current)
    if not _clipboard_history:
        print("剪貼簿歷史是空的")
        return
    for i, item in enumerate(_clipboard_history[-20:], 1):
        preview = item[:80].replace("\n", "↵")
        print(f"  [{i}] {preview}")


# ── 虛擬桌面 ─────────────────────────────────────────

def vdesktop(action: str):
    if action == "left":
        pyautogui.hotkey("ctrl", "win", "left")
        print("✅ 切換到左邊虛擬桌面")
    elif action == "right":
        pyautogui.hotkey("ctrl", "win", "right")
        print("✅ 切換到右邊虛擬桌面")
    elif action == "new":
        pyautogui.hotkey("ctrl", "win", "d")
        print("✅ 已建立新虛擬桌面")
    else:
        print(f"❌ 未知動作：{action}（可用：left/right/new）")


# ── 電源管理 ─────────────────────────────────────────

def power(action: str):
    cmds = {
        "sleep":    "rundll32.exe powrprof.dll,SetSuspendState 0,1,0",
        "restart":  "shutdown /r /t 5",
        "shutdown": "shutdown /s /t 5",
    }
    if action not in cmds:
        print(f"❌ 未知動作：{action}（可用：sleep/restart/shutdown）")
        return
    subprocess.run(["powershell.exe", "-Command", cmds[action]])
    print(f"✅ 已執行：{action}")


# ── 藍牙管理 ─────────────────────────────────────────

def bt_scan():
    try:
        import asyncio
        import bleak
        async def _scan():
            devices = await bleak.BleakScanner.discover(timeout=5.0)
            return devices
        devices = asyncio.run(_scan())
        if not devices:
            print("找不到藍牙裝置")
            return
        for d in devices:
            print(f"  {d.address}  {d.name or '(未知)'}")
    except Exception as e:
        print(f"❌ 藍牙掃描失敗：{e}")

def bt_connect(mac: str):
    try:
        import asyncio
        import bleak
        async def _connect():
            async with bleak.BleakClient(mac) as client:
                print(f"✅ 已連線：{mac}（服務數：{len(client.services)}）")
                await asyncio.sleep(3)
        asyncio.run(_connect())
    except Exception as e:
        print(f"❌ 連線失敗：{e}")


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
        "stt":           lambda: stt(),
        "ocr":           lambda: ocr(args[0] if args else ""),
        "workflow_run":  lambda: workflow_run(args[0]),
        "workflow_save": lambda: workflow_save(args[0], " ".join(args[1:])),
        "screen_watch":  lambda: screen_watch(args[0], args[1], float(args[2]) if len(args)>2 else 2.0),
        "monitors":      lambda: monitors(),
        "zip":           lambda: zip_files(args[0], args[1]),
        "unzip":         lambda: unzip(args[0], args[1]),
        "download":      lambda: download(args[0], args[1] if len(args)>1 else ""),
        "print_file":    lambda: print_file(args[0]),
        "wifi_list":       lambda: wifi_list(),
        "wifi_connect":    lambda: wifi_connect(args[0], args[1]),
        "screen_stream":   lambda: screen_stream(int(args[0]) if args else 10),
        "wake_listen":     lambda: wake_listen(args[0] if args else "小牛馬"),
        "drag":            lambda: drag(*args),
        "right_menu":      lambda: right_menu(args[0], args[1], args[2] if len(args)>2 else ""),
        "ai_plan":         lambda: ai_plan(" ".join(args)),
        "clipboard_history": lambda: clipboard_history(),
        "vdesktop":        lambda: vdesktop(args[0]),
        "power":           lambda: power(args[0]),
        "bt_scan":         lambda: bt_scan(),
        "bt_connect":      lambda: bt_connect(args[0]),
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
