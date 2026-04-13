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
  run_python <程式碼>           直接執行 Python 程式碼
  run_shell <指令>              執行 PowerShell 指令
  word_read <路徑>              讀取 Word 文件
  word_write <路徑> <內容>      寫入 Word 文件
  excel_read <路徑> [工作表]    讀取 Excel
  excel_write <路徑> <工作表> <資料JSON> 寫入 Excel
  pdf_read <路徑>               讀取 PDF 文字
  screen_diff <間隔秒> <區域>   偵測螢幕區域變化
  scrape <網址> <CSS選擇器>     爬取網頁指定內容
  img_edit <動作> <路徑> [參數] 圖片編輯（crop/resize/text/merge）
  gdrive_upload <檔案> <資料夾ID> 上傳到 Google Drive
  gdrive_download <檔案ID> <路徑> 從 Google Drive 下載
  db_query <資料庫路徑> <SQL>   執行 SQLite 查詢
  db_mysql <host> <db> <SQL>    執行 MySQL 查詢
  encrypt <檔案路徑> <密碼>     加密檔案
  decrypt <檔案路徑> <密碼>     解密檔案
  clipboard_watch <秒數>        監控剪貼簿變化
  qr_gen <內容> [路徑]          生成 QR Code
  qr_scan [圖片路徑]            掃描 QR Code（截圖或指定圖片）
  screen_record [秒數] [輸出路徑]  螢幕錄影存 mp4
  webcam [輸出路徑]             攝影機拍照
  translate <文字> [目標語言] [來源語言]  翻譯（目標預設 zh-TW）
  chart <類型> <資料JSON> [標題] [輸出路徑]  生成圖表（line/bar/pie）
  pptx_read <路徑>              讀取 PowerPoint
  pptx_create <路徑> <slides_json>  建立 PowerPoint
  api_call <方法> <URL> [headers_json] [body_json]  呼叫 REST API
  watchdog <程序名> <腳本路徑> [秒數]  守護程序，崩潰自動重啟
  ssh_run <host> <user> <password> <指令>  SSH 執行遠端指令
  sftp_upload <host> <user> <pass> <local> <remote>  SFTP 上傳
  sftp_download <host> <user> <pass> <remote> <local>  SFTP 下載
  net_ping <host> [次數]        Ping 網路主機
  net_traceroute <host>         路由追蹤
  net_portscan <host> [ports]   Port 掃描
  win_service <list|start|stop> [服務名]  Windows 服務管理
  pdf_merge <paths_json> <輸出路徑>  合併 PDF
  pdf_split <路徑> <輸出資料夾>  分割 PDF 每頁
  pdf_watermark <路徑> <文字> [輸出路徑]  PDF 加浮水印
  audio_convert <輸入> <輸出>   音訊格式轉換
  audio_trim <輸入> <起始ms> <結束ms> [輸出]  音訊剪輯
  discord_notify <webhook_url> <訊息>  Discord 推播
  line_notify <token> <訊息>    LINE Notify 推播
  disk_clean <list|clean>       磁碟暫存清理
  backup <來源> <目標資料夾>    備份壓縮
  registry_read <key_path> [value_name]  讀取登錄檔
  registry_write <key_path> <value_name> <value>  寫入登錄檔
  video_screenshot <路徑> [秒數] [輸出]  影片截取畫面
  video_trim <路徑> <起始秒> <結束秒> [輸出]  影片剪輯
  monitor_list                  列出所有螢幕資訊
  email_read <host> <user> <pass> [資料夾] [數量]  讀取 IMAP 郵件
  gcal_list [天數]              列出 Google Calendar 行程
  gcal_add <標題> <開始> <結束> [說明]  新增行程（時間格式 2026-04-13T10:00:00）
  global_hotkey <熱鍵> <指令> [秒數]  監聽全域快捷鍵
  git_op <動作> [repo] [參數]   Git 操作（status/log/pull/add/commit/push/diff）
  hw_monitor                    硬體監控（GPU/電池/溫度）
  report_gen <標題> <資料JSON> [輸出路徑]  生成 HTML 報告
  dropbox_upload <local> <remote> [token]  上傳到 Dropbox
  dropbox_download <remote> <local> [token]  從 Dropbox 下載
  docker_op <動作> [容器名]     Docker 操作（list/start/stop/logs/images）
  pdf_to_images <路徑> [輸出資料夾] [dpi]  PDF 轉圖片
  barcode_scan [圖片路徑]       掃描條碼或 QR Code
  nlp_summarize <文字>          AI 文字摘要
  nlp_sentiment <文字>          AI 情緒分析
  vpn_control <list|connect|disconnect> [名稱] [帳號] [密碼]  VPN 控制
  sys_restore <create|list> [說明]  Windows 系統還原點
  disk_analyze [路徑] [數量]    磁碟空間分析
  face_detect [圖片路徑] [輸出路徑]  人臉偵測
  video_to_gif <路徑> [起始秒] [持續秒] [輸出] [fps]  影片轉 GIF
  excel_chart <路徑> <工作表> [類型] [標題]  Excel 生成圖表
  speedtest                     網路速度測試
  screenshot_compare [圖1] [圖2] [輸出]  截圖比對差異
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


# ── 程式碼執行 ───────────────────────────────────────

def run_python(code: str):
    import traceback, contextlib
    buf = io.StringIO()
    try:
        with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
            exec(code, {"__builtins__": __builtins__})
        result = buf.getvalue()
        print(result if result else "✅ 執行完成（無輸出）")
    except Exception:
        print(f"❌ 執行錯誤：\n{traceback.format_exc()}")

def run_shell(cmd: str):
    result = subprocess.run(
        ["powershell.exe", "-Command", cmd],
        capture_output=True, text=True, encoding="cp950", errors="replace"
    )
    output = result.stdout + result.stderr
    print(output.strip() or "✅ 執行完成")


# ── 文件處理 ─────────────────────────────────────────

def word_read(path: str):
    from docx import Document
    doc = Document(path)
    text = "\n".join(p.text for p in doc.paragraphs if p.text.strip())
    print(text[:3000])

def word_write(path: str, content: str):
    from docx import Document
    doc = Document()
    for line in content.split("\n"):
        doc.add_paragraph(line)
    doc.save(path)
    print(f"✅ 已寫入：{path}")

def excel_read(path: str, sheet: str = ""):
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[sheet] if sheet and sheet in wb.sheetnames else wb.active
    for row in ws.iter_rows(values_only=True):
        print("\t".join(str(c) if c is not None else "" for c in row))

def excel_write(path: str, sheet: str, data_json: str):
    import openpyxl, json
    data = json.loads(data_json)
    try:
        wb = openpyxl.load_workbook(path)
    except FileNotFoundError:
        wb = openpyxl.Workbook()
    ws = wb[sheet] if sheet in wb.sheetnames else wb.create_sheet(sheet)
    for row in data:
        ws.append(row)
    wb.save(path)
    print(f"✅ 已寫入 Excel：{path} [{sheet}]")

def pdf_read(path: str):
    import fitz
    doc = fitz.open(path)
    text = "\n".join(page.get_text() for page in doc)
    print(text[:3000])


# ── 螢幕變化偵測 ──────────────────────────────────────

def screen_diff(interval: float = 1.0, duration: float = 30.0, region=None):
    import numpy as np
    import cv2
    print(f"👁 監控螢幕變化（{duration}秒）...")
    prev = None
    end = time.time() + duration
    changes = 0
    while time.time() < end:
        img = pyautogui.screenshot(region=region)
        arr = np.array(img.convert("L"))
        if prev is not None:
            diff = cv2.absdiff(prev, arr)
            score = diff.mean()
            if score > 5:
                changes += 1
                ts = datetime.now().strftime("%H:%M:%S")
                print(f"  [{ts}] 偵測到變化！差異分數：{score:.1f}")
        prev = arr
        time.sleep(interval)
    print(f"✅ 監控結束，共偵測到 {changes} 次變化")


# ── 網頁爬蟲 ─────────────────────────────────────────

def scrape(url: str, selector: str = "body"):
    from bs4 import BeautifulSoup
    res = requests.get(url, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
    res.raise_for_status()
    soup = BeautifulSoup(res.text, "html.parser")
    elements = soup.select(selector)
    if not elements:
        print(f"找不到選擇器：{selector}")
        return
    for el in elements[:10]:
        text = el.get_text(strip=True)
        if text:
            print(text[:200])


# ── 圖片編輯 ─────────────────────────────────────────

def img_edit(action: str, path: str, *args):
    from PIL import Image, ImageDraw, ImageFont
    img = Image.open(path)

    if action == "crop":
        x1, y1, x2, y2 = int(args[0]), int(args[1]), int(args[2]), int(args[3])
        img = img.crop((x1, y1, x2, y2))
    elif action == "resize":
        w, h = int(args[0]), int(args[1])
        img = img.resize((w, h), Image.LANCZOS)
    elif action == "text":
        text = args[0]
        x, y = int(args[1]) if len(args) > 1 else img.width // 2, int(args[2]) if len(args) > 2 else img.height - 50
        draw = ImageDraw.Draw(img)
        try:
            font = ImageFont.truetype("C:/Windows/Fonts/msjhbd.ttc", 36)
        except Exception:
            font = ImageFont.load_default()
        draw.text((x, y), text, font=font, fill=(255, 255, 255))
    elif action == "merge":
        img2 = Image.open(args[0])
        new = Image.new("RGB", (img.width + img2.width, max(img.height, img2.height)))
        new.paste(img, (0, 0))
        new.paste(img2, (img.width, 0))
        img = new
    elif action == "grayscale":
        img = img.convert("L")

    out_path = args[-1] if args and str(args[-1]).endswith((".png",".jpg",".jpeg")) else path
    img.save(out_path)
    print(f"✅ 圖片已儲存：{out_path}")


# ── Google Drive ──────────────────────────────────────

def gdrive_upload(file_path: str, folder_id: str = ""):
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload
        from google.oauth2.credentials import Credentials
        creds_path = Path("C:/Users/blue_/gdrive_token.json")
        if not creds_path.exists():
            print("❌ 找不到 C:/Users/blue_/gdrive_token.json，請先完成 Google 授權")
            return
        creds = Credentials.from_authorized_user_file(str(creds_path))
        service = build("drive", "v3", credentials=creds)
        meta = {"name": Path(file_path).name}
        if folder_id:
            meta["parents"] = [folder_id]
        media = MediaFileUpload(file_path)
        f = service.files().create(body=meta, media_body=media, fields="id").execute()
        print(f"✅ 已上傳，檔案 ID：{f.get('id')}")
    except Exception as e:
        print(f"❌ 上傳失敗：{e}")

def gdrive_download(file_id: str, save_path: str):
    try:
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaIoBaseDownload
        from google.oauth2.credentials import Credentials
        creds = Credentials.from_authorized_user_file("C:/Users/blue_/gdrive_token.json")
        service = build("drive", "v3", credentials=creds)
        req = service.files().get_media(fileId=file_id)
        with open(save_path, "wb") as f:
            downloader = MediaIoBaseDownload(f, req)
            done = False
            while not done:
                _, done = downloader.next_chunk()
        print(f"✅ 已下載：{save_path}")
    except Exception as e:
        print(f"❌ 下載失敗：{e}")


# ── 資料庫操作 ───────────────────────────────────────

def db_query(db_path: str, sql: str):
    conn = sqlite3.connect(db_path)
    try:
        cur = conn.execute(sql)
        if cur.description:
            headers = [d[0] for d in cur.description]
            print("\t".join(headers))
            print("-" * 60)
            for row in cur.fetchall():
                print("\t".join(str(c) for c in row))
        else:
            conn.commit()
            print(f"✅ 執行成功，影響 {cur.rowcount} 列")
    finally:
        conn.close()

def db_mysql(host: str, database: str, sql: str, user: str = "root", password: str = ""):
    import pymysql
    conn = pymysql.connect(host=host, user=user, password=password, database=database, charset="utf8mb4")
    try:
        with conn.cursor() as cur:
            cur.execute(sql)
            if cur.description:
                headers = [d[0] for d in cur.description]
                print("\t".join(headers))
                for row in cur.fetchall():
                    print("\t".join(str(c) for c in row))
            else:
                conn.commit()
                print(f"✅ 執行成功，影響 {cur.rowcount} 列")
    finally:
        conn.close()


# ── 加密/解密 ────────────────────────────────────────

def encrypt(file_path: str, password: str):
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    import base64
    salt = b"xinyang501_salt_"
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=100000)
    key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
    f = Fernet(key)
    data = Path(file_path).read_bytes()
    out_path = file_path + ".enc"
    Path(out_path).write_bytes(f.encrypt(data))
    print(f"✅ 已加密：{out_path}")

def decrypt(file_path: str, password: str):
    from cryptography.fernet import Fernet
    from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
    from cryptography.hazmat.primitives import hashes
    import base64
    salt = b"xinyang501_salt_"
    kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=100000)
    key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
    f = Fernet(key)
    data = Path(file_path).read_bytes()
    out_path = file_path.replace(".enc", ".dec")
    Path(out_path).write_bytes(f.decrypt(data))
    print(f"✅ 已解密：{out_path}")


# ── 剪貼簿監控 ───────────────────────────────────────

def clipboard_watch(duration: float = 30.0):
    import pyperclip
    print(f"📋 監控剪貼簿 {duration} 秒...")
    prev = pyperclip.paste()
    end = time.time() + duration
    changes = []
    while time.time() < end:
        current = pyperclip.paste()
        if current != prev and current:
            ts = datetime.now().strftime("%H:%M:%S")
            preview = current[:80].replace("\n", "↵")
            print(f"  [{ts}] 新內容：{preview}")
            changes.append(current)
            prev = current
        time.sleep(0.5)
    print(f"✅ 監控結束，共偵測到 {len(changes)} 次變化")


# ── QR Code ──────────────────────────────────────────

def qr_gen(content: str, save_path: str = ""):
    import qrcode
    qr = qrcode.make(content)
    if not save_path:
        save_path = str(SCREENSHOT_DIR / f"qr_{datetime.now().strftime('%Y%m%d_%H%M%S')}.png")
    qr.save(save_path)
    print(f"✅ QR Code 已生成：{save_path}")

def qr_scan(image_path: str = ""):
    try:
        from pyzbar.pyzbar import decode
        from PIL import Image
        if not image_path:
            img = pyautogui.screenshot()
        else:
            img = Image.open(image_path)
        results = decode(img)
        if not results:
            print("❌ 未偵測到 QR Code")
            return
        for r in results:
            print(f"✅ 掃描結果：{r.data.decode('utf-8')}")
    except Exception as e:
        print(f"❌ 掃描失敗：{e}")


# ── 螢幕錄影 ────────────────────────────────────────

def screen_record(duration: float = 10.0, output: str = ""):
    try:
        import mss, cv2, numpy as np, time as t
        out_path = output or str(Path.home() / "Desktop" / f"record_{datetime.now().strftime('%H%M%S')}.mp4")
        with mss.mss() as sct:
            mon = sct.monitors[1]
            w, h = mon["width"], mon["height"]
            fps = 10
            writer = cv2.VideoWriter(out_path, cv2.VideoWriter_fourcc(*"mp4v"), fps, (w, h))
            end = t.time() + duration
            while t.time() < end:
                frame = np.array(sct.grab(mon))
                frame = cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR)
                writer.write(frame)
                t.sleep(1 / fps)
            writer.release()
        print(f"✅ 錄影完成：{out_path}")
    except Exception as e:
        print(f"❌ 錄影失敗：{e}")


def webcam_capture(output: str = ""):
    try:
        import cv2
        out_path = output or str(Path.home() / "Desktop" / f"webcam_{datetime.now().strftime('%H%M%S')}.jpg")
        cap = cv2.VideoCapture(0)
        ret, frame = cap.read()
        cap.release()
        if ret:
            cv2.imwrite(out_path, frame)
            print(f"✅ 已拍照：{out_path}")
        else:
            print("❌ 無法存取攝影機")
    except Exception as e:
        print(f"❌ 攝影機失敗：{e}")


# ── 翻譯 ────────────────────────────────────────────

def translate(text: str, target: str = "zh-TW", source: str = "auto"):
    try:
        from deep_translator import GoogleTranslator
        result = GoogleTranslator(source=source, target=target).translate(text)
        print(result)
    except Exception as e:
        print(f"❌ 翻譯失敗：{e}")


# ── 資料視覺化 ───────────────────────────────────────

def chart(chart_type: str, data_json: str, title: str = "", output: str = ""):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt
        import json
        data = json.loads(data_json)
        out_path = output or str(Path.home() / "Desktop" / f"chart_{datetime.now().strftime('%H%M%S')}.png")
        fig, ax = plt.subplots()
        if chart_type == "line":
            for label, values in data.items():
                ax.plot(values, label=label)
            ax.legend()
        elif chart_type == "bar":
            ax.bar(list(data.keys()), list(data.values()))
        elif chart_type == "pie":
            ax.pie(list(data.values()), labels=list(data.keys()), autopct="%1.1f%%")
        if title:
            ax.set_title(title)
        plt.tight_layout()
        plt.savefig(out_path)
        plt.close()
        print(f"✅ 圖表已存：{out_path}")
    except Exception as e:
        print(f"❌ 圖表生成失敗：{e}")


# ── PowerPoint ───────────────────────────────────────

def pptx_read(path: str):
    try:
        from pptx import Presentation
        prs = Presentation(path)
        for i, slide in enumerate(prs.slides, 1):
            texts = [sh.text for sh in slide.shapes if sh.has_text_frame]
            print(f"[投影片 {i}] " + " | ".join(t for t in texts if t.strip()))
    except Exception as e:
        print(f"❌ 讀取失敗：{e}")

def pptx_create(path: str, slides_json: str):
    try:
        from pptx import Presentation
        from pptx.util import Inches, Pt
        import json
        slides = json.loads(slides_json)
        prs = Presentation()
        for s in slides:
            layout = prs.slide_layouts[1]
            slide = prs.slides.add_slide(layout)
            if "title" in s:
                slide.shapes.title.text = s["title"]
            if "body" in s and slide.placeholders[1]:
                slide.placeholders[1].text = s["body"]
        prs.save(path)
        print(f"✅ 已建立簡報：{path}")
    except Exception as e:
        print(f"❌ 建立失敗：{e}")


# ── REST API ─────────────────────────────────────────

def api_call(method: str, url: str, headers_json: str = "{}", body_json: str = "{}"):
    try:
        import json
        headers = json.loads(headers_json)
        body = json.loads(body_json)
        resp = requests.request(method.upper(), url, headers=headers, json=body if body else None, timeout=30)
        try:
            print(json.dumps(resp.json(), ensure_ascii=False, indent=2)[:3000])
        except Exception:
            print(resp.text[:3000])
    except Exception as e:
        print(f"❌ API 呼叫失敗：{e}")


# ── 程序守護 ─────────────────────────────────────────

def watchdog(process_name: str, script: str, duration: float = 60.0):
    try:
        import psutil, time as t
        print(f"🐕 開始監控 [{process_name}]，持續 {duration} 秒...")
        end = t.time() + duration
        restarts = 0
        while t.time() < end:
            running = any(p.name().lower() == process_name.lower() for p in psutil.process_iter())
            if not running:
                subprocess.Popen(["pythonw" if script.endswith(".py") else script, script] if script.endswith(".py") else [script])
                restarts += 1
                print(f"⚠️ [{process_name}] 已重啟（第 {restarts} 次）")
            t.sleep(5)
        print(f"✅ 守護結束，共重啟 {restarts} 次")
    except Exception as e:
        print(f"❌ 守護失敗：{e}")


# ── SSH ──────────────────────────────────────────────

def ssh_run(host: str, user: str, password: str, command: str, port: int = 22):
    try:
        import paramiko
        client = paramiko.SSHClient()
        client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        client.connect(host, port=port, username=user, password=password, timeout=15)
        _, stdout, stderr = client.exec_command(command)
        out = stdout.read().decode(errors="replace")
        err = stderr.read().decode(errors="replace")
        client.close()
        print((out + err).strip() or "（執行完畢，無輸出）")
    except Exception as e:
        print(f"❌ SSH 失敗：{e}")

def sftp_upload(host: str, user: str, password: str, local: str, remote: str, port: int = 22):
    try:
        import paramiko
        t = paramiko.Transport((host, port))
        t.connect(username=user, password=password)
        sftp = paramiko.SFTPClient.from_transport(t)
        sftp.put(local, remote)
        sftp.close(); t.close()
        print(f"✅ 已上傳：{local} → {remote}")
    except Exception as e:
        print(f"❌ SFTP 上傳失敗：{e}")

def sftp_download(host: str, user: str, password: str, remote: str, local: str, port: int = 22):
    try:
        import paramiko
        t = paramiko.Transport((host, port))
        t.connect(username=user, password=password)
        sftp = paramiko.SFTPClient.from_transport(t)
        sftp.get(remote, local)
        sftp.close(); t.close()
        print(f"✅ 已下載：{remote} → {local}")
    except Exception as e:
        print(f"❌ SFTP 下載失敗：{e}")


# ── 網路診斷 ─────────────────────────────────────────

def net_ping(host: str, count: int = 4):
    result = subprocess.run(["ping", "-n", str(count), host], capture_output=True, text=True, encoding="cp950", errors="replace")
    print(result.stdout.strip())

def net_traceroute(host: str):
    result = subprocess.run(["tracert", host], capture_output=True, text=True, encoding="cp950", errors="replace", timeout=60)
    print(result.stdout[:3000])

def net_portscan(host: str, ports: str = "22,80,443,3306,3389,8080"):
    import socket
    results = []
    for p in [int(x) for x in ports.split(",")]:
        try:
            s = socket.socket()
            s.settimeout(1)
            r = s.connect_ex((host, p))
            results.append(f"Port {p}: {'開放 ✅' if r == 0 else '關閉 ❌'}")
            s.close()
        except Exception:
            results.append(f"Port {p}: 錯誤")
    print("\n".join(results))


# ── Windows 服務 ─────────────────────────────────────

def win_service(action: str, name: str = ""):
    try:
        if action == "list":
            r = subprocess.run(["powershell.exe", "-Command",
                "Get-Service | Select-Object Name,Status | Format-Table -AutoSize"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout[:3000])
        elif action in ("start", "stop"):
            cmd = f"{'Start' if action=='start' else 'Stop'}-Service -Name '{name}' -Force"
            r = subprocess.run(["powershell.exe", "-Command", cmd],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout or r.stderr or f"✅ {action} {name}")
    except Exception as e:
        print(f"❌ 服務操作失敗：{e}")


# ── PDF 編輯 ─────────────────────────────────────────

def pdf_merge(paths_json: str, output: str):
    try:
        import fitz, json
        paths = json.loads(paths_json)
        writer = fitz.open()
        for p in paths:
            writer.insert_pdf(fitz.open(p))
        writer.save(output)
        print(f"✅ 已合併 {len(paths)} 個 PDF：{output}")
    except Exception as e:
        print(f"❌ 合併失敗：{e}")

def pdf_split(path: str, output_dir: str):
    try:
        import fitz
        doc = fitz.open(path)
        Path(output_dir).mkdir(parents=True, exist_ok=True)
        for i, page in enumerate(doc):
            out = fitz.open()
            out.insert_pdf(doc, from_page=i, to_page=i)
            out.save(str(Path(output_dir) / f"page_{i+1}.pdf"))
        print(f"✅ 已分割 {len(doc)} 頁到：{output_dir}")
    except Exception as e:
        print(f"❌ 分割失敗：{e}")

def pdf_watermark(path: str, text: str, output: str = ""):
    try:
        import fitz
        doc = fitz.open(path)
        out_path = output or path.replace(".pdf", "_wm.pdf")
        for page in doc:
            page.insert_text((page.rect.width/2 - 50, page.rect.height/2),
                text, fontsize=40, color=(0.8, 0.8, 0.8), rotate=45)
        doc.save(out_path)
        print(f"✅ 已加浮水印：{out_path}")
    except Exception as e:
        print(f"❌ 浮水印失敗：{e}")


# ── 音訊處理 ─────────────────────────────────────────

def audio_convert(input_path: str, output_path: str):
    try:
        from pydub import AudioSegment
        fmt = Path(output_path).suffix.lstrip(".")
        AudioSegment.from_file(input_path).export(output_path, format=fmt)
        print(f"✅ 已轉換：{output_path}")
    except Exception as e:
        print(f"❌ 音訊轉換失敗：{e}")

def audio_trim(input_path: str, start_ms: int, end_ms: int, output_path: str = ""):
    try:
        from pydub import AudioSegment
        audio = AudioSegment.from_file(input_path)[start_ms:end_ms]
        out = output_path or input_path.replace(".", "_trim.")
        audio.export(out, format=Path(out).suffix.lstrip("."))
        print(f"✅ 已剪輯：{out}")
    except Exception as e:
        print(f"❌ 音訊剪輯失敗：{e}")


# ── 推播通知 ─────────────────────────────────────────

def discord_notify(webhook_url: str, message: str):
    try:
        resp = requests.post(webhook_url, json={"content": message}, timeout=10)
        print(f"✅ Discord 已發送（{resp.status_code}）")
    except Exception as e:
        print(f"❌ Discord 發送失敗：{e}")

def line_notify(token: str, message: str):
    try:
        resp = requests.post(
            "https://notify-api.line.me/api/notify",
            headers={"Authorization": f"Bearer {token}"},
            data={"message": message}, timeout=10
        )
        print(f"✅ LINE 已發送（{resp.status_code}）")
    except Exception as e:
        print(f"❌ LINE 發送失敗：{e}")


# ── 磁碟清理 ─────────────────────────────────────────

def disk_clean(action: str = "list"):
    try:
        import tempfile, shutil
        tmp = Path(tempfile.gettempdir())
        if action == "list":
            files = list(tmp.rglob("*"))
            total = sum(f.stat().st_size for f in files if f.is_file())
            print(f"暫存資料夾：{tmp}\n檔案數：{len(files)}\n佔用空間：{total/1024/1024:.1f} MB")
        elif action == "clean":
            count = 0
            for f in tmp.iterdir():
                try:
                    if f.is_file():
                        f.unlink(); count += 1
                    elif f.is_dir():
                        shutil.rmtree(f, ignore_errors=True); count += 1
                except Exception:
                    pass
            print(f"✅ 已清理 {count} 個暫存項目")
    except Exception as e:
        print(f"❌ 磁碟清理失敗：{e}")

def backup(src: str, dest: str):
    try:
        import shutil
        out = Path(dest) / f"{Path(src).name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        shutil.make_archive(str(out).replace(".zip",""), "zip", src)
        print(f"✅ 備份完成：{out}")
    except Exception as e:
        print(f"❌ 備份失敗：{e}")


# ── Windows 登錄檔 ────────────────────────────────────

def registry_read(key_path: str, value_name: str = ""):
    try:
        import winreg
        parts = key_path.split("\\", 1)
        roots = {"HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
                 "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
                 "HKLM": winreg.HKEY_LOCAL_MACHINE, "HKCU": winreg.HKEY_CURRENT_USER}
        root = roots[parts[0]]
        with winreg.OpenKey(root, parts[1]) as k:
            if value_name:
                val, _ = winreg.QueryValueEx(k, value_name)
                print(f"{value_name} = {val}")
            else:
                i = 0
                while True:
                    try:
                        n, v, _ = winreg.EnumValue(k, i)
                        print(f"{n} = {v}")
                        i += 1
                    except OSError:
                        break
    except Exception as e:
        print(f"❌ 讀取登錄檔失敗：{e}")

def registry_write(key_path: str, value_name: str, value: str):
    try:
        import winreg
        parts = key_path.split("\\", 1)
        roots = {"HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
                 "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
                 "HKLM": winreg.HKEY_LOCAL_MACHINE, "HKCU": winreg.HKEY_CURRENT_USER}
        root = roots[parts[0]]
        with winreg.OpenKey(root, parts[1], 0, winreg.KEY_SET_VALUE) as k:
            winreg.SetValueEx(k, value_name, 0, winreg.REG_SZ, value)
        print(f"✅ 已寫入：{value_name} = {value}")
    except Exception as e:
        print(f"❌ 寫入登錄檔失敗：{e}")


# ── 影片處理 ─────────────────────────────────────────

def video_screenshot(path: str, second: float = 0, output: str = ""):
    try:
        import cv2
        cap = cv2.VideoCapture(path)
        cap.set(cv2.CAP_PROP_POS_MSEC, second * 1000)
        ret, frame = cap.read()
        cap.release()
        out = output or path.replace(".mp4", f"_frame{int(second)}s.jpg")
        if ret:
            cv2.imwrite(out, frame)
            print(f"✅ 已擷取畫面：{out}")
        else:
            print("❌ 無法讀取影片")
    except Exception as e:
        print(f"❌ 影片截圖失敗：{e}")

def video_trim(path: str, start_sec: float, end_sec: float, output: str = ""):
    try:
        out = output or path.replace(".mp4", f"_trim.mp4")
        subprocess.run([
            "ffmpeg", "-y", "-i", path,
            "-ss", str(start_sec), "-to", str(end_sec),
            "-c", "copy", out
        ], capture_output=True)
        print(f"✅ 已剪輯：{out}")
    except Exception as e:
        print(f"❌ 影片剪輯失敗：{e}")


# ── 多螢幕管理 ───────────────────────────────────────

def monitor_list():
    try:
        from screeninfo import get_monitors
        for m in get_monitors():
            print(f"{'主螢幕 ' if m.is_primary else '副螢幕 '}{m.width}x{m.height} @({m.x},{m.y}) name={m.name}")
    except Exception as e:
        print(f"❌ 取得螢幕資訊失敗：{e}")


# ── Email 讀取 (IMAP) ────────────────────────────────

def email_read(host: str, user: str, password: str, folder: str = "INBOX", count: int = 5):
    try:
        import imapclient, email as _email
        from email.header import decode_header
        client = imapclient.IMAPClient(host, ssl=True)
        client.login(user, password)
        client.select_folder(folder)
        msgs = client.search(["ALL"])
        recent = msgs[-count:] if len(msgs) >= count else msgs
        results = []
        for uid in reversed(recent):
            raw = client.fetch([uid], ["RFC822"])[uid][b"RFC822"]
            msg = _email.message_from_bytes(raw)
            subj_raw, enc = decode_header(msg["Subject"])[0]
            subject = subj_raw.decode(enc or "utf-8") if isinstance(subj_raw, bytes) else subj_raw
            sender = msg["From"]
            date = msg["Date"]
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode(errors="replace")[:200]
                        break
            else:
                body = msg.get_payload(decode=True).decode(errors="replace")[:200]
            results.append(f"[{date}]\n寄件人：{sender}\n主旨：{subject}\n{body}\n{'─'*30}")
        client.logout()
        print("\n".join(results))
    except Exception as e:
        print(f"❌ 讀取郵件失敗：{e}")


# ── Google Calendar ───────────────────────────────────

def gcal_list(days: int = 7):
    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        from datetime import timezone, timedelta
        creds_path = Path("C:/Users/blue_/claude-telegram-bot/gcal_token.json")
        if not creds_path.exists():
            print("❌ 未找到 Google Calendar 憑證（gcal_token.json）")
            return
        creds = Credentials.from_authorized_user_file(str(creds_path))
        service = build("calendar", "v3", credentials=creds)
        now = datetime.now(timezone.utc)
        end = now + timedelta(days=days)
        events = service.events().list(
            calendarId="primary",
            timeMin=now.isoformat(),
            timeMax=end.isoformat(),
            maxResults=20, singleEvents=True, orderBy="startTime"
        ).execute().get("items", [])
        if not events:
            print(f"未來 {days} 天沒有行程")
            return
        for e in events:
            start = e["start"].get("dateTime", e["start"].get("date"))
            print(f"📅 {start}  {e.get('summary','（無標題）')}")
    except Exception as e:
        print(f"❌ 讀取行事曆失敗：{e}")

def gcal_add(title: str, start: str, end: str, description: str = ""):
    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        creds_path = Path("C:/Users/blue_/claude-telegram-bot/gcal_token.json")
        if not creds_path.exists():
            print("❌ 未找到 Google Calendar 憑證（gcal_token.json）")
            return
        creds = Credentials.from_authorized_user_file(str(creds_path))
        service = build("calendar", "v3", credentials=creds)
        event = {
            "summary": title,
            "description": description,
            "start": {"dateTime": start, "timeZone": "Asia/Taipei"},
            "end": {"dateTime": end, "timeZone": "Asia/Taipei"},
        }
        created = service.events().insert(calendarId="primary", body=event).execute()
        print(f"✅ 行程已新增：{created.get('htmlLink')}")
    except Exception as e:
        print(f"❌ 新增行程失敗：{e}")


# ── 全域快捷鍵 ───────────────────────────────────────

def global_hotkey_listen(hotkey: str, command: str, duration: float = 60.0):
    try:
        import keyboard as kb, time as t
        triggered = []
        def on_trigger():
            triggered.append(datetime.now().strftime("%H:%M:%S"))
            subprocess.run(command, shell=True)
        kb.add_hotkey(hotkey, on_trigger)
        print(f"🎹 監聽快捷鍵 [{hotkey}]，持續 {duration} 秒...")
        t.sleep(duration)
        kb.remove_all_hotkeys()
        print(f"✅ 共觸發 {len(triggered)} 次：{triggered}")
    except Exception as e:
        print(f"❌ 快捷鍵監聽失敗：{e}")


# ── Git 操作 ─────────────────────────────────────────

def git_op(action: str, repo: str = ".", message: str = "", remote: str = "origin", branch: str = "master"):
    try:
        import git as _git
        repo_obj = _git.Repo(repo)
        if action == "status":
            print(repo_obj.git.status())
        elif action == "log":
            for c in list(repo_obj.iter_commits())[:10]:
                print(f"{c.hexsha[:7]} [{c.authored_datetime.strftime('%Y-%m-%d %H:%M')}] {c.message.strip()[:60]}")
        elif action == "pull":
            result = repo_obj.remotes[remote].pull()
            print(f"✅ Pull 完成：{result[0].commit.hexsha[:7]}")
        elif action == "add":
            repo_obj.git.add(A=True)
            print("✅ 已 git add -A")
        elif action == "commit":
            repo_obj.index.commit(message or "auto commit")
            print(f"✅ 已 commit：{message}")
        elif action == "push":
            repo_obj.remotes[remote].push(branch)
            print(f"✅ 已 push 到 {remote}/{branch}")
        elif action == "diff":
            print(repo_obj.git.diff()[:3000] or "（無變更）")
    except Exception as e:
        print(f"❌ Git 操作失敗：{e}")


# ── 硬體監控進階 ─────────────────────────────────────

def hw_monitor():
    try:
        import psutil
        cpu_temp = ""
        try:
            temps = psutil.sensors_temperatures()
            if temps:
                for name, entries in temps.items():
                    for e in entries[:2]:
                        cpu_temp += f"{name}: {e.current}°C  "
        except Exception:
            cpu_temp = "（不支援溫度感測）"

        battery = psutil.sensors_battery()
        bat_str = f"{battery.percent:.0f}% {'充電中' if battery.power_plugged else '使用電池'}" if battery else "無電池"

        gpu_str = ""
        try:
            import GPUtil
            gpus = GPUtil.getGPUs()
            for g in gpus:
                gpu_str += f"\nGPU [{g.name}] 使用率：{g.load*100:.0f}% | 記憶體：{g.memoryUsed:.0f}/{g.memoryTotal:.0f}MB | 溫度：{g.temperature}°C"
        except Exception:
            gpu_str = "\nGPU：（未偵測到 NVIDIA GPU）"

        print(
            f"🌡 溫度：{cpu_temp}\n"
            f"🔋 電池：{bat_str}"
            f"{gpu_str}"
        )
    except Exception as e:
        print(f"❌ 硬體監控失敗：{e}")


# ── 自動報告生成 ─────────────────────────────────────

def report_gen(title: str, data_json: str, output: str = ""):
    try:
        import jinja2, json
        data = json.loads(data_json)
        out_path = output or str(Path.home() / "Desktop" / f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
        template_str = """<!DOCTYPE html>
<html><head><meta charset="utf-8">
<style>body{font-family:sans-serif;margin:40px}table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ccc;padding:8px;text-align:left}th{background:#4472C4;color:white}
tr:nth-child(even){background:#f2f2f2}h1{color:#4472C4}</style></head>
<body><h1>{{ title }}</h1><p>生成時間：{{ time }}</p>
{% for section, rows in data.items() %}
<h2>{{ section }}</h2>
{% if rows is iterable and rows is not string %}
{% if rows[0] is mapping %}
<table><tr>{% for k in rows[0].keys() %}<th>{{ k }}</th>{% endfor %}</tr>
{% for row in rows %}<tr>{% for v in row.values() %}<td>{{ v }}</td>{% endfor %}</tr>{% endfor %}
</table>
{% else %}<ul>{% for item in rows %}<li>{{ item }}</li>{% endfor %}</ul>{% endif %}
{% else %}<p>{{ rows }}</p>{% endif %}
{% endfor %}</body></html>"""
        tmpl = jinja2.Template(template_str)
        html = tmpl.render(title=title, data=data, time=datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        Path(out_path).write_text(html, encoding="utf-8")
        print(f"✅ 報告已生成：{out_path}")
    except Exception as e:
        print(f"❌ 報告生成失敗：{e}")


# ── Dropbox ──────────────────────────────────────────

def dropbox_upload(local: str, remote: str, token: str = ""):
    try:
        import dropbox as dbx
        tok = token or os.getenv("DROPBOX_TOKEN", "")
        if not tok:
            print("❌ 請設定 DROPBOX_TOKEN 環境變數")
            return
        d = dbx.Dropbox(tok)
        with open(local, "rb") as f:
            d.files_upload(f.read(), remote, mode=dbx.files.WriteMode.overwrite)
        print(f"✅ 已上傳到 Dropbox：{remote}")
    except Exception as e:
        print(f"❌ Dropbox 上傳失敗：{e}")

def dropbox_download(remote: str, local: str, token: str = ""):
    try:
        import dropbox as dbx
        tok = token or os.getenv("DROPBOX_TOKEN", "")
        if not tok:
            print("❌ 請設定 DROPBOX_TOKEN 環境變數")
            return
        d = dbx.Dropbox(tok)
        _, res = d.files_download(remote)
        Path(local).write_bytes(res.content)
        print(f"✅ 已從 Dropbox 下載：{local}")
    except Exception as e:
        print(f"❌ Dropbox 下載失敗：{e}")


# ── Docker ───────────────────────────────────────────

def docker_op(action: str, name: str = ""):
    try:
        import docker as _docker
        client = _docker.from_env()
        if action == "list":
            for c in client.containers.list(all=True):
                print(f"[{c.status}] {c.name}  {c.image.tags}")
        elif action == "start":
            client.containers.get(name).start()
            print(f"✅ 容器 [{name}] 已啟動")
        elif action == "stop":
            client.containers.get(name).stop()
            print(f"✅ 容器 [{name}] 已停止")
        elif action == "logs":
            logs = client.containers.get(name).logs(tail=50).decode(errors="replace")
            print(logs)
        elif action == "images":
            for img in client.images.list():
                print(f"{img.tags}  {img.short_id}")
    except Exception as e:
        print(f"❌ Docker 操作失敗：{e}")


# ── PDF 轉圖片 ───────────────────────────────────────

def pdf_to_images(path: str, output_dir: str = "", dpi: int = 150):
    try:
        import fitz
        doc = fitz.open(path)
        out = Path(output_dir) if output_dir else Path(path).parent / (Path(path).stem + "_imgs")
        out.mkdir(parents=True, exist_ok=True)
        for i, page in enumerate(doc):
            mat = fitz.Matrix(dpi/72, dpi/72)
            pix = page.get_pixmap(matrix=mat)
            img_path = str(out / f"page_{i+1}.png")
            pix.save(img_path)
        print(f"✅ 已轉換 {len(doc)} 頁到：{out}")
    except Exception as e:
        print(f"❌ PDF 轉圖片失敗：{e}")


# ── 條碼掃描 ─────────────────────────────────────────

def barcode_scan(image_path: str = ""):
    try:
        from pyzbar.pyzbar import decode
        from PIL import Image
        img = Image.open(image_path) if image_path else pyautogui.screenshot()
        results = decode(img)
        if not results:
            print("❌ 未偵測到條碼或 QR Code")
            return
        for r in results:
            print(f"類型：{r.type}  內容：{r.data.decode('utf-8', errors='replace')}")
    except Exception as e:
        print(f"❌ 條碼掃描失敗：{e}")


# ── NLP 文字分析 ─────────────────────────────────────

def nlp_summarize(text: str):
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=512,
            messages=[{"role": "user", "content": f"請用繁體中文摘要以下文字（100字以內）：\n\n{text}"}]
        )
        print(msg.content[0].text)
    except Exception as e:
        print(f"❌ 摘要失敗：{e}")

def nlp_sentiment(text: str):
    try:
        import anthropic
        client = anthropic.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
        msg = client.messages.create(
            model="claude-haiku-4-5-20251001",
            max_tokens=128,
            messages=[{"role": "user", "content": f"分析以下文字的情緒，只回覆：正面/負面/中性 + 一句說明：\n\n{text}"}]
        )
        print(msg.content[0].text)
    except Exception as e:
        print(f"❌ 情緒分析失敗：{e}")


# ── VPN 控制 ─────────────────────────────────────────

def vpn_control(action: str, name: str = "", user: str = "", password: str = ""):
    try:
        if action == "list":
            r = subprocess.run(["powershell.exe", "-Command",
                "Get-VpnConnection | Select-Object Name,ConnectionStatus | Format-Table"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout or "（未設定 VPN）")
        elif action == "connect":
            r = subprocess.run(["rasdial", name, user, password],
                capture_output=True, text=True, encoding="cp950", errors="replace")
            print(r.stdout.strip())
        elif action == "disconnect":
            r = subprocess.run(["rasdial", name, "/disconnect"],
                capture_output=True, text=True, encoding="cp950", errors="replace")
            print(r.stdout.strip())
    except Exception as e:
        print(f"❌ VPN 操作失敗：{e}")


# ── 系統還原點 ───────────────────────────────────────

def sys_restore(action: str, description: str = ""):
    try:
        if action == "create":
            ps = f"Checkpoint-Computer -Description '{description or 'Claude Auto Restore'}' -RestorePointType MODIFY_SETTINGS"
            r = subprocess.run(["powershell.exe", "-Command", ps],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout or "✅ 還原點已建立")
        elif action == "list":
            r = subprocess.run(["powershell.exe", "-Command",
                "Get-ComputerRestorePoint | Select-Object SequenceNumber,Description,CreationTime | Format-Table"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            print(r.stdout or "（無還原點）")
    except Exception as e:
        print(f"❌ 系統還原點操作失敗：{e}")


# ── 磁碟分析 ─────────────────────────────────────────

def disk_analyze(path: str = "C:/", top: int = 10):
    try:
        import psutil
        usage = psutil.disk_usage(path)
        print(f"磁碟：{path}\n總容量：{usage.total/1024**3:.1f} GB\n已使用：{usage.used/1024**3:.1f} GB ({usage.percent}%)\n可用：{usage.free/1024**3:.1f} GB\n")
        sizes = []
        try:
            for item in Path(path).iterdir():
                try:
                    if item.is_dir():
                        size = sum(f.stat().st_size for f in item.rglob("*") if f.is_file())
                    else:
                        size = item.stat().st_size
                    sizes.append((size, str(item)))
                except Exception:
                    pass
        except Exception:
            pass
        sizes.sort(reverse=True)
        print(f"佔用最多的前 {top} 個項目：")
        for size, name in sizes[:top]:
            print(f"  {size/1024**3:.2f} GB  {name}")
    except Exception as e:
        print(f"❌ 磁碟分析失敗：{e}")


# ── 人臉偵測 ─────────────────────────────────────────

def face_detect(image_path: str = "", output: str = ""):
    try:
        import cv2, numpy as np
        if not image_path:
            img_pil = pyautogui.screenshot()
            img = cv2.cvtColor(np.array(img_pil), cv2.COLOR_RGB2BGR)
        else:
            img = cv2.imread(image_path)
        cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
        gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)
        faces = cascade.detectMultiScale(gray, 1.1, 4)
        for (x, y, w, h) in faces:
            cv2.rectangle(img, (x, y), (x+w, y+h), (0, 255, 0), 2)
        out_path = output or str(Path.home() / "Desktop" / f"faces_{datetime.now().strftime('%H%M%S')}.jpg")
        cv2.imwrite(out_path, img)
        print(f"✅ 偵測到 {len(faces)} 張人臉，已存：{out_path}")
    except Exception as e:
        print(f"❌ 人臉偵測失敗：{e}")


# ── 影片轉 GIF ───────────────────────────────────────

def video_to_gif(path: str, start: float = 0, duration: float = 5.0, output: str = "", fps: int = 10):
    try:
        import imageio
        out = output or path.replace(".mp4", ".gif")
        reader = imageio.get_reader(path)
        meta = reader.get_meta_data()
        video_fps = meta.get("fps", 30)
        start_frame = int(start * video_fps)
        end_frame = int((start + duration) * video_fps)
        frames = []
        for i, frame in enumerate(reader):
            if i < start_frame:
                continue
            if i >= end_frame:
                break
            frames.append(frame)
        imageio.mimsave(out, frames, fps=fps)
        print(f"✅ GIF 已生成：{out}（{len(frames)} 幀）")
    except Exception as e:
        print(f"❌ 影片轉 GIF 失敗：{e}")


# ── Excel 圖表 ───────────────────────────────────────

def excel_chart(path: str, sheet: str, chart_type: str = "bar", title: str = ""):
    try:
        import openpyxl
        from openpyxl.chart import BarChart, LineChart, PieChart, Reference
        wb = openpyxl.load_workbook(path)
        ws = wb[sheet] if sheet in wb.sheetnames else wb.active
        max_row = ws.max_row
        max_col = ws.max_column
        chart_map = {"bar": BarChart, "line": LineChart, "pie": PieChart}
        chart = chart_map.get(chart_type, BarChart)()
        chart.title = title or sheet
        chart.style = 10
        data = Reference(ws, min_col=2, min_row=1, max_row=max_row, max_col=max_col)
        cats = Reference(ws, min_col=1, min_row=2, max_row=max_row)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(cats)
        ws.add_chart(chart, "A" + str(max_row + 2))
        wb.save(path)
        print(f"✅ 圖表已加入：{path}")
    except Exception as e:
        print(f"❌ Excel 圖表失敗：{e}")


# ── 網路速度測試 ─────────────────────────────────────

def speedtest_run():
    try:
        import speedtest as st
        print("⏳ 測速中（約需 30 秒）...")
        s = st.Speedtest()
        s.get_best_server()
        download = s.download() / 1_000_000
        upload = s.upload() / 1_000_000
        ping = s.results.ping
        server = s.results.server.get("name","")
        print(f"📶 網路速度測試結果\n下載：{download:.1f} Mbps\n上傳：{upload:.1f} Mbps\nPing：{ping:.0f} ms\n伺服器：{server}")
    except Exception as e:
        print(f"❌ 速度測試失敗：{e}")


# ── 螢幕截圖比對 ─────────────────────────────────────

def screenshot_compare(img1_path: str = "", img2_path: str = "", output: str = ""):
    try:
        import cv2, numpy as np
        from PIL import Image
        if not img1_path:
            img1 = np.array(pyautogui.screenshot())
        else:
            img1 = cv2.imread(img1_path)
        if not img2_path:
            import time as t; t.sleep(2)
            img2 = np.array(pyautogui.screenshot())
        else:
            img2 = cv2.imread(img2_path)
        img1_bgr = cv2.cvtColor(img1, cv2.COLOR_RGB2BGR) if img1_path == "" else img1
        img2_bgr = cv2.cvtColor(img2, cv2.COLOR_RGB2BGR) if img2_path == "" else img2
        h, w = min(img1_bgr.shape[0], img2_bgr.shape[0]), min(img1_bgr.shape[1], img2_bgr.shape[1])
        img1_bgr, img2_bgr = img1_bgr[:h,:w], img2_bgr[:h,:w]
        diff = cv2.absdiff(img1_bgr, img2_bgr)
        gray = cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY)
        _, thresh = cv2.threshold(gray, 30, 255, cv2.THRESH_BINARY)
        contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        result = img2_bgr.copy()
        cv2.drawContours(result, contours, -1, (0, 0, 255), 2)
        out = output or str(Path.home() / "Desktop" / f"diff_{datetime.now().strftime('%H%M%S')}.png")
        cv2.imwrite(out, result)
        changed = cv2.countNonZero(thresh)
        total = h * w
        pct = changed / total * 100
        print(f"✅ 差異：{pct:.2f}%，已標記差異區域：{out}")
    except Exception as e:
        print(f"❌ 截圖比對失敗：{e}")


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
        "run_python":      lambda: run_python(" ".join(args)),
        "run_shell":       lambda: run_shell(" ".join(args)),
        "word_read":       lambda: word_read(args[0]),
        "word_write":      lambda: word_write(args[0], " ".join(args[1:])),
        "excel_read":      lambda: excel_read(args[0], args[1] if len(args)>1 else None),
        "excel_write":     lambda: excel_write(args[0], args[1], " ".join(args[2:])),
        "pdf_read":        lambda: pdf_read(args[0]),
        "screen_diff":     lambda: screen_diff(float(args[0]) if args else 2.0, args[1] if len(args)>1 else "full"),
        "scrape":          lambda: scrape(args[0], args[1] if len(args)>1 else "body"),
        "img_edit":        lambda: img_edit(args[0], args[1], *args[2:]),
        "gdrive_upload":   lambda: gdrive_upload(args[0], args[1] if len(args)>1 else "root"),
        "gdrive_download": lambda: gdrive_download(args[0], args[1]),
        "db_query":        lambda: db_query(args[0], " ".join(args[1:])),
        "db_mysql":        lambda: db_mysql(args[0], args[1], " ".join(args[2:])),
        "encrypt":         lambda: encrypt(args[0], args[1]),
        "decrypt":         lambda: decrypt(args[0], args[1]),
        "clipboard_watch": lambda: clipboard_watch(float(args[0]) if args else 30.0),
        "qr_gen":          lambda: qr_gen(args[0], args[1] if len(args)>1 else ""),
        "qr_scan":         lambda: qr_scan(args[0] if args else ""),
        "screen_record":   lambda: screen_record(float(args[0]) if args else 10.0, args[1] if len(args)>1 else ""),
        "webcam":          lambda: webcam_capture(args[0] if args else ""),
        "translate":       lambda: translate(" ".join(args[:1] if len(args)==1 else [args[0]]), args[1] if len(args)>1 else "zh-TW", args[2] if len(args)>2 else "auto"),
        "chart":           lambda: chart(args[0], args[1], args[2] if len(args)>2 else "", args[3] if len(args)>3 else ""),
        "pptx_read":       lambda: pptx_read(args[0]),
        "pptx_create":     lambda: pptx_create(args[0], " ".join(args[1:])),
        "api_call":        lambda: api_call(args[0], args[1], args[2] if len(args)>2 else "{}", args[3] if len(args)>3 else "{}"),
        "watchdog":        lambda: watchdog(args[0], args[1], float(args[2]) if len(args)>2 else 60.0),
        "ssh_run":         lambda: ssh_run(args[0], args[1], args[2], " ".join(args[3:])),
        "sftp_upload":     lambda: sftp_upload(args[0], args[1], args[2], args[3], args[4]),
        "sftp_download":   lambda: sftp_download(args[0], args[1], args[2], args[3], args[4]),
        "net_ping":        lambda: net_ping(args[0], int(args[1]) if len(args)>1 else 4),
        "net_traceroute":  lambda: net_traceroute(args[0]),
        "net_portscan":    lambda: net_portscan(args[0], args[1] if len(args)>1 else "22,80,443,3306,3389,8080"),
        "win_service":     lambda: win_service(args[0], args[1] if len(args)>1 else ""),
        "pdf_merge":       lambda: pdf_merge(args[0], args[1]),
        "pdf_split":       lambda: pdf_split(args[0], args[1]),
        "pdf_watermark":   lambda: pdf_watermark(args[0], args[1], args[2] if len(args)>2 else ""),
        "audio_convert":   lambda: audio_convert(args[0], args[1]),
        "audio_trim":      lambda: audio_trim(args[0], int(args[1]), int(args[2]), args[3] if len(args)>3 else ""),
        "discord_notify":  lambda: discord_notify(args[0], " ".join(args[1:])),
        "line_notify":     lambda: line_notify(args[0], " ".join(args[1:])),
        "disk_clean":      lambda: disk_clean(args[0] if args else "list"),
        "backup":          lambda: backup(args[0], args[1]),
        "registry_read":   lambda: registry_read(args[0], args[1] if len(args)>1 else ""),
        "registry_write":  lambda: registry_write(args[0], args[1], " ".join(args[2:])),
        "video_screenshot":lambda: video_screenshot(args[0], float(args[1]) if len(args)>1 else 0, args[2] if len(args)>2 else ""),
        "video_trim":      lambda: video_trim(args[0], float(args[1]), float(args[2]), args[3] if len(args)>3 else ""),
        "monitor_list":    lambda: monitor_list(),
        "email_read":      lambda: email_read(args[0], args[1], args[2], args[3] if len(args)>3 else "INBOX", int(args[4]) if len(args)>4 else 5),
        "gcal_list":       lambda: gcal_list(int(args[0]) if args else 7),
        "gcal_add":        lambda: gcal_add(args[0], args[1], args[2], args[3] if len(args)>3 else ""),
        "global_hotkey":   lambda: global_hotkey_listen(args[0], args[1], float(args[2]) if len(args)>2 else 60.0),
        "git_op":          lambda: git_op(args[0], args[1] if len(args)>1 else ".", args[2] if len(args)>2 else "", args[3] if len(args)>3 else "origin", args[4] if len(args)>4 else "master"),
        "hw_monitor":      lambda: hw_monitor(),
        "report_gen":      lambda: report_gen(args[0], args[1], args[2] if len(args)>2 else ""),
        "dropbox_upload":  lambda: dropbox_upload(args[0], args[1], args[2] if len(args)>2 else ""),
        "dropbox_download":lambda: dropbox_download(args[0], args[1], args[2] if len(args)>2 else ""),
        "docker_op":       lambda: docker_op(args[0], args[1] if len(args)>1 else ""),
        "pdf_to_images":   lambda: pdf_to_images(args[0], args[1] if len(args)>1 else "", int(args[2]) if len(args)>2 else 150),
        "barcode_scan":    lambda: barcode_scan(args[0] if args else ""),
        "nlp_summarize":   lambda: nlp_summarize(" ".join(args)),
        "nlp_sentiment":   lambda: nlp_sentiment(" ".join(args)),
        "vpn_control":     lambda: vpn_control(args[0], args[1] if len(args)>1 else "", args[2] if len(args)>2 else "", args[3] if len(args)>3 else ""),
        "sys_restore":     lambda: sys_restore(args[0], " ".join(args[1:]) if len(args)>1 else ""),
        "disk_analyze":    lambda: disk_analyze(args[0] if args else "C:/", int(args[1]) if len(args)>1 else 10),
        "face_detect":     lambda: face_detect(args[0] if args else "", args[1] if len(args)>1 else ""),
        "video_to_gif":    lambda: video_to_gif(args[0], float(args[1]) if len(args)>1 else 0, float(args[2]) if len(args)>2 else 5.0, args[3] if len(args)>3 else "", int(args[4]) if len(args)>4 else 10),
        "excel_chart":     lambda: excel_chart(args[0], args[1], args[2] if len(args)>2 else "bar", args[3] if len(args)>3 else ""),
        "speedtest":       lambda: speedtest_run(),
        "screenshot_compare": lambda: screenshot_compare(args[0] if args else "", args[1] if len(args)>1 else "", args[2] if len(args)>2 else ""),
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
