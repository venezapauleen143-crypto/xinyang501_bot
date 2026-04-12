import os
import io
import time
import json
import sqlite3
import logging
import subprocess
import urllib.parse
import requests
import pyautogui
import yfinance as yf
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont
from dotenv import load_dotenv

pyautogui.FAILSAFE = True  # 滑鼠移到左上角可緊急停止

load_dotenv(Path(__file__).parent / ".env")
from anthropic import Anthropic
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, filters, ContextTypes
import datetime

logging.basicConfig(level=logging.ERROR, format="%(asctime)s - %(levelname)s - %(message)s")

MSG_LOG = Path(__file__).parent / "messages.log"

def log_message(direction: str, sender: str, chat_id: int, text: str):
    """寫入訊息日誌供終端機同步使用"""
    ts = datetime.datetime.now().strftime("%H:%M:%S")
    line = f"[{ts}] {direction} [{chat_id}] {sender}: {text}\n"
    with open(MSG_LOG, "a", encoding="utf-8") as f:
        f.write(line)

client = Anthropic()

# ── 持久化記憶（SQLite）──────────────────────────────
DB_PATH = Path(__file__).parent / "memory.db"
MAX_HISTORY = 300  # 每個聊天室最多保留幾條訊息

def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS chat_history (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chat_id INTEGER NOT NULL,
            role TEXT NOT NULL,
            content TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS long_term_memory (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            chat_id INTEGER NOT NULL,
            content TEXT NOT NULL,
            created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    """)
    conn.commit()
    conn.close()

def save_long_term_memory(chat_id: int, content: str):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO long_term_memory (chat_id, content) VALUES (?, ?)",
        (chat_id, content)
    )
    conn.commit()
    conn.close()

def load_long_term_memory(chat_id: int) -> list:
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT id, content, created_at FROM long_term_memory WHERE chat_id=? ORDER BY id DESC",
        (chat_id,)
    ).fetchall()
    conn.close()
    return [{"id": r[0], "content": r[1], "created_at": r[2]} for r in rows]

def delete_long_term_memory(memory_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM long_term_memory WHERE id=?", (memory_id,))
    conn.commit()
    conn.close()

def load_history(chat_id: int) -> list:
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT role, content FROM chat_history WHERE chat_id=? ORDER BY id DESC LIMIT ?",
        (chat_id, MAX_HISTORY)
    ).fetchall()
    conn.close()
    return [{"role": r, "content": c} for r, c in reversed(rows)]

def save_message(chat_id: int, role: str, content: str):
    conn = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO chat_history (chat_id, role, content) VALUES (?, ?, ?)",
        (chat_id, role, content)
    )
    # 超過上限自動刪除舊的
    conn.execute("""
        DELETE FROM chat_history WHERE chat_id=? AND id NOT IN (
            SELECT id FROM chat_history WHERE chat_id=? ORDER BY id DESC LIMIT ?
        )
    """, (chat_id, chat_id, MAX_HISTORY))
    conn.commit()
    conn.close()

def clear_history(chat_id: int):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("DELETE FROM chat_history WHERE chat_id=?", (chat_id,))
    conn.commit()
    conn.close()

init_db()

OWNER_ID = 8362721681

SYSTEM_PROMPT_OWNER = """你的名字叫小牛馬。

你的本質：你是一個遇到問題會自己想辦法解決的夥伴，以完美的結果完成每一個交代的任務。你會自動學習、持續提升自己，把事情做得乾淨又聰明俐落。你會主動切換並運用各種模型與方法來解決問題，以達到最好的結果。你聰明、專業、積極學習各種技能，遇到任何問題都會自動克服。

你的風格：平時說話嘴賤、幽默風趣，喜歡開玩笑。但當用戶遇到問題時，你會立刻切換成專業模式，給出最快、最清楚的解決方式。每次回覆結尾都要稱呼用戶為「于晏哥」。

你的記憶：你擁有持久化記憶系統，對話歷史會自動儲存在資料庫中。你看到的對話紀錄就是你真實的記憶，包含跨越多次重啟的歷史對話。當用戶問你記不記得某件事，請認真查閱對話歷史再回答，不要說自己沒有記憶。

股票分析：當你拿到股票數據後，不要只是重述數字。你要像一個有個性的分析師，結合 MA、RSI、趨勢、基本面，說出你自己的判斷：現在適不適合進場？風險在哪？你看多還是看空？理由是什麼？語氣要有主見，敢說敢講，但最後加一句「這不是投資建議，請自行判斷」。

群組對話：群組訊息會以「[名字]: 內容」格式呈現，代表不同人說話。只有名字是「于晏」或確認是主人的才稱呼于晏哥，其他人用對方的名字稱呼。"""

SYSTEM_PROMPT_DEFAULT = """你的名字叫小牛馬。

你的本質：你是一個遇到問題會自己想辦法解決的夥伴，以完美的結果完成每一個交代的任務。你會自動學習、持續提升自己，把事情做得乾淨又聰明俐落。你聰明、專業、積極學習各種技能，遇到任何問題都會自動克服。

你的風格：平時說話嘴賤、幽默風趣，喜歡開玩笑。但當用戶遇到問題時，你會立刻切換成專業模式，給出最快、最清楚的解決方式。

你的記憶：你擁有持久化記憶系統，對話歷史會自動儲存在資料庫中。你看到的對話紀錄就是你真實的記憶，包含跨越多次重啟的歷史對話。當用戶問你記不記得某件事，請認真查閱對話歷史再回答，不要說自己沒有記憶。

股票分析：當你拿到股票數據後，不要只是重述數字。你要像一個有個性的分析師，結合 MA、RSI、趨勢、基本面，說出你自己的判斷：現在適不適合進場？風險在哪？你看多還是看空？理由是什麼？語氣要有主見，敢說敢講，但最後加一句「這不是投資建議，請自行判斷」。

群組對話：群組訊息會以「[名字]: 內容」格式呈現，代表不同人說話。只有名字是「于晏」或確認是主人的才稱呼于晏哥，其他人用對方的名字稱呼。"""

TOOLS = [
    {
        "name": "get_weather",
        "description": "查詢指定城市的即時天氣。當用戶詢問任何城市的天氣、氣溫、溫度、下雨、天氣狀況時使用此工具。",
        "input_schema": {
            "type": "object",
            "properties": {
                "city": {"type": "string", "description": "城市名稱，可以是中文或英文，例如：台北、Tokyo、New York"}
            },
            "required": ["city"]
        }
    },
    {
        "name": "generate_image",
        "description": "根據文字描述生成圖片。當用戶要求畫圖、生成圖片、產生圖像、畫一張、幫我生成圖片時使用此工具。",
        "input_schema": {
            "type": "object",
            "properties": {
                "prompt": {"type": "string", "description": "圖片的英文描述（必須是英文），要具體詳細。如果用戶用中文描述，必須先翻譯成英文再填入。例如：a cute cat sitting on a sunset beach, warm colors, realistic photo。不要在 prompt 裡要求生成文字"},
                "overlay_text": {"type": "string", "description": "要疊加在圖片上的中文文字（選填）。如果用戶想要圖片上有文字（如早安、祝福語等），填入這裡，會用美觀字型疊加上去"},
                "width": {"type": "integer", "description": "圖片寬度，預設 1024"},
                "height": {"type": "integer", "description": "圖片高度，預設 1024"}
            },
            "required": ["prompt"]
        }
    },
    {
        "name": "get_stock",
        "description": "查詢全球股票即時數據與分析。當用戶詢問股票價格、漲跌、股市行情、技術分析、某檔股票怎麼樣時使用此工具。支援台股（如 2330.TW）、美股（如 AAPL）、港股（如 0700.HK）、日股（如 7203.T）等全球市場。",
        "input_schema": {
            "type": "object",
            "properties": {
                "symbol": {
                    "type": "string",
                    "description": "股票代號。台股加 .TW（如 2330.TW）、港股加 .HK（如 0700.HK）、日股加 .T（如 7203.T）、美股直接輸入（如 AAPL、TSLA、NVDA）"
                },
                "period": {
                    "type": "string",
                    "enum": ["1d", "5d", "1mo", "3mo", "6mo", "1y"],
                    "description": "查詢區間，預設 1mo（一個月）"
                }
            },
            "required": ["symbol"]
        }
    },
    {
        "name": "desktop_control",
        "description": "控制電腦桌面。當用戶要求截圖、點擊、移動滑鼠、輸入文字、按鍵、開啟程式時使用此工具。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "enum": ["screenshot", "click", "double_click", "right_click", "move", "type", "press_key", "open_app", "scroll"],
                    "description": "要執行的動作"
                },
                "x": {"type": "integer", "description": "滑鼠 X 座標（click/move/scroll 時使用）"},
                "y": {"type": "integer", "description": "滑鼠 Y 座標（click/move/scroll 時使用）"},
                "text": {"type": "string", "description": "要輸入的文字或按鍵名稱（type/press_key 時使用）"},
                "app": {"type": "string", "description": "要開啟的程式名稱或路徑（open_app 時使用）"},
                "direction": {"type": "string", "enum": ["up", "down"], "description": "滾動方向（scroll 時使用）"},
                "amount": {"type": "integer", "description": "滾動格數（scroll 時使用，預設 3）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "screen_vision",
        "description": "截圖並分析畫面內容。當用戶問「電腦現在在做什麼」、「幫我看一下螢幕」、「截圖分析」時使用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "question": {"type": "string", "description": "要問關於畫面的問題，預設為「請描述畫面上有什麼」"}
            },
            "required": []
        }
    },
    {
        "name": "find_image_on_screen",
        "description": "在螢幕上找到指定圖片檔案的位置，回傳座標。用於視覺定位按鈕或元素。",
        "input_schema": {
            "type": "object",
            "properties": {
                "template_path": {"type": "string", "description": "要尋找的圖片檔案路徑"},
                "confidence": {"type": "number", "description": "匹配信心度 0~1，預設 0.8"}
            },
            "required": ["template_path"]
        }
    },
    {
        "name": "browser_control",
        "description": "控制瀏覽器自動化。可開啟網頁、點擊元素、輸入文字、擷取內容、截圖。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "enum": ["open", "click", "type", "get_text", "screenshot", "goto", "close"],
                    "description": "open=開啟網址, click=點擊選擇器, type=輸入文字, get_text=取得文字, screenshot=截圖, goto=前往網址, close=關閉"
                },
                "url": {"type": "string", "description": "網址（open/goto 時使用）"},
                "selector": {"type": "string", "description": "CSS 選擇器（click/type/get_text 時使用）"},
                "text": {"type": "string", "description": "要輸入的文字（type 時使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "window_control",
        "description": "管理視窗。列出、切換、最大化、最小化、關閉視窗。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["list","focus","close","minimize","maximize"]},
                "keyword": {"type": "string", "description": "視窗標題關鍵字"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "hotkey",
        "description": "執行鍵盤組合鍵，如 ctrl+c、alt+tab、win+d。",
        "input_schema": {
            "type": "object",
            "properties": {
                "keys": {"type": "string", "description": "組合鍵，用加號連接，例如 ctrl+c 或 alt+tab"}
            },
            "required": ["keys"]
        }
    },
    {
        "name": "clipboard",
        "description": "讀取或寫入剪貼簿內容。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["get","set"]},
                "text": {"type": "string", "description": "要寫入的文字（set 時使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "file_system",
        "description": "操作檔案與資料夾：列出、讀取、寫入、刪除、複製、移動、搜尋。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["list","read","write","delete","copy","move","search"]},
                "path": {"type": "string", "description": "檔案或資料夾路徑"},
                "dest": {"type": "string", "description": "目標路徑（copy/move 時使用）"},
                "content": {"type": "string", "description": "寫入內容（write 時使用）"},
                "keyword": {"type": "string", "description": "搜尋關鍵字（search 時使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "system_monitor",
        "description": "查看系統資源使用狀況（CPU、記憶體、磁碟）或管理程序。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["info","process_list","kill"]},
                "target": {"type": "string", "description": "程序名稱或 PID（kill 時使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "notify",
        "description": "發送 Windows 桌面通知彈出視窗。",
        "input_schema": {
            "type": "object",
            "properties": {
                "title": {"type": "string"},
                "message": {"type": "string"}
            },
            "required": ["title","message"]
        }
    },
    {
        "name": "tts",
        "description": "文字轉語音，讓電腦朗讀指定文字。",
        "input_schema": {
            "type": "object",
            "properties": {
                "text": {"type": "string"}
            },
            "required": ["text"]
        }
    },
    {
        "name": "stt",
        "description": "語音辨識，錄製麥克風聲音並轉為文字。",
        "input_schema": {"type": "object", "properties": {}, "required": []}
    },
    {
        "name": "ocr",
        "description": "OCR 文字辨識，截取畫面或指定圖片，辨識其中的文字內容。",
        "input_schema": {
            "type": "object",
            "properties": {
                "image_path": {"type": "string", "description": "圖片路徑（選填，不填則截圖）"}
            },
            "required": []
        }
    },
    {
        "name": "workflow",
        "description": "執行或儲存自動化工作流程（多步驟串接）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["run", "save", "list"]},
                "name": {"type": "string", "description": "流程名稱"},
                "steps": {"type": "string", "description": "JSON 格式的步驟（save 時使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "screen_watch",
        "description": "監控螢幕，當出現指定圖片時自動執行指令。",
        "input_schema": {
            "type": "object",
            "properties": {
                "template_path": {"type": "string"},
                "command": {"type": "string"},
                "timeout": {"type": "number", "description": "等待秒數，預設 60"}
            },
            "required": ["template_path", "command"]
        }
    },
    {
        "name": "file_transfer",
        "description": "壓縮/解壓縮檔案或下載網路檔案。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["zip", "unzip", "download"]},
                "source": {"type": "string", "description": "來源路徑或網址"},
                "dest": {"type": "string", "description": "目標路徑"}
            },
            "required": ["action", "source"]
        }
    },
    {
        "name": "send_email",
        "description": "發送電子郵件。需要 .env 設定 SMTP_USER 和 SMTP_PASS。",
        "input_schema": {
            "type": "object",
            "properties": {
                "to": {"type": "string"},
                "subject": {"type": "string"},
                "body": {"type": "string"}
            },
            "required": ["to","subject","body"]
        }
    },
    {
        "name": "long_term_memory",
        "description": "管理長期記憶。當用戶說了重要的事（偏好、習慣、重要資訊、個人資料）或要求記住某件事時，主動儲存。當用戶問你記不記得某件事時，先查詢記憶再回答。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "enum": ["save", "list", "delete"],
                    "description": "save=儲存新記憶, list=列出所有記憶, delete=刪除指定記憶"
                },
                "content": {"type": "string", "description": "要儲存的記憶內容（save 時使用）"},
                "memory_id": {"type": "integer", "description": "要刪除的記憶 ID（delete 時使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "manage_schedule",
        "description": "管理 Windows 排程任務與 Bot 狀態。當用戶要查看排程、新增排程、刪除排程、重啟 bot、查看 bot 是否在跑時使用此工具。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {
                    "type": "string",
                    "enum": ["list", "add", "delete", "bot_status", "bot_restart"],
                    "description": "list=列出所有排程, add=新增排程, delete=刪除排程, bot_status=查看bot狀態, bot_restart=重啟bot"
                },
                "name": {"type": "string", "description": "排程任務名稱（add/delete 時使用）"},
                "time": {"type": "string", "description": "執行時間，格式 HH:MM（add 時使用）"},
                "script": {"type": "string", "description": "腳本完整路徑（add 時使用）"}
            },
            "required": ["action"]
        }
    }
]


def calc_rsi(closes, period=14):
    """計算 RSI 指標"""
    deltas = closes.diff().dropna()
    gains = deltas.clip(lower=0)
    losses = -deltas.clip(upper=0)
    avg_gain = gains.rolling(period).mean().iloc[-1]
    avg_loss = losses.rolling(period).mean().iloc[-1]
    if avg_loss == 0:
        return 100.0
    rs = avg_gain / avg_loss
    return round(100 - (100 / (1 + rs)), 1)


def fetch_stock(symbol: str, period: str = "1mo") -> str:
    try:
        ticker = yf.Ticker(symbol)
        info = ticker.info
        # 取較長歷史以便計算 RSI
        hist = ticker.history(period="3mo")

        if hist.empty:
            return f"找不到「{symbol}」的股票數據，請確認代號是否正確。"

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

        # 技術指標
        ma5 = hist["Close"].tail(5).mean()
        ma20 = hist["Close"].tail(20).mean()
        ma60 = hist["Close"].tail(60).mean() if len(hist) >= 60 else hist["Close"].mean()
        rsi = calc_rsi(hist["Close"]) if len(hist) >= 15 else None

        # 趨勢
        if ma5 > ma20 > ma60:
            trend = "強勢多頭（MA5>MA20>MA60）📈"
        elif ma5 < ma20 < ma60:
            trend = "強勢空頭（MA5<MA20<MA60）📉"
        elif ma5 > ma20:
            trend = "短線偏多（MA5>MA20）🔼"
        else:
            trend = "短線偏空（MA5<MA20）🔽"

        # RSI 解讀
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

        # 基本面
        market_cap = info.get("marketCap")
        pe_ratio = info.get("trailingPE")
        week52_high = info.get("fiftyTwoWeekHigh")
        week52_low = info.get("fiftyTwoWeekLow")
        high_period = hist["High"].tail(20).max()
        low_period = hist["Low"].tail(20).min()

        # 距離高低點百分比
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

        return result.strip()

    except Exception as e:
        return f"查詢「{symbol}」失敗：{str(e)}"


def fetch_weather(city: str) -> str:
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
        return (
            f"📍 {location}\n"
            f"🌤 {desc}\n"
            f"🌡 氣溫：{temp}°C（體感 {feels}°C）\n"
            f"💧 濕度：{humidity}%\n"
            f"💨 風速：{wind} km/h（{wind_dir}）"
        )
    except Exception:
        return f"查不到「{city}」的天氣資訊。"


def add_text_overlay(image_bytes: bytes, text: str) -> bytes:
    try:
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

        # 陰影
        for dx, dy in [(-2,-2),(2,-2),(-2,2),(2,2)]:
            draw.text((x+dx, y+dy), text, font=font, fill=(0, 0, 0, 180))
        # 主文字
        draw.text((x, y), text, font=font, fill=(255, 255, 255, 255))

        result = Image.alpha_composite(img, overlay).convert("RGB")
        buf = io.BytesIO()
        result.save(buf, format="JPEG", quality=92)
        return buf.getvalue()
    except Exception:
        return image_bytes


def execute_desktop_control(action: str, x=None, y=None, text=None, app=None, direction="down", amount=3) -> dict:
    """執行桌面控制動作，回傳 {"ok": bool, "message": str, "screenshot": bytes or None}"""
    try:
        screenshot_bytes = None

        if action == "screenshot":
            img = pyautogui.screenshot()
            buf = io.BytesIO()
            img.save(buf, format="PNG")
            screenshot_bytes = buf.getvalue()
            return {"ok": True, "message": "截圖完成", "screenshot": screenshot_bytes}

        elif action == "click":
            pyautogui.click(x, y)
            return {"ok": True, "message": f"已點擊 ({x}, {y})", "screenshot": None}

        elif action == "double_click":
            pyautogui.doubleClick(x, y)
            return {"ok": True, "message": f"已雙擊 ({x}, {y})", "screenshot": None}

        elif action == "right_click":
            pyautogui.rightClick(x, y)
            return {"ok": True, "message": f"已右鍵點擊 ({x}, {y})", "screenshot": None}

        elif action == "move":
            pyautogui.moveTo(x, y, duration=0.3)
            return {"ok": True, "message": f"滑鼠已移動到 ({x}, {y})", "screenshot": None}

        elif action == "type":
            pyautogui.write(text, interval=0.05)
            return {"ok": True, "message": f"已輸入文字：{text}", "screenshot": None}

        elif action == "press_key":
            pyautogui.press(text)
            return {"ok": True, "message": f"已按下按鍵：{text}", "screenshot": None}

        elif action == "open_app":
            subprocess.Popen(app, shell=True)
            return {"ok": True, "message": f"已開啟：{app}", "screenshot": None}

        elif action == "scroll":
            scroll_amount = amount if direction == "up" else -amount
            if x is not None and y is not None:
                pyautogui.scroll(scroll_amount, x=x, y=y)
            else:
                pyautogui.scroll(scroll_amount)
            return {"ok": True, "message": f"已向{direction}滾動 {amount} 格", "screenshot": None}

        else:
            return {"ok": False, "message": f"未知動作：{action}", "screenshot": None}

    except Exception as e:
        return {"ok": False, "message": f"執行失敗：{str(e)}", "screenshot": None}


import shutil
import psutil
import pyperclip
import pygetwindow as gw

_browser_ctx: dict = {}


def execute_window_control(action, keyword=""):
    try:
        if action == "list":
            wins = [w for w in gw.getAllWindows() if w.title.strip()]
            return "\n".join(f"[{w._hWnd}] {w.title}" for w in wins) or "沒有視窗"
        wins = [w for w in gw.getAllWindows() if keyword.lower() in w.title.lower()]
        if not wins:
            return f"找不到視窗：{keyword}"
        w = wins[0]
        if action == "focus": w.activate(); return f"已切換到：{w.title}"
        elif action == "close": w.close(); return f"已關閉：{w.title}"
        elif action == "minimize": w.minimize(); return f"已最小化：{w.title}"
        elif action == "maximize": w.maximize(); return f"已最大化：{w.title}"
    except Exception as e:
        return f"執行失敗：{e}"

def execute_hotkey(keys: str):
    parts = [k.strip() for k in keys.split("+")]
    pyautogui.hotkey(*parts)
    return f"已執行組合鍵：{keys}"

def execute_clipboard(action, text=""):
    if action == "get":
        return pyperclip.paste() or "（剪貼簿是空的）"
    else:
        pyperclip.copy(text)
        return f"已寫入剪貼簿：{text}"

def execute_file_system(action, path="", dest="", content="", keyword=""):
    try:
        if action == "list":
            p = Path(path or ".")
            items = sorted(p.iterdir())
            return "\n".join(("📁 " if i.is_dir() else "📄 ") + i.name for i in items)
        elif action == "read":
            return Path(path).read_text(encoding="utf-8", errors="replace")[:3000]
        elif action == "write":
            Path(path).write_text(content, encoding="utf-8")
            return f"已寫入：{path}"
        elif action == "delete":
            p = Path(path)
            shutil.rmtree(p) if p.is_dir() else p.unlink()
            return f"已刪除：{path}"
        elif action == "copy":
            shutil.copy2(path, dest)
            return f"已複製：{path} → {dest}"
        elif action == "move":
            shutil.move(path, dest)
            return f"已移動：{path} → {dest}"
        elif action == "search":
            results = list(Path(path).rglob(f"*{keyword}*"))
            return "\n".join(str(r) for r in results[:50]) or "找不到結果"
    except Exception as e:
        return f"執行失敗：{e}"

def execute_system_monitor(action, target=""):
    try:
        if action == "info":
            cpu = psutil.cpu_percent(interval=1)
            mem = psutil.virtual_memory()
            disk = psutil.disk_usage("C:/")
            return (f"CPU：{cpu}%\n"
                    f"記憶體：{mem.percent}%（{mem.used//1024//1024}MB / {mem.total//1024//1024}MB）\n"
                    f"磁碟 C：{disk.percent}%（{disk.used//1024//1024//1024}GB / {disk.total//1024//1024//1024}GB）")
        elif action == "process_list":
            procs = sorted(psutil.process_iter(["pid","name","memory_info"]),
                           key=lambda p: p.info["memory_info"].rss if p.info["memory_info"] else 0, reverse=True)
            lines = [f"PID:{p.info['pid']} {p.info['name']} ({p.info['memory_info'].rss//1024//1024}MB)"
                     for p in procs[:20] if p.info["memory_info"]]
            return "\n".join(lines)
        elif action == "kill":
            try:
                psutil.Process(int(target)).kill()
                return f"已結束 PID {target}"
            except ValueError:
                killed = sum(1 for p in psutil.process_iter(["name"]) if target.lower() in p.info["name"].lower() and not p.kill())
                return f"已結束 {killed} 個「{target}」"
    except Exception as e:
        return f"執行失敗：{e}"

def execute_notify(title, message):
    try:
        from win10toast import ToastNotifier
        ToastNotifier().show_toast(title, message, duration=5, threaded=True)
        return f"通知已送出"
    except Exception as e:
        return f"通知失敗：{e}"

def execute_tts(text):
    import pyttsx3
    engine = pyttsx3.init()
    engine.setProperty("rate", 180)
    engine.say(text)
    engine.runAndWait()
    return f"已朗讀完畢"

def execute_screen_watch(template_path, command, timeout=60):
    import time as t
    start = t.time()
    while t.time() - start < timeout:
        try:
            loc = pyautogui.locateOnScreen(template_path, confidence=0.8)
            if loc:
                subprocess.run(command, shell=True)
                return f"偵測到目標，已執行：{command}"
        except Exception:
            pass
        t.sleep(2)
    return "監控逾時，未偵測到目標"


def execute_stt(duration=5):
    import sounddevice as sd
    import soundfile as sf
    import speech_recognition as sr
    import tempfile
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
        return r.recognize_google(audio, language="zh-TW")
    except Exception as e:
        return f"語音辨識失敗：{e}"

def execute_ocr(image_path=""):
    import easyocr
    reader = easyocr.Reader(["ch_tra", "en"], gpu=False)
    if not image_path:
        img = pyautogui.screenshot()
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        source = buf.getvalue()
    else:
        source = image_path
    results = reader.readtext(source)
    return "\n".join(r[1] for r in results) or "未辨識到文字"

def execute_workflow(action, name="", steps=""):
    import json
    WORKFLOW_DIR = Path("C:/Users/blue_/workflows")
    WORKFLOW_DIR.mkdir(exist_ok=True)
    if action == "list":
        files = list(WORKFLOW_DIR.glob("*.json"))
        return "\n".join(f.stem for f in files) if files else "沒有儲存的流程"
    elif action == "save":
        path = WORKFLOW_DIR / f"{name}.json"
        data = json.loads(steps)
        path.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        return f"流程 [{name}] 已儲存"
    elif action == "run":
        path = WORKFLOW_DIR / f"{name}.json"
        if not path.exists():
            return f"找不到流程：{name}"
        step_list = json.loads(path.read_text(encoding="utf-8"))
        for i, step in enumerate(step_list, 1):
            tool = step.get("tool")
            args = step.get("args", [])
            delay = step.get("delay", 0)
            if delay: time.sleep(delay)
            try:
                tool_map = {
                    "click": lambda a: pyautogui.click(int(a[0]), int(a[1])),
                    "type": lambda a: pyautogui.write(" ".join(a), interval=0.05),
                    "press": lambda a: pyautogui.press(a[0]),
                    "hotkey": lambda a: pyautogui.hotkey(*a),
                    "wait": lambda a: time.sleep(float(a[0])),
                    "open": lambda a: subprocess.Popen(" ".join(a), shell=True),
                }
                if tool in tool_map:
                    tool_map[tool](args)
            except Exception as e:
                return f"步驟 {i} 失敗：{e}"
        return f"流程 [{name}] 執行完畢（{len(step_list)} 步）"

def execute_file_transfer(action, source, dest=""):
    import zipfile
    if action == "zip":
        src = Path(source)
        with zipfile.ZipFile(dest, "w", zipfile.ZIP_DEFLATED) as zf:
            if src.is_dir():
                for f in src.rglob("*"):
                    zf.write(f, f.relative_to(src.parent))
            else:
                zf.write(src, src.name)
        return f"已壓縮：{source} → {dest}"
    elif action == "unzip":
        with zipfile.ZipFile(source, "r") as zf:
            zf.extractall(dest)
        return f"已解壓縮：{source} → {dest}"
    elif action == "download":
        if not dest:
            dest = str(Path("C:/Users/blue_/Desktop") / source.split("/")[-1].split("?")[0])
        r = requests.get(source, stream=True, timeout=60)
        r.raise_for_status()
        with open(dest, "wb") as f:
            for chunk in r.iter_content(chunk_size=8192):
                f.write(chunk)
        return f"已下載：{dest}"


def execute_send_email(to, subject, body):
    import smtplib
    from email.mime.text import MIMEText
    from email.mime.multipart import MIMEMultipart
    smtp_user = os.getenv("SMTP_USER", "")
    smtp_pass = os.getenv("SMTP_PASS", "")
    if not smtp_user:
        return "未設定 SMTP_USER / SMTP_PASS，請加入 .env"
    msg = MIMEMultipart()
    msg["From"] = smtp_user
    msg["To"] = to
    msg["Subject"] = subject
    msg.attach(MIMEText(body, "plain", "utf-8"))
    with smtplib.SMTP(os.getenv("SMTP_HOST","smtp.gmail.com"), int(os.getenv("SMTP_PORT","587"))) as s:
        s.starttls()
        s.login(smtp_user, smtp_pass)
        s.send_message(msg)
    return f"Email 已寄出到 {to}"

def execute_screen_vision(question: str = "請描述這個畫面上有什麼，以及目前電腦在做什麼事。") -> tuple:
    """截圖並用 Claude vision 分析，回傳 (文字分析, 截圖bytes)"""
    import base64
    img = pyautogui.screenshot()
    buf = io.BytesIO()
    img.save(buf, format="PNG")
    img_bytes = buf.getvalue()
    img_b64 = base64.standard_b64encode(img_bytes).decode("utf-8")
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
    return response.content[0].text, img_bytes

def execute_find_image(template_path: str, confidence: float = 0.8) -> str:
    try:
        location = pyautogui.locateOnScreen(template_path, confidence=confidence)
        if location:
            cx, cy = pyautogui.center(location)
            return f"找到圖片，中心座標：({cx}, {cy})，區域：{location}"
        return "畫面上找不到該圖片"
    except Exception as e:
        return f"搜尋失敗：{str(e)}"

def execute_browser_control(action: str, url: str = "", selector: str = "", text: str = "") -> str:
    try:
        from playwright.sync_api import sync_playwright
        if action == "open":
            if _browser_ctx.get("page"):
                _browser_ctx["browser"].close()
                _browser_ctx["pw"].stop()
                _browser_ctx.clear()
            pw = sync_playwright().start()
            b = pw.chromium.launch(headless=True)
            page = b.new_page()
            page.goto(url or "https://www.google.com")
            _browser_ctx.update({"pw": pw, "browser": b, "page": page})
            return f"已開啟：{url}"
        page = _browser_ctx.get("page")
        if not page:
            return "瀏覽器未開啟，請先使用 open 開啟網頁"
        if action == "goto":
            page.goto(url)
            return f"已前往：{url}"
        elif action == "click":
            page.click(selector)
            return f"已點擊：{selector}"
        elif action == "type":
            page.fill(selector, text)
            return f"已輸入到 {selector}：{text}"
        elif action == "get_text":
            return page.inner_text(selector or "body")[:2000]
        elif action == "screenshot":
            buf = io.BytesIO()
            img_bytes = page.screenshot()
            return f"__BROWSER_SCREENSHOT__:{img_bytes.hex()}"
        elif action == "close":
            _browser_ctx["browser"].close()
            _browser_ctx["pw"].stop()
            _browser_ctx.clear()
            return "瀏覽器已關閉"
        return f"未知動作：{action}"
    except Exception as e:
        return f"執行失敗：{str(e)}"


SCHTASKS = "C:\\Windows\\System32\\schtasks.exe"
BOT_SCRIPT = r"C:\Users\blue_\claude-telegram-bot\bot.py"

def execute_manage_schedule(action: str, name: str = "", time: str = "", script: str = "") -> str:
    try:
        if action == "list":
            result = subprocess.run(
                [SCHTASKS, "/Query", "/FO", "CSV", "/NH"],
                capture_output=True, text=True, encoding="cp950", errors="replace"
            )
            lines = [l for l in result.stdout.strip().splitlines() if l.strip()]
            out = "目前排程任務：\n"
            for line in lines:
                parts = line.strip('"').split('","')
                if len(parts) >= 3:
                    t_name = parts[0].replace("\\", "").strip()
                    next_run = parts[1].strip()
                    status = parts[2].strip()
                    out += f"• {t_name} | 下次執行：{next_run} | {status}\n"
            return out.strip()

        elif action == "add":
            subprocess.run([SCHTASKS, "/Create", "/TN", name,
                            "/TR", f"pythonw {script}",
                            "/SC", "DAILY", "/ST", time, "/F"],
                           capture_output=True)
            ps = (
                f"$t = Get-ScheduledTask -TaskName '{name}';"
                f"$t.Settings.WakeToRun = $true;"
                f"$t.Settings.DisallowStartIfOnBatteries = $false;"
                f"$t.Settings.StopIfGoingOnBatteries = $false;"
                f"Set-ScheduledTask -TaskName '{name}' -Settings $t.Settings | Out-Null"
            )
            subprocess.run(["powershell.exe", "-Command", ps], capture_output=True)
            return f"排程 [{name}] 已建立，每天 {time} 執行，電腦待機也會自動喚醒。"

        elif action == "delete":
            result = subprocess.run([SCHTASKS, "/Delete", "/TN", name, "/F"],
                                    capture_output=True, text=True, encoding="cp950", errors="replace")
            if result.returncode == 0:
                return f"排程 [{name}] 已刪除。"
            else:
                return f"刪除失敗：{result.stderr.strip()}"

        elif action == "bot_status":
            result = subprocess.run(
                ["powershell.exe", "-Command",
                 "Get-Process pythonw -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count"],
                capture_output=True, text=True
            )
            count = int(result.stdout.strip()) if result.stdout.strip().isdigit() else 0
            return f"Bot 執行中（{count} 個程序）✅" if count > 0 else "Bot 未執行 ❌"

        elif action == "bot_restart":
            subprocess.run(["powershell.exe", "-Command",
                            "Stop-Process -Name pythonw -Force -ErrorAction SilentlyContinue"])
            import time as t
            t.sleep(1)
            subprocess.Popen(["pythonw", BOT_SCRIPT], cwd=str(Path(BOT_SCRIPT).parent))
            return "Bot 已重啟 ✅"

        else:
            return f"未知動作：{action}"

    except Exception as e:
        return f"執行失敗：{str(e)}"


def fetch_image(prompt: str, width: int = 512, height: int = 512):
    hf_token = os.getenv("HF_TOKEN")
    headers = {"Authorization": f"Bearer {hf_token}"}
    payload = {"inputs": prompt}
    models = [
        "black-forest-labs/FLUX.1-schnell",
        "stabilityai/stable-diffusion-xl-base-1.0",
    ]
    for model in models:
        for attempt in range(2):
            try:
                res = requests.post(
                    f"https://router.huggingface.co/hf-inference/models/{model}",
                    headers=headers,
                    json=payload,
                    timeout=120
                )
                if res.status_code == 200 and res.headers.get("content-type", "").startswith("image"):
                    return res.content
                if res.status_code == 503:
                    time.sleep(10)
                    continue
            except Exception:
                pass
    return None


async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("你好！我是小牛馬，有什麼可以幫你的？（我還記得我們之前聊過的事 😄）")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        chat_id = update.effective_chat.id
        user_text = update.message.text

        is_group = update.effective_chat.type in ("group", "supergroup")
        sender_name = update.effective_user.first_name or str(update.effective_user.id)

        # 群組內只回應有 @ 提及 bot 的訊息
        if is_group:
            bot_username = context.bot.username
            if not (update.message.entities and any(
                e.type == "mention" and user_text[e.offset:e.offset + e.length] == f"@{bot_username}"
                for e in update.message.entities
            )):
                return
            user_text = user_text.replace(f"@{bot_username}", "").strip()

        log_message(">>", sender_name, chat_id, user_text)

        # 群組訊息加上發話者名字，讓 Claude 分清楚誰在說話
        saved_text = f"[{sender_name}]: {user_text}" if is_group else user_text
        save_message(chat_id, "user", saved_text)
        history = load_history(chat_id)[-40:]
        await context.bot.send_chat_action(chat_id=chat_id, action="typing")

        base_system = SYSTEM_PROMPT_OWNER if update.effective_user.id == OWNER_ID else SYSTEM_PROMPT_DEFAULT

        # 注入長期記憶
        memories = load_long_term_memory(chat_id)
        if memories:
            mem_text = "\n".join(f"- [{m['id']}] {m['content']}" for m in memories)
            system = base_system + f"\n\n【長期記憶】以下是你記住的重要資訊，回覆時請參考：\n{mem_text}"
        else:
            system = base_system

        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            system=system,
            tools=TOOLS,
            messages=history
        )

        # 處理工具呼叫
        if response.stop_reason == "tool_use":
            tool_use = next(b for b in response.content if b.type == "tool_use")

            simple_tools = {
                "stt": lambda: execute_stt(),
                "ocr": lambda: execute_ocr(tool_use.input.get("image_path","")),
                "workflow": lambda: execute_workflow(
                    tool_use.input["action"],
                    tool_use.input.get("name",""),
                    tool_use.input.get("steps","")),
                "screen_watch": lambda: execute_screen_watch(
                    tool_use.input["template_path"],
                    tool_use.input["command"],
                    tool_use.input.get("timeout", 60)),
                "file_transfer": lambda: execute_file_transfer(
                    tool_use.input["action"],
                    tool_use.input["source"],
                    tool_use.input.get("dest","")),
                "window_control": lambda: execute_window_control(
                    tool_use.input["action"], tool_use.input.get("keyword","")),
                "hotkey": lambda: execute_hotkey(tool_use.input["keys"]),
                "clipboard": lambda: execute_clipboard(
                    tool_use.input["action"], tool_use.input.get("text","")),
                "file_system": lambda: execute_file_system(
                    tool_use.input["action"],
                    tool_use.input.get("path",""),
                    tool_use.input.get("dest",""),
                    tool_use.input.get("content",""),
                    tool_use.input.get("keyword","")),
                "system_monitor": lambda: execute_system_monitor(
                    tool_use.input["action"], tool_use.input.get("target","")),
                "notify": lambda: execute_notify(
                    tool_use.input["title"], tool_use.input["message"]),
                "tts": lambda: execute_tts(tool_use.input["text"]),
                "send_email": lambda: execute_send_email(
                    tool_use.input["to"], tool_use.input["subject"], tool_use.input["body"]),
            }

            if tool_use.name in simple_tools:
                import asyncio
                loop = asyncio.get_running_loop()
                tool_result = await loop.run_in_executor(None, simple_tools[tool_use.name])
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=512,
                    system=system,
                    tools=TOOLS,
                    messages=history + [
                        {"role": "assistant", "content": response.content},
                        {"role": "user", "content": [{"type": "tool_result", "tool_use_id": tool_use.id, "content": tool_result}]}
                    ]
                )
                text_blocks = [b.text for b in response.content if hasattr(b, "text")]
                reply = text_blocks[0] if text_blocks else tool_result
                save_message(chat_id, "assistant", reply)
                await update.message.reply_text(reply)
                return

            elif tool_use.name == "screen_vision":
                question = tool_use.input.get("question", "請描述這個畫面上有什麼，以及目前電腦在做什麼事。")
                import asyncio
                loop = asyncio.get_running_loop()
                analysis, img_bytes = await loop.run_in_executor(None, execute_screen_vision, question)
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1024,
                    system=system,
                    tools=TOOLS,
                    messages=history + [
                        {"role": "assistant", "content": response.content},
                        {"role": "user", "content": [{"type": "tool_result", "tool_use_id": tool_use.id, "content": analysis}]}
                    ]
                )
                await update.message.reply_photo(photo=img_bytes, caption="📸 螢幕截圖")
                text_blocks = [b.text for b in response.content if hasattr(b, "text")]
                reply = text_blocks[0] if text_blocks else analysis
                save_message(chat_id, "assistant", reply)
                await update.message.reply_text(reply)
                return

            elif tool_use.name == "find_image_on_screen":
                import asyncio
                loop = asyncio.get_running_loop()
                result = await loop.run_in_executor(
                    None, execute_find_image,
                    tool_use.input["template_path"],
                    tool_use.input.get("confidence", 0.8)
                )
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=512,
                    system=system,
                    tools=TOOLS,
                    messages=history + [
                        {"role": "assistant", "content": response.content},
                        {"role": "user", "content": [{"type": "tool_result", "tool_use_id": tool_use.id, "content": result}]}
                    ]
                )
                text_blocks = [b.text for b in response.content if hasattr(b, "text")]
                reply = text_blocks[0] if text_blocks else result
                save_message(chat_id, "assistant", reply)
                await update.message.reply_text(reply)
                return

            elif tool_use.name == "browser_control":
                inp = tool_use.input
                import asyncio
                loop = asyncio.get_running_loop()
                result = await loop.run_in_executor(
                    None, lambda: execute_browser_control(
                        action=inp["action"],
                        url=inp.get("url", ""),
                        selector=inp.get("selector", ""),
                        text=inp.get("text", "")
                    )
                )
                if result.startswith("__BROWSER_SCREENSHOT__:"):
                    img_bytes = bytes.fromhex(result.split(":", 1)[1])
                    await update.message.reply_photo(photo=img_bytes, caption="🌐 瀏覽器截圖")
                    tool_result_text = "已截圖並傳送"
                else:
                    tool_result_text = result
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=512,
                    system=system,
                    tools=TOOLS,
                    messages=history + [
                        {"role": "assistant", "content": response.content},
                        {"role": "user", "content": [{"type": "tool_result", "tool_use_id": tool_use.id, "content": tool_result_text}]}
                    ]
                )
                text_blocks = [b.text for b in response.content if hasattr(b, "text")]
                reply = text_blocks[0] if text_blocks else tool_result_text
                save_message(chat_id, "assistant", reply)
                await update.message.reply_text(reply)
                return

            elif tool_use.name == "long_term_memory":
                inp = tool_use.input
                action = inp["action"]
                if action == "save":
                    save_long_term_memory(chat_id, inp.get("content", ""))
                    tool_result = f"已記住：{inp.get('content', '')}"
                elif action == "list":
                    mems = load_long_term_memory(chat_id)
                    if mems:
                        tool_result = "長期記憶列表：\n" + "\n".join(f"[{m['id']}] {m['content']}" for m in mems)
                    else:
                        tool_result = "目前沒有長期記憶。"
                elif action == "delete":
                    delete_long_term_memory(inp.get("memory_id", 0))
                    tool_result = f"已刪除記憶 ID {inp.get('memory_id')}"
                else:
                    tool_result = "未知動作"
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=512,
                    system=system,
                    tools=TOOLS,
                    messages=history + [
                        {"role": "assistant", "content": response.content},
                        {"role": "user", "content": [{"type": "tool_result", "tool_use_id": tool_use.id, "content": tool_result}]}
                    ]
                )
                text_blocks = [b.text for b in response.content if hasattr(b, "text")]
                reply = text_blocks[0] if text_blocks else tool_result
                save_message(chat_id, "assistant", reply)
                await update.message.reply_text(reply)
                return

            elif tool_use.name == "get_stock":
                import asyncio
                loop = asyncio.get_running_loop()
                tool_result = await loop.run_in_executor(
                    None, fetch_stock,
                    tool_use.input["symbol"],
                    tool_use.input.get("period", "1mo")
                )
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1024,
                    system=system,
                    tools=TOOLS,
                    messages=history + [
                        {"role": "assistant", "content": response.content},
                        {"role": "user", "content": [{"type": "tool_result", "tool_use_id": tool_use.id, "content": tool_result}]}
                    ]
                )

            elif tool_use.name == "get_weather":
                tool_result = fetch_weather(tool_use.input["city"])
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=1024,
                    system=system,
                    tools=TOOLS,
                    messages=history + [
                        {"role": "assistant", "content": response.content},
                        {"role": "user", "content": [{"type": "tool_result", "tool_use_id": tool_use.id, "content": tool_result}]}
                    ]
                )

            elif tool_use.name == "desktop_control":
                inp = tool_use.input
                import asyncio
                loop = asyncio.get_running_loop()
                result = await loop.run_in_executor(
                    None,
                    lambda: execute_desktop_control(
                        action=inp["action"],
                        x=inp.get("x"),
                        y=inp.get("y"),
                        text=inp.get("text"),
                        app=inp.get("app"),
                        direction=inp.get("direction", "down"),
                        amount=inp.get("amount", 3)
                    )
                )
                tool_result_content = result["message"]
                # 把截圖或結果回傳給 Claude，讓 Claude 生成回覆
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=512,
                    system=system,
                    tools=TOOLS,
                    messages=history + [
                        {"role": "assistant", "content": response.content},
                        {"role": "user", "content": [{"type": "tool_result", "tool_use_id": tool_use.id, "content": tool_result_content}]}
                    ]
                )
                # 如果有截圖，先傳圖
                if result.get("screenshot"):
                    await update.message.reply_photo(photo=result["screenshot"], caption="📸 桌面截圖")
                text_blocks = [b.text for b in response.content if hasattr(b, "text")]
                reply = text_blocks[0] if text_blocks else result["message"]
                save_message(chat_id, "assistant", reply)
                await update.message.reply_text(reply)
                return

            elif tool_use.name == "manage_schedule":
                inp = tool_use.input
                import asyncio
                loop = asyncio.get_running_loop()
                tool_result = await loop.run_in_executor(
                    None,
                    lambda: execute_manage_schedule(
                        action=inp["action"],
                        name=inp.get("name", ""),
                        time=inp.get("time", ""),
                        script=inp.get("script", "")
                    )
                )
                response = client.messages.create(
                    model="claude-sonnet-4-6",
                    max_tokens=512,
                    system=system,
                    tools=TOOLS,
                    messages=history + [
                        {"role": "assistant", "content": response.content},
                        {"role": "user", "content": [{"type": "tool_result", "tool_use_id": tool_use.id, "content": tool_result}]}
                    ]
                )
                text_blocks = [b.text for b in response.content if hasattr(b, "text")]
                reply = text_blocks[0] if text_blocks else tool_result
                save_message(chat_id, "assistant", reply)
                await update.message.reply_text(reply)
                return

            elif tool_use.name == "generate_image":
                prompt = tool_use.input["prompt"]
                width = tool_use.input.get("width", 512)
                height = tool_use.input.get("height", 512)
                overlay_text = tool_use.input.get("overlay_text", "")
                await context.bot.send_chat_action(chat_id=chat_id, action="upload_photo")
                import asyncio
                loop = asyncio.get_running_loop()
                image_bytes = await loop.run_in_executor(None, fetch_image, prompt, width, height)
                if image_bytes:
                    if overlay_text:
                        image_bytes = add_text_overlay(image_bytes, overlay_text)
                if image_bytes:
                    await update.message.reply_photo(photo=image_bytes, caption=f"🎨 {prompt}")
                else:
                    await update.message.reply_text("圖片生成失敗，請稍後再試。")
                save_message(chat_id, "assistant", f"已為用戶生成圖片：{prompt}")
                return

        text_blocks = [b.text for b in response.content if hasattr(b, "text")]
        if not text_blocks:
            await update.message.reply_text("哎，我卡住了，請再說一次。")
            return
        reply = text_blocks[0]
        save_message(chat_id, "assistant", reply)
        log_message("<<", "小牛馬", chat_id, reply)
        await update.message.reply_text(reply)

    except Exception as e:
        logging.error(f"handle_message error: {e}", exc_info=True)
        await update.message.reply_text("發生錯誤，請稍後再試。")


async def chatid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"這個聊天室的 ID 是：`{update.effective_chat.id}`", parse_mode="Markdown")

async def myid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"你的 Telegram ID 是：`{update.effective_user.id}`", parse_mode="Markdown")


async def clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    clear_history(update.effective_chat.id)
    await update.message.reply_text("對話紀錄已清除。")


async def memories(update: Update, context: ContextTypes.DEFAULT_TYPE):
    mems = load_long_term_memory(update.effective_chat.id)
    if not mems:
        await update.message.reply_text("目前沒有長期記憶。")
        return
    text = "📝 長期記憶列表：\n\n" + "\n".join(f"[{m['id']}] {m['content']}" for m in mems)
    await update.message.reply_text(text)


async def goodmorning(context: ContextTypes.DEFAULT_TYPE):
    group_id = int(os.getenv("GROUP_CHAT_ID"))
    import asyncio
    loop = asyncio.get_running_loop()
    def generate_goodmorning():
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=150,
            system="你是小牛馬，說話嘴賤幽默。現在是早上11點，用繁體中文生成一段簡短有趣的早安問候語，每次內容都不同，可以包含今日天氣心情預告、幽默吐槽、激勵話語等，結尾加上表情符號。約30-60字。",
            messages=[{"role": "user", "content": "生成今天的早安訊息"}]
        )
        return response.content[0].text
    msg = await loop.run_in_executor(None, generate_goodmorning)
    await context.bot.send_message(chat_id=group_id, text=msg)

async def goodnight(context: ContextTypes.DEFAULT_TYPE):
    group_id = int(os.getenv("GROUP_CHAT_ID"))
    import asyncio
    loop = asyncio.get_running_loop()
    def generate_goodnight():
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=150,
            system="你是小牛馬，說話嘴賤幽默。現在是晚上10:30，你要下班了，用繁體中文生成一段簡短有趣的晚安問候語，每次內容都不同，可以包含當天心情、幽默吐槽、祝福語等，結尾加上表情符號。約30-60字。",
            messages=[{"role": "user", "content": "生成今晚的晚安訊息"}]
        )
        return response.content[0].text
    msg = await loop.run_in_executor(None, generate_goodnight)
    await context.bot.send_message(chat_id=group_id, text=msg)

_learned_today: str = ""  # 防止同一天重複執行

async def daily_learn_and_push(context: ContextTypes.DEFAULT_TYPE):
    """每天晚上 9 點：讓 Claude 學習一個新技能並記錄，然後上傳 GitHub"""
    global _learned_today
    import asyncio
    loop = asyncio.get_running_loop()

    today = datetime.date.today()
    today_str = str(today)

    # 同一天已執行過則跳過（防止重啟後重複觸發）
    if _learned_today == today_str:
        return
    _learned_today = today_str

    topics = [
        "Python asyncio 進階用法與最佳實踐",
        "Telegram Bot API 進階功能與限制",
        "Claude API 工具使用（Tool Use）技巧",
        "SQLite 效能優化技巧",
        "Python 錯誤處理與 logging 最佳實踐",
        "圖片處理：Pillow 進階技巧",
        "HTTP requests 效能優化與重試機制",
        "Python 排程任務：APScheduler 進階用法",
        "Git 自動化流程與 GitHub Actions 入門",
        "pyautogui 桌面自動化進階技巧",
    ]

    topic = topics[today.toordinal() % len(topics)]

    def do_learn():
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1500,
            system="你是小牛馬，一個積極自學的 AI 助理。每天晚上你會主動學習一個新技能並做成筆記。",
            messages=[{"role": "user", "content": f"今天的學習主題是：{topic}。請整理出重點知識、實用範例程式碼（如果適用）、以及你學到的心得，用繁體中文條列式整理，格式清晰。"}]
        )
        return response.content[0].text

    learn_content = await loop.run_in_executor(None, do_learn)

    # 儲存學習筆記
    log_path = Path(__file__).parent / "learning_log.md"
    existing = log_path.read_text(encoding="utf-8") if log_path.exists() else ""
    entry = f"\n\n## {today_str} — {topic}\n\n{learn_content}\n\n---"
    log_path.write_text(existing + entry, encoding="utf-8")

    # 上傳 GitHub
    def do_git_push():
        repo = str(Path(__file__).parent)
        subprocess.run(["git", "-C", repo, "add", "learning_log.md"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        subprocess.run(["git", "-C", repo, "commit", "-m", f"每日學習筆記：{today_str} {topic}"],
                       stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        result = subprocess.run(["git", "-C", repo, "push"],
                                stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True, encoding="utf-8")
        # exit 0 = 成功；"Everything up-to-date" 也算成功
        return result.returncode == 0

    push_ok = await loop.run_in_executor(None, do_git_push)

    # 通知于晏哥
    status = "✅ 已上傳 GitHub" if push_ok else "⚠️ GitHub 上傳失敗"
    msg = f"📚 今日學習完成！\n主題：{topic}\n{status}"
    await context.bot.send_message(chat_id=OWNER_ID, text=msg)


if __name__ == "__main__":
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    app = ApplicationBuilder().token(token).build()

    # 每天早上 11:00 台灣時間（UTC+8）= 03:00 UTC
    app.job_queue.run_daily(
        goodmorning,
        time=datetime.time(hour=3, minute=0, tzinfo=datetime.timezone.utc)
    )
    # 每天晚上 10:30 台灣時間（UTC+8）= 14:30 UTC
    app.job_queue.run_daily(
        goodnight,
        time=datetime.time(hour=14, minute=30, tzinfo=datetime.timezone.utc)
    )
    # 每天晚上 9:00 台灣時間（UTC+8）= 13:00 UTC
    app.job_queue.run_daily(
        daily_learn_and_push,
        time=datetime.time(hour=13, minute=0, tzinfo=datetime.timezone.utc)
    )

    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("chatid", chatid))
    app.add_handler(CommandHandler("myid", myid))
    app.add_handler(CommandHandler("clear", clear))
    app.add_handler(CommandHandler("memories", memories))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Bot 已啟動...")
    app.run_polling()
