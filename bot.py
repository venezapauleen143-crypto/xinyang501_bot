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
        "name": "ai_plan",
        "description": "AI 自動規劃並執行多步驟任務。用戶給一個目標，自動拆解成步驟並執行。",
        "input_schema": {
            "type": "object",
            "properties": {
                "goal": {"type": "string", "description": "要達成的目標"}
            },
            "required": ["goal"]
        }
    },
    {
        "name": "screen_stream",
        "description": "螢幕串流，持續截圖並傳送給用戶。",
        "input_schema": {
            "type": "object",
            "properties": {
                "duration": {"type": "integer", "description": "持續秒數，預設 10"},
                "interval": {"type": "number", "description": "截圖間隔秒數，預設 2"}
            },
            "required": []
        }
    },
    {
        "name": "drag",
        "description": "拖曳滑鼠從一個位置到另一個位置。",
        "input_schema": {
            "type": "object",
            "properties": {
                "x1": {"type": "integer"}, "y1": {"type": "integer"},
                "x2": {"type": "integer"}, "y2": {"type": "integer"},
                "duration": {"type": "number", "description": "拖曳時間，預設 0.5 秒"}
            },
            "required": ["x1","y1","x2","y2"]
        }
    },
    {
        "name": "power_control",
        "description": "電源管理：睡眠、重開機、關機。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["sleep","restart","shutdown"]}
            },
            "required": ["action"]
        }
    },
    {
        "name": "virtual_desktop",
        "description": "切換或建立 Windows 虛擬桌面。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["left","right","new"]}
            },
            "required": ["action"]
        }
    },
    {
        "name": "bluetooth",
        "description": "掃描或連線藍牙裝置。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["scan","connect"]},
                "mac": {"type": "string", "description": "藍牙 MAC 位址（connect 時使用）"}
            },
            "required": ["action"]
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
    },
    {
        "name": "reminder",
        "description": "設定一次性提醒/鬧鐘，到時發出通知和語音。",
        "input_schema": {
            "type": "object",
            "properties": {
                "time": {"type": "string", "description": "HH:MM 格式時間，或純數字代表幾秒後"},
                "message": {"type": "string"}
            },
            "required": ["time", "message"]
        }
    },
    {
        "name": "webpage_shot",
        "description": "對指定網址截取整頁截圖，或監控網頁內容變化。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["screenshot", "monitor"]},
                "url": {"type": "string"},
                "selector": {"type": "string", "description": "監控的 CSS 選擇器（monitor 使用）"},
                "interval": {"type": "number", "description": "監控間隔秒數（monitor 使用）"},
                "duration": {"type": "number", "description": "監控總秒數（monitor 使用）"}
            },
            "required": ["action", "url"]
        }
    },
    {
        "name": "file_tools",
        "description": "批次重新命名檔案、同步資料夾。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["batch_rename", "sync"]},
                "path": {"type": "string", "description": "資料夾路徑"},
                "dest": {"type": "string", "description": "目標資料夾（sync 使用）"},
                "pattern": {"type": "string", "description": "正規表達式（batch_rename 使用）"},
                "replacement": {"type": "string", "description": "替換字串"},
                "ext": {"type": "string", "description": "副檔名過濾（batch_rename 使用）"}
            },
            "required": ["action", "path"]
        }
    },
    {
        "name": "image_tools",
        "description": "壓縮圖片、批次處理圖片資料夾、OCR+翻譯。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["compress", "batch", "ocr_translate"]},
                "path": {"type": "string"},
                "quality": {"type": "integer", "description": "壓縮品質 0-100（compress/batch 使用）"},
                "width": {"type": "integer"}, "height": {"type": "integer"},
                "target_lang": {"type": "string", "description": "翻譯目標語言（ocr_translate 使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "lookup",
        "description": "查詢 IP 地理位置或外幣匯率。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["ip", "currency"]},
                "ip": {"type": "string", "description": "IP 位址（ip 使用，留空查自己）"},
                "amount": {"type": "number", "description": "金額（currency 使用）"},
                "from_cur": {"type": "string", "description": "來源幣別，如 USD"},
                "to_cur": {"type": "string", "description": "目標幣別，如 TWD"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "system_tools",
        "description": "Windows 事件日誌、USB 裝置、防火牆、印表機、網路芳鄰、字型列表。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["event_log","usb_list","firewall_list","firewall_add","firewall_remove","printer_list","printer_jobs","net_share_list","net_share_connect","net_share_disconnect","font_list","rdp"]},
                "name": {"type": "string"},
                "host": {"type": "string"},
                "port": {"type": "integer"},
                "direction": {"type": "string", "enum": ["in","out"]},
                "share_path": {"type": "string"},
                "drive": {"type": "string"},
                "user": {"type": "string"},
                "password": {"type": "string"},
                "keyword": {"type": "string"},
                "log_name": {"type": "string"},
                "level": {"type": "string"},
                "count": {"type": "integer"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "tts_advanced",
        "description": "使用微軟 Edge TTS 合成更自然的語音，支援多種中文聲音。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["speak","list_voices"]},
                "text": {"type": "string"},
                "voice": {"type": "string", "description": "語音名稱，如 zh-TW-YunJheNeural（女）或 zh-TW-YunJheNeural（男）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "send_voice",
        "description": "將文字轉成語音並以 Telegram 語音訊息傳送給用戶。當用戶說「用語音說」、「語音回答」、「唸給我聽」、「說出來」等時使用此工具。",
        "input_schema": {
            "type": "object",
            "properties": {
                "text": {"type": "string", "description": "要轉換成語音的文字內容"},
                "voice": {"type": "string", "description": "語音名稱，預設 zh-TW-YunJheNeural（女聲）。男聲用 zh-TW-YunJheNeural"}
            },
            "required": ["text"]
        }
    },
    {
        "name": "todo_list",
        "description": "管理本地任務清單：新增、列出、完成、刪除任務。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["add","list","done","delete","clear"]},
                "task": {"type": "string", "description": "任務內容（add 使用）"},
                "id": {"type": "integer", "description": "任務 ID（done/delete 使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "sysres_chart",
        "description": "即時監控 CPU/RAM 使用率並生成圖表。",
        "input_schema": {
            "type": "object",
            "properties": {
                "duration": {"type": "integer", "description": "監控秒數，預設 10"}
            },
            "required": []
        }
    },
    {
        "name": "password_mgr",
        "description": "加密儲存或查詢帳號密碼（本地加密，需主密碼）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["save","get"]},
                "site": {"type": "string"},
                "username": {"type": "string", "description": "帳號（save 使用）"},
                "password": {"type": "string", "description": "密碼（save 使用）"},
                "master": {"type": "string", "description": "主密碼"}
            },
            "required": ["action", "site", "master"]
        }
    },
    {
        "name": "clipboard_image",
        "description": "讀取或寫入剪貼簿中的圖片。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["get","set"]},
                "path": {"type": "string", "description": "圖片路徑（set 使用）或儲存路徑（get 使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "email_control",
        "description": "讀取 IMAP 郵件收件匣。",
        "input_schema": {
            "type": "object",
            "properties": {
                "host": {"type": "string", "description": "IMAP 伺服器，如 imap.gmail.com"},
                "user": {"type": "string"},
                "password": {"type": "string"},
                "folder": {"type": "string", "description": "資料夾名稱，預設 INBOX"},
                "count": {"type": "integer", "description": "讀取封數，預設 5"}
            },
            "required": ["host", "user", "password"]
        }
    },
    {
        "name": "calendar",
        "description": "查詢或新增 Google Calendar 行事曆事件。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["list","add"]},
                "days": {"type": "integer", "description": "往後幾天（list 使用），預設 7"},
                "title": {"type": "string", "description": "行程標題（add 使用）"},
                "start": {"type": "string", "description": "開始時間，格式 2026-04-13T10:00:00（add 使用）"},
                "end": {"type": "string", "description": "結束時間（add 使用）"},
                "description": {"type": "string", "description": "說明（add 使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "global_hotkey",
        "description": "監聽全域快捷鍵，觸發時執行指定指令。",
        "input_schema": {
            "type": "object",
            "properties": {
                "hotkey": {"type": "string", "description": "快捷鍵組合，如 ctrl+shift+a"},
                "command": {"type": "string", "description": "觸發時執行的 shell 指令"},
                "duration": {"type": "number", "description": "監聽秒數，預設 60"}
            },
            "required": ["hotkey", "command"]
        }
    },
    {
        "name": "git",
        "description": "執行 Git 操作：狀態查看、提交、推送、拉取等。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["status","log","pull","add","commit","push","diff"]},
                "repo": {"type": "string", "description": "repo 路徑，預設當前目錄"},
                "message": {"type": "string", "description": "commit 訊息（commit 使用）"},
                "branch": {"type": "string", "description": "分支名稱，預設 master"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "hardware",
        "description": "查看進階硬體資訊：GPU 使用率/溫度、電池狀態、CPU 溫度。",
        "input_schema": {"type": "object", "properties": {}, "required": []}
    },
    {
        "name": "report",
        "description": "根據提供的資料生成 HTML 格式報告並存檔。",
        "input_schema": {
            "type": "object",
            "properties": {
                "title": {"type": "string"},
                "data": {"type": "string", "description": "JSON 格式資料，如 {\"section\": [{\"col1\": val}]}"},
                "output": {"type": "string", "description": "輸出路徑（選填）"}
            },
            "required": ["title", "data"]
        }
    },
    {
        "name": "dropbox",
        "description": "上傳或下載 Dropbox 雲端檔案。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["upload","download"]},
                "local": {"type": "string", "description": "本地路徑"},
                "remote": {"type": "string", "description": "Dropbox 路徑，如 /folder/file.txt"},
                "token": {"type": "string", "description": "Dropbox access token（選填，可用環境變數）"}
            },
            "required": ["action", "local", "remote"]
        }
    },
    {
        "name": "docker",
        "description": "管理 Docker 容器：列出、啟動、停止、查看日誌。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["list","start","stop","logs","images"]},
                "name": {"type": "string", "description": "容器名稱（start/stop/logs 使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "pdf_image",
        "description": "將 PDF 每頁轉成圖片（PNG）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "output_dir": {"type": "string", "description": "輸出資料夾（選填）"},
                "dpi": {"type": "integer", "description": "解析度，預設 150"}
            },
            "required": ["path"]
        }
    },
    {
        "name": "barcode",
        "description": "掃描圖片或截圖中的條碼、QR Code。",
        "input_schema": {
            "type": "object",
            "properties": {
                "image_path": {"type": "string", "description": "圖片路徑（選填，不填則截圖）"}
            },
            "required": []
        }
    },
    {
        "name": "nlp",
        "description": "AI 文字分析：摘要長文或分析情緒。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["summarize","sentiment"]},
                "text": {"type": "string"}
            },
            "required": ["action", "text"]
        }
    },
    {
        "name": "vpn",
        "description": "VPN 連線管理：列出、連線、斷線。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["list","connect","disconnect"]},
                "name": {"type": "string", "description": "VPN 名稱"},
                "user": {"type": "string"},
                "password": {"type": "string"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "restore_point",
        "description": "建立或列出 Windows 系統還原點。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["create","list"]},
                "description": {"type": "string", "description": "還原點說明（create 使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "disk_analyze",
        "description": "分析磁碟空間使用，列出佔用最多的資料夾。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string", "description": "分析路徑，預設 C:/"},
                "top": {"type": "integer", "description": "顯示前幾名，預設 10"}
            },
            "required": []
        }
    },
    {
        "name": "face_detect",
        "description": "從截圖或圖片中偵測人臉，並標記位置。",
        "input_schema": {
            "type": "object",
            "properties": {
                "image_path": {"type": "string", "description": "圖片路徑（選填，不填則截圖）"},
                "output": {"type": "string", "description": "輸出路徑（選填）"}
            },
            "required": []
        }
    },
    {
        "name": "video_gif",
        "description": "將影片片段轉成 GIF 動圖。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "start": {"type": "number", "description": "起始秒，預設 0"},
                "duration": {"type": "number", "description": "持續秒數，預設 5"},
                "output": {"type": "string", "description": "輸出路徑（選填）"},
                "fps": {"type": "integer", "description": "GIF fps，預設 10"}
            },
            "required": ["path"]
        }
    },
    {
        "name": "excel_chart",
        "description": "在 Excel 工作表中生成圖表（長條/折線/圓餅）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "path": {"type": "string"},
                "sheet": {"type": "string"},
                "type": {"type": "string", "enum": ["bar","line","pie"], "description": "圖表類型，預設 bar"},
                "title": {"type": "string"}
            },
            "required": ["path", "sheet"]
        }
    },
    {
        "name": "speedtest",
        "description": "測試目前網路的上下載速度和 Ping 值。",
        "input_schema": {"type": "object", "properties": {}, "required": []}
    },
    {
        "name": "screenshot_compare",
        "description": "比對兩張截圖或圖片的差異，標記變化區域。",
        "input_schema": {
            "type": "object",
            "properties": {
                "img1": {"type": "string", "description": "第一張圖片路徑（選填，不填則即時截圖）"},
                "img2": {"type": "string", "description": "第二張圖片路徑（選填，不填則 2 秒後再截圖）"},
                "output": {"type": "string", "description": "輸出路徑（選填）"}
            },
            "required": []
        }
    },
    {
        "name": "screen_record",
        "description": "螢幕錄影或攝影機拍照。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["record","webcam"]},
                "duration": {"type": "number", "description": "錄影秒數（record 使用）"},
                "output": {"type": "string", "description": "輸出檔案路徑（選填）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "translate",
        "description": "翻譯文字，支援中英日韓等多語言。",
        "input_schema": {
            "type": "object",
            "properties": {
                "text": {"type": "string"},
                "target": {"type": "string", "description": "目標語言代碼，如 zh-TW, en, ja, ko，預設 zh-TW"},
                "source": {"type": "string", "description": "來源語言，預設 auto 自動偵測"}
            },
            "required": ["text"]
        }
    },
    {
        "name": "chart",
        "description": "根據數據生成折線圖、長條圖、圓餅圖並存成圖片。",
        "input_schema": {
            "type": "object",
            "properties": {
                "type": {"type": "string", "enum": ["line","bar","pie"]},
                "data": {"type": "string", "description": "JSON 格式資料，如 {\"A\": [1,2,3]} 或 {\"A\": 30, \"B\": 70}"},
                "title": {"type": "string", "description": "圖表標題（選填）"},
                "output": {"type": "string", "description": "輸出路徑（選填）"}
            },
            "required": ["type", "data"]
        }
    },
    {
        "name": "pptx_control",
        "description": "讀取或建立 PowerPoint 簡報。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["read","create"]},
                "path": {"type": "string"},
                "slides": {"type": "string", "description": "JSON 格式投影片，如 [{\"title\":\"標題\",\"body\":\"內容\"}]（create 使用）"}
            },
            "required": ["action", "path"]
        }
    },
    {
        "name": "api_call",
        "description": "呼叫任意外部 HTTP REST API。",
        "input_schema": {
            "type": "object",
            "properties": {
                "method": {"type": "string", "enum": ["GET","POST","PUT","DELETE","PATCH"]},
                "url": {"type": "string"},
                "headers": {"type": "string", "description": "JSON 格式 headers（選填）"},
                "body": {"type": "string", "description": "JSON 格式 body（選填）"}
            },
            "required": ["method", "url"]
        }
    },
    {
        "name": "watchdog",
        "description": "守護指定程序，若崩潰則自動重啟。",
        "input_schema": {
            "type": "object",
            "properties": {
                "process": {"type": "string", "description": "程序名稱，如 pythonw.exe"},
                "script": {"type": "string", "description": "崩潰時要執行的腳本路徑"},
                "duration": {"type": "number", "description": "守護秒數，預設 60"}
            },
            "required": ["process", "script"]
        }
    },
    {
        "name": "ssh_sftp",
        "description": "SSH 遠端執行指令，或 SFTP 上傳/下載檔案。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["ssh_run","sftp_upload","sftp_download"]},
                "host": {"type": "string"},
                "user": {"type": "string"},
                "password": {"type": "string"},
                "command": {"type": "string", "description": "SSH 指令（ssh_run 使用）"},
                "local": {"type": "string", "description": "本地路徑"},
                "remote": {"type": "string", "description": "遠端路徑"}
            },
            "required": ["action", "host", "user", "password"]
        }
    },
    {
        "name": "network_diag",
        "description": "網路診斷：Ping、路由追蹤、Port 掃描。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["ping","traceroute","portscan"]},
                "host": {"type": "string"},
                "ports": {"type": "string", "description": "要掃描的 port 列表，如 22,80,443（portscan 使用）"}
            },
            "required": ["action", "host"]
        }
    },
    {
        "name": "win_service",
        "description": "管理 Windows 系統服務：列出、啟動、停止。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["list","start","stop"]},
                "name": {"type": "string", "description": "服務名稱（start/stop 使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "pdf_edit",
        "description": "PDF 進階編輯：合併、分割、加浮水印。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["merge","split","watermark"]},
                "path": {"type": "string", "description": "PDF 路徑"},
                "output": {"type": "string", "description": "輸出路徑"},
                "paths": {"type": "string", "description": "JSON 格式多個 PDF 路徑（merge 使用）"},
                "text": {"type": "string", "description": "浮水印文字（watermark 使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "audio_process",
        "description": "音訊處理：格式轉換、剪輯。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["convert","trim"]},
                "input": {"type": "string", "description": "輸入檔案路徑"},
                "output": {"type": "string", "description": "輸出檔案路徑"},
                "start_ms": {"type": "integer", "description": "起始毫秒（trim 使用）"},
                "end_ms": {"type": "integer", "description": "結束毫秒（trim 使用）"}
            },
            "required": ["action", "input"]
        }
    },
    {
        "name": "push_notify",
        "description": "發送推播通知到 Discord 或 LINE。",
        "input_schema": {
            "type": "object",
            "properties": {
                "platform": {"type": "string", "enum": ["discord","line"]},
                "message": {"type": "string"},
                "webhook_or_token": {"type": "string", "description": "Discord Webhook URL 或 LINE Notify Token"}
            },
            "required": ["platform", "message", "webhook_or_token"]
        }
    },
    {
        "name": "disk_backup",
        "description": "磁碟暫存清理或資料夾備份。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["list_temp","clean_temp","backup"]},
                "src": {"type": "string", "description": "備份來源路徑（backup 使用）"},
                "dest": {"type": "string", "description": "備份目標資料夾（backup 使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "registry",
        "description": "讀取或寫入 Windows 登錄檔。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["read","write"]},
                "key": {"type": "string", "description": "登錄檔路徑，如 HKCU\\Software\\MyApp"},
                "value_name": {"type": "string", "description": "值名稱"},
                "value": {"type": "string", "description": "要寫入的值（write 使用）"}
            },
            "required": ["action", "key"]
        }
    },
    {
        "name": "video_process",
        "description": "影片處理：截取指定秒數畫面、剪輯片段。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["screenshot","trim"]},
                "path": {"type": "string"},
                "second": {"type": "number", "description": "截取的秒數（screenshot 使用）"},
                "start": {"type": "number", "description": "起始秒（trim 使用）"},
                "end": {"type": "number", "description": "結束秒（trim 使用）"},
                "output": {"type": "string", "description": "輸出路徑（選填）"}
            },
            "required": ["action", "path"]
        }
    },
    {
        "name": "monitor_config",
        "description": "列出所有螢幕顯示器的資訊（解析度、位置、主副螢幕）。",
        "input_schema": {"type": "object", "properties": {}, "required": []}
    },
    {
        "name": "run_code",
        "description": "直接執行 Python 程式碼或 PowerShell 指令。",
        "input_schema": {
            "type": "object",
            "properties": {
                "type": {"type": "string", "enum": ["python", "shell"], "description": "python=執行Python, shell=執行PowerShell"},
                "code": {"type": "string", "description": "要執行的程式碼或指令"}
            },
            "required": ["type", "code"]
        }
    },
    {
        "name": "document_control",
        "description": "讀取或寫入 Word、Excel、PDF 文件。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["word_read","word_write","excel_read","excel_write","pdf_read"]},
                "path": {"type": "string", "description": "文件路徑"},
                "content": {"type": "string", "description": "要寫入的內容（word_write/excel_write 使用）"},
                "sheet": {"type": "string", "description": "Excel 工作表名稱（excel_read/excel_write 使用）"}
            },
            "required": ["action", "path"]
        }
    },
    {
        "name": "web_scrape",
        "description": "爬取網頁指定內容，或偵測螢幕區域變化。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["scrape", "screen_diff"]},
                "url": {"type": "string", "description": "網址（scrape 使用）"},
                "selector": {"type": "string", "description": "CSS 選擇器（scrape 使用）"},
                "interval": {"type": "number", "description": "偵測間隔秒數（screen_diff 使用）"},
                "region": {"type": "string", "description": "螢幕區域（screen_diff 使用）"}
            },
            "required": ["action"]
        }
    },
    {
        "name": "image_edit",
        "description": "圖片編輯：裁切、縮放、加文字、合併圖片。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["crop","resize","text","merge"]},
                "path": {"type": "string", "description": "圖片路徑"},
                "params": {"type": "string", "description": "參數（crop: x y w h; resize: w h; text: x y 文字; merge: 圖片2路徑）"}
            },
            "required": ["action", "path"]
        }
    },
    {
        "name": "cloud_storage",
        "description": "上傳或下載 Google Drive 檔案。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["upload","download"]},
                "path": {"type": "string", "description": "本地檔案路徑"},
                "drive_id": {"type": "string", "description": "Google Drive 檔案或資料夾 ID"}
            },
            "required": ["action", "path"]
        }
    },
    {
        "name": "database",
        "description": "執行 SQLite 或 MySQL 資料庫查詢。",
        "input_schema": {
            "type": "object",
            "properties": {
                "type": {"type": "string", "enum": ["sqlite","mysql"]},
                "db": {"type": "string", "description": "SQLite: 資料庫路徑；MySQL: host"},
                "name": {"type": "string", "description": "MySQL 資料庫名稱"},
                "sql": {"type": "string", "description": "SQL 指令"}
            },
            "required": ["type", "db", "sql"]
        }
    },
    {
        "name": "encrypt_file",
        "description": "加密或解密檔案。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["encrypt","decrypt"]},
                "path": {"type": "string", "description": "檔案路徑"},
                "password": {"type": "string", "description": "密碼"}
            },
            "required": ["action", "path", "password"]
        }
    },
    {
        "name": "qr_code",
        "description": "生成或掃描 QR Code，或監控剪貼簿變化。",
        "input_schema": {
            "type": "object",
            "properties": {
                "action": {"type": "string", "enum": ["qr_gen","qr_scan","clipboard_watch"]},
                "content": {"type": "string", "description": "QR Code 內容（qr_gen 使用）"},
                "path": {"type": "string", "description": "圖片路徑（qr_scan 使用）或儲存路徑（qr_gen 使用）"},
                "duration": {"type": "number", "description": "監控秒數（clipboard_watch 使用）"}
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

def execute_ai_plan(goal: str) -> str:
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=2048,
        system="""你是電腦自動化規劃師。把目標拆解成可執行步驟，以 JSON 陣列回傳：
[{"tool":"click","args":[x,y],"delay":0},{"tool":"type","args":["文字"]}]
可用工具：click/type/press/hotkey/open/screenshot/wait/notify/move/scroll
只回傳 JSON。""",
        messages=[{"role": "user", "content": f"目標：{goal}"}]
    )
    import json
    plan_text = response.content[0].text.strip()
    try:
        steps = json.loads(plan_text)
        results = []
        tool_map = {
            "click": lambda a: pyautogui.click(int(a[0]), int(a[1])),
            "type": lambda a: pyautogui.write(" ".join(str(x) for x in a), interval=0.05),
            "press": lambda a: pyautogui.press(a[0]),
            "hotkey": lambda a: pyautogui.hotkey(*a),
            "open": lambda a: subprocess.Popen(" ".join(str(x) for x in a), shell=True),
            "wait": lambda a: time.sleep(float(a[0])),
            "move": lambda a: pyautogui.moveTo(int(a[0]), int(a[1]), duration=0.3),
            "scroll": lambda a: pyautogui.scroll(int(a[1]) if a[0]=="up" else -int(a[1])),
        }
        for i, step in enumerate(steps, 1):
            t = step.get("tool"); a = step.get("args", []); d = step.get("delay", 0)
            if d: time.sleep(d)
            if t in tool_map:
                tool_map[t](a)
                results.append(f"步驟 {i} ✅ {t}")
            else:
                results.append(f"步驟 {i} ⚠️ 未知：{t}")
        return f"目標「{goal}」執行完畢\n" + "\n".join(results)
    except Exception as e:
        return f"規劃結果：{plan_text}\n執行錯誤：{e}"

def execute_screen_stream(duration=10, interval=2):
    screenshots = []
    end = time.time() + duration
    while time.time() < end:
        img = pyautogui.screenshot()
        buf = io.BytesIO()
        img.save(buf, format="PNG")
        screenshots.append(buf.getvalue())
        time.sleep(interval)
    return screenshots

def execute_drag(x1, y1, x2, y2, dur=0.5):
    pyautogui.moveTo(int(x1), int(y1))
    pyautogui.dragTo(int(x2), int(y2), duration=float(dur), button="left")
    return f"已拖曳 ({x1},{y1}) → ({x2},{y2})"

def execute_power(action):
    cmds = {
        "sleep": "rundll32.exe powrprof.dll,SetSuspendState 0,1,0",
        "restart": "shutdown /r /t 5",
        "shutdown": "shutdown /s /t 5",
    }
    subprocess.run(["powershell.exe", "-Command", cmds[action]])
    return f"已執行：{action}"

def execute_vdesktop(action):
    acts = {"left": ("ctrl","win","left"), "right": ("ctrl","win","right"), "new": ("ctrl","win","d")}
    pyautogui.hotkey(*acts[action])
    return f"虛擬桌面：{action}"

def execute_bluetooth(action, mac=""):
    try:
        import asyncio, bleak
        if action == "scan":
            async def _scan():
                return await bleak.BleakScanner.discover(timeout=5.0)
            devices = asyncio.run(_scan())
            return "\n".join(f"{d.address} {d.name or '(未知)'}" for d in devices) or "找不到裝置"
        elif action == "connect":
            async def _conn():
                async with bleak.BleakClient(mac) as c:
                    return f"已連線：{mac}（服務數：{len(c.services)}）"
            return asyncio.run(_conn())
    except Exception as e:
        return f"藍牙操作失敗：{e}"


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


def execute_reminder(time_str, message):
    import threading, time as t
    def _remind():
        if time_str.isdigit():
            t.sleep(int(time_str))
        else:
            import datetime as dt
            now = dt.datetime.now()
            target = dt.datetime.strptime(time_str, "%H:%M").replace(year=now.year, month=now.month, day=now.day)
            if target < now: target = target.replace(day=now.day+1)
            t.sleep((target-now).total_seconds())
        try:
            from win10toast import ToastNotifier
            ToastNotifier().show_toast("⏰ 提醒", message, duration=10)
        except Exception: pass
        try:
            import pyttsx3; e = pyttsx3.init(); e.say(message); e.runAndWait()
        except Exception: pass
    threading.Thread(target=_remind, daemon=True).start()
    return f"✅ 提醒已設定：{time_str} → {message}"


def execute_webpage_shot(action, url, selector="body", interval=60.0, duration=3600.0):
    try:
        if action == "screenshot":
            from playwright.sync_api import sync_playwright
            out = str(Path.home() / "Desktop" / f"webpage_{datetime.now().strftime('%H%M%S')}.png")
            with sync_playwright() as p:
                browser = p.chromium.launch()
                page = browser.new_page(viewport={"width": 1280, "height": 800})
                page.goto(url, timeout=30000)
                page.wait_for_load_state("networkidle")
                page.screenshot(path=out, full_page=True)
                browser.close()
            return f"✅ 網頁截圖已存：{out}"
        elif action == "monitor":
            import hashlib, time as t
            from bs4 import BeautifulSoup
            def _fetch():
                r = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
                soup = BeautifulSoup(r.text, "html.parser")
                text = "\n".join(e.get_text(strip=True) for e in soup.select(selector))
                return hashlib.md5(text.encode()).hexdigest(), text[:150]
            last_hash, _ = _fetch()
            end = t.time() + duration; changes = []
            while t.time() < end:
                t.sleep(interval)
                new_hash, snippet = _fetch()
                if new_hash != last_hash:
                    changes.append(f"[{datetime.now().strftime('%H:%M:%S')}] {snippet}")
                    last_hash = new_hash
            return f"監控結束，共 {len(changes)} 次變化\n" + "\n".join(changes[:5])
    except Exception as e:
        return f"❌ 失敗：{e}"


def execute_file_tools(action, path, dest="", pattern="", replacement="", ext=""):
    try:
        import re, filecmp, shutil
        if action == "batch_rename":
            files = [f for f in Path(path).iterdir() if f.is_file() and (not ext or f.suffix.lower()==ext.lower())]
            count = 0
            for f in sorted(files):
                new_name = re.sub(pattern, replacement, f.stem) + f.suffix
                if new_name != f.name: f.rename(f.parent / new_name); count += 1
            return f"✅ 已重新命名 {count} 個檔案"
        elif action == "sync":
            dest_path = Path(dest); dest_path.mkdir(parents=True, exist_ok=True)
            copied = 0
            for item in Path(path).rglob("*"):
                rel = item.relative_to(path); dst = dest_path / rel
                if item.is_dir(): dst.mkdir(parents=True, exist_ok=True)
                elif item.is_file() and (not dst.exists() or not filecmp.cmp(str(item), str(dst), shallow=False)):
                    shutil.copy2(str(item), str(dst)); copied += 1
            return f"✅ 同步完成：{copied} 個檔案更新"
    except Exception as e:
        return f"❌ 失敗：{e}"


def execute_image_tools(action, path="", quality=75, width=0, height=0, target_lang="zh-TW"):
    try:
        if action == "compress":
            from PIL import Image
            img = Image.open(path)
            if img.mode in ("RGBA","P"): img = img.convert("RGB")
            out = path.replace(".", f"_q{quality}.")
            img.save(out, optimize=True, quality=quality)
            orig = Path(path).stat().st_size; new = Path(out).stat().st_size
            return f"✅ {orig//1024}KB → {new//1024}KB（節省 {(1-new/orig)*100:.1f}%）：{out}"
        elif action == "batch":
            from PIL import Image as _Img
            folder = Path(path); out_dir = folder / "output"; out_dir.mkdir(exist_ok=True)
            count = 0
            for f in list(folder.glob("*.jpg")) + list(folder.glob("*.png")) + list(folder.glob("*.jpeg")):
                img = _Img.open(f)
                if width and height: img = img.resize((width, height))
                if img.mode in ("RGBA","P"): img = img.convert("RGB")
                img.save(str(out_dir/f.name), optimize=True, quality=quality); count += 1
            return f"✅ 批次處理 {count} 張 → {out_dir}"
        elif action == "ocr_translate":
            import easyocr, numpy as np
            from PIL import Image
            from deep_translator import GoogleTranslator
            reader = easyocr.Reader(["ch_tra","en"], gpu=False)
            img = Image.open(path) if path else pyautogui.screenshot()
            text = " ".join(r[1] for r in reader.readtext(np.array(img)))
            if not text.strip(): return "❌ 未辨識到文字"
            translated = GoogleTranslator(source="auto", target=target_lang).translate(text)
            return f"原文：{text[:200]}\n翻譯：{translated}"
    except Exception as e:
        return f"❌ 失敗：{e}"


def execute_lookup(action, ip="", amount=1.0, from_cur="USD", to_cur="TWD"):
    try:
        if action == "ip":
            d = requests.get(f"http://ip-api.com/json/{ip}?lang=zh-TW", timeout=10).json()
            if d.get("status") == "fail": return f"❌ {d.get('message')}"
            return f"IP：{d['query']}\n國家：{d['country']}\n城市：{d['city']}\nISP：{d['isp']}\n時區：{d['timezone']}\n座標：{d['lat']},{d['lon']}"
        elif action == "currency":
            d = requests.get(f"https://api.frankfurter.app/latest?amount={amount}&from={from_cur.upper()}&to={to_cur.upper()}", timeout=10).json()
            rate = d["rates"].get(to_cur.upper())
            return f"💱 {amount} {from_cur.upper()} = {rate:.4f} {to_cur.upper()}（{d.get('date')}）"
    except Exception as e:
        return f"❌ 查詢失敗：{e}"


def execute_system_tools(action, **kwargs):
    try:
        if action == "event_log":
            log_name = kwargs.get("log_name","System"); level = kwargs.get("level","Error"); count = kwargs.get("count",10)
            r = subprocess.run(["powershell.exe","-Command",f"Get-WinEvent -LogName '{log_name}' -MaxEvents {count} | Where-Object {{$_.LevelDisplayName -eq '{level}'}} | Select-Object TimeCreated,Message | Format-List"], capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30)
            return r.stdout[:2000] or f"（無 {level} 事件）"
        elif action == "usb_list":
            r = subprocess.run(["powershell.exe","-Command","Get-PnpDevice | Where-Object {$_.Class -eq 'USB' -and $_.Status -eq 'OK'} | Select-Object FriendlyName | Format-Table"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout[:2000] or "（無 USB 裝置）"
        elif action == "firewall_list":
            r = subprocess.run(["powershell.exe","-Command","Get-NetFirewallRule | Where-Object {$_.Enabled -eq 'True'} | Select-Object DisplayName,Direction,Action | Format-Table -AutoSize"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout[:2000]
        elif action == "firewall_add":
            name=kwargs.get("name",""); direction=kwargs.get("direction","in"); port=kwargs.get("port",80)
            r = subprocess.run(["powershell.exe","-Command",f"New-NetFirewallRule -DisplayName '{name}' -Direction {direction} -Protocol TCP -LocalPort {port} -Action Allow"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout or f"✅ 防火牆規則已新增：{name}"
        elif action == "firewall_remove":
            r = subprocess.run(["powershell.exe","-Command",f"Remove-NetFirewallRule -DisplayName '{kwargs.get('name','')}'"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout or f"✅ 防火牆規則已刪除"
        elif action == "printer_list":
            r = subprocess.run(["powershell.exe","-Command","Get-Printer | Select-Object Name,PrinterStatus | Format-Table"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout or "（無印表機）"
        elif action == "printer_jobs":
            r = subprocess.run(["powershell.exe","-Command","Get-PrintJob -PrinterName (Get-Printer | Select-Object -First 1 -ExpandProperty Name) | Format-Table"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout or "（列印佇列為空）"
        elif action == "net_share_list":
            r = subprocess.run(["net","use"], capture_output=True, text=True, encoding="cp950", errors="replace")
            return r.stdout
        elif action == "net_share_connect":
            args = ["net","use",kwargs.get("drive","Z:"),kwargs.get("share_path","")]
            if kwargs.get("user"): args += [f"/user:{kwargs['user']}", kwargs.get("password","")]
            r = subprocess.run(args, capture_output=True, text=True, encoding="cp950", errors="replace")
            return r.stdout or f"✅ 已連線"
        elif action == "net_share_disconnect":
            r = subprocess.run(["net","use",kwargs.get("drive","Z:"),"/delete"], capture_output=True, text=True, encoding="cp950", errors="replace")
            return r.stdout or "✅ 已中斷"
        elif action == "font_list":
            kw = kwargs.get("keyword","")
            r = subprocess.run(["powershell.exe","-Command","[System.Reflection.Assembly]::LoadWithPartialName('System.Drawing') | Out-Null; [System.Drawing.FontFamily]::Families | Select-Object -ExpandProperty Name"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            fonts = [f for f in r.stdout.strip().splitlines() if kw.lower() in f.lower()] if kw else r.stdout.strip().splitlines()
            return "\n".join(fonts[:50]) + (f"\n...共 {len(fonts)} 個" if len(fonts)>50 else "")
        elif action == "rdp":
            subprocess.Popen(["mstsc", f"/v:{kwargs.get('host','')}"])
            return f"✅ 正在開啟 RDP：{kwargs.get('host','')}"
    except Exception as e:
        return f"❌ 失敗：{e}"


def generate_voice_ogg(text: str, voice: str = "zh-TW-YunJheNeural") -> bytes:
    """生成語音並回傳 OGG OPUS bytes（Telegram voice message 格式）"""
    import edge_tts, asyncio, tempfile, subprocess as sp
    import imageio_ffmpeg

    ffmpeg_exe = imageio_ffmpeg.get_ffmpeg_exe()
    tmp_mp3 = tempfile.NamedTemporaryFile(suffix=".mp3", delete=False)
    tmp_mp3.close()
    tmp_ogg = tempfile.NamedTemporaryFile(suffix=".ogg", delete=False)
    tmp_ogg.close()

    # 生成 MP3
    async def _gen():
        await edge_tts.Communicate(text, voice).save(tmp_mp3.name)
    asyncio.run(_gen())

    # 轉換成 OGG OPUS（Telegram voice 格式）
    sp.run([
        ffmpeg_exe, "-y", "-i", tmp_mp3.name,
        "-c:a", "libopus", "-b:a", "64k",
        tmp_ogg.name
    ], capture_output=True)

    data = Path(tmp_ogg.name).read_bytes()
    Path(tmp_mp3.name).unlink(missing_ok=True)
    Path(tmp_ogg.name).unlink(missing_ok=True)
    return data


def execute_tts_advanced(action, text="", voice="zh-TW-YunJheNeural"):
    try:
        import edge_tts, asyncio
        if action == "speak":
            out = str(Path.home() / "Desktop" / f"tts_{datetime.now().strftime('%H%M%S')}.mp3")
            async def _gen():
                await edge_tts.Communicate(text, voice).save(out)
            asyncio.run(_gen())
            subprocess.Popen(["powershell.exe","-Command",f"Start-Process '{out}'"])
            return f"✅ Edge TTS 語音已播放：{voice}"
        elif action == "list_voices":
            async def _list(): return await edge_tts.list_voices()
            voices = asyncio.run(_list())
            return "\n".join(f"{v['ShortName']}  {v['FriendlyName']}" for v in voices if v["Locale"].startswith("zh"))
    except Exception as e:
        return f"❌ Edge TTS 失敗：{e}"


def execute_todo(action, task="", todo_id=0):
    try:
        db = str(Path("C:/Users/blue_/claude-telegram-bot/todo.db"))
        conn = sqlite3.connect(db)
        conn.execute("CREATE TABLE IF NOT EXISTS todos (id INTEGER PRIMARY KEY AUTOINCREMENT, task TEXT, done INTEGER DEFAULT 0, created TEXT)")
        conn.commit()
        if action == "add":
            conn.execute("INSERT INTO todos (task,created) VALUES (?,?)", (task, datetime.now().strftime("%Y-%m-%d %H:%M")))
            conn.commit(); conn.close(); return f"✅ 已新增：{task}"
        elif action == "list":
            rows = conn.execute("SELECT id,task,done,created FROM todos ORDER BY done,id").fetchall()
            conn.close()
            return "\n".join(f"{'✅' if r[2] else '⬜'} [{r[0]}] {r[1]}" for r in rows) or "（清單為空）"
        elif action == "done":
            conn.execute("UPDATE todos SET done=1 WHERE id=?", (todo_id,)); conn.commit(); conn.close(); return f"✅ 任務 #{todo_id} 已完成"
        elif action == "delete":
            conn.execute("DELETE FROM todos WHERE id=?", (todo_id,)); conn.commit(); conn.close(); return f"✅ 任務 #{todo_id} 已刪除"
        elif action == "clear":
            conn.execute("DELETE FROM todos WHERE done=1"); conn.commit(); conn.close(); return "✅ 已清除所有已完成任務"
    except Exception as e:
        return f"❌ 任務清單失敗：{e}"


def execute_sysres_chart(duration=10):
    try:
        import psutil, matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt, time as t
        cpu_vals, mem_vals = [], []
        for _ in range(duration):
            cpu_vals.append(psutil.cpu_percent(interval=1))
            mem_vals.append(psutil.virtual_memory().percent)
        out = str(Path.home() / "Desktop" / f"sysres_{datetime.now().strftime('%H%M%S')}.png")
        fig, ax = plt.subplots()
        ax.plot(range(1,duration+1), cpu_vals, label="CPU %", color="blue")
        ax.plot(range(1,duration+1), mem_vals, label="RAM %", color="orange")
        ax.set_ylim(0,100); ax.legend(); ax.set_title("系統資源使用率")
        plt.tight_layout(); plt.savefig(out); plt.close()
        return out
    except Exception as e:
        return f"❌ 失敗：{e}"


def execute_password_mgr(action, site, master, username="", password=""):
    try:
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
        from cryptography.hazmat.primitives import hashes
        from cryptography.fernet import Fernet
        import base64
        kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=b"pwd_manager_v1", iterations=100000)
        key = base64.urlsafe_b64encode(kdf.derive(master.encode()))
        f = Fernet(key)
        db = str(Path("C:/Users/blue_/claude-telegram-bot/passwords.db"))
        conn = sqlite3.connect(db)
        conn.execute("CREATE TABLE IF NOT EXISTS passwords (id INTEGER PRIMARY KEY AUTOINCREMENT, site TEXT, username TEXT, password TEXT)")
        conn.commit()
        if action == "save":
            enc = f.encrypt(password.encode()).decode()
            conn.execute("INSERT INTO passwords (site,username,password) VALUES (?,?,?)", (site, username, enc))
            conn.commit(); conn.close(); return f"✅ 密碼已儲存：{site}"
        elif action == "get":
            rows = conn.execute("SELECT site,username,password FROM passwords WHERE site LIKE ?", (f"%{site}%",)).fetchall()
            conn.close()
            if not rows: return f"（找不到 {site} 的密碼）"
            results = []
            for s, u, enc_p in rows:
                try: results.append(f"🔑 {s}\n帳號：{u}\n密碼：{f.decrypt(enc_p.encode()).decode()}")
                except Exception: results.append(f"❌ {s} 解密失敗（主密碼錯誤？）")
            return "\n".join(results)
    except Exception as e:
        return f"❌ 密碼管理失敗：{e}"


def execute_clipboard_image(action, path=""):
    try:
        import win32clipboard
        from PIL import Image
        import io as _io
        if action == "get":
            win32clipboard.OpenClipboard()
            try: data = win32clipboard.GetClipboardData(win32clipboard.CF_DIB)
            finally: win32clipboard.CloseClipboard()
            out = path or str(Path.home() / "Desktop" / f"clipboard_{datetime.now().strftime('%H%M%S')}.png")
            Image.open(_io.BytesIO(data)).save(out)
            return f"✅ 剪貼簿圖片已存：{out}"
        elif action == "set":
            img = Image.open(path).convert("RGB")
            buf = _io.BytesIO(); img.save(buf,"BMP"); data = buf.getvalue()[14:]
            win32clipboard.OpenClipboard()
            try: win32clipboard.EmptyClipboard(); win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
            finally: win32clipboard.CloseClipboard()
            return f"✅ 圖片已複製到剪貼簿"
    except Exception as e:
        return f"❌ 剪貼簿圖片失敗：{e}"


def execute_email_read(host, user, password, folder="INBOX", count=5):
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
            body = ""
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain":
                        body = part.get_payload(decode=True).decode(errors="replace")[:150]; break
            else:
                body = msg.get_payload(decode=True).decode(errors="replace")[:150]
            results.append(f"寄件人：{msg['From']}\n主旨：{subject}\n日期：{msg['Date']}\n{body}\n{'─'*30}")
        client.logout()
        return "\n".join(results) if results else "（收件匣為空）"
    except Exception as e:
        return f"❌ 讀取郵件失敗：{e}"


def execute_calendar(action, days=7, title="", start="", end="", description=""):
    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        from datetime import timezone, timedelta
        creds_path = Path("C:/Users/blue_/claude-telegram-bot/gcal_token.json")
        if not creds_path.exists():
            return "❌ 未找到 Google Calendar 憑證（gcal_token.json）"
        creds = Credentials.from_authorized_user_file(str(creds_path))
        service = build("calendar", "v3", credentials=creds)
        if action == "list":
            now = datetime.now(timezone.utc)
            events = service.events().list(
                calendarId="primary", timeMin=now.isoformat(),
                timeMax=(now + timedelta(days=days)).isoformat(),
                maxResults=20, singleEvents=True, orderBy="startTime"
            ).execute().get("items", [])
            if not events:
                return f"未來 {days} 天沒有行程"
            return "\n".join(f"📅 {e['start'].get('dateTime',e['start'].get('date'))}  {e.get('summary','（無標題）')}" for e in events)
        elif action == "add":
            event = {"summary": title, "description": description,
                     "start": {"dateTime": start, "timeZone": "Asia/Taipei"},
                     "end": {"dateTime": end, "timeZone": "Asia/Taipei"}}
            created = service.events().insert(calendarId="primary", body=event).execute()
            return f"✅ 行程已新增：{created.get('summary')}"
    except Exception as e:
        return f"❌ 行事曆操作失敗：{e}"


def execute_global_hotkey(hotkey, command, duration=60.0):
    try:
        import keyboard as kb, time as t
        triggered = []
        def on_trigger():
            triggered.append(datetime.now().strftime("%H:%M:%S"))
            subprocess.run(command, shell=True)
        kb.add_hotkey(hotkey, on_trigger)
        t.sleep(duration)
        kb.remove_all_hotkeys()
        return f"✅ 快捷鍵 [{hotkey}] 共觸發 {len(triggered)} 次"
    except Exception as e:
        return f"❌ 快捷鍵監聽失敗：{e}"


def execute_git(action, repo=".", message="", branch="master"):
    try:
        import git as _git
        r = _git.Repo(repo)
        if action == "status": return r.git.status()
        elif action == "log":
            return "\n".join(f"{c.hexsha[:7]} [{c.authored_datetime.strftime('%m-%d %H:%M')}] {c.message.strip()[:60]}" for c in list(r.iter_commits())[:10])
        elif action == "pull":
            result = r.remotes.origin.pull(); return f"✅ Pull 完成"
        elif action == "add": r.git.add(A=True); return "✅ git add -A 完成"
        elif action == "commit": r.index.commit(message or "auto commit"); return f"✅ committed: {message}"
        elif action == "push": r.remotes.origin.push(branch); return f"✅ pushed to origin/{branch}"
        elif action == "diff": return r.git.diff()[:2000] or "（無變更）"
        return "未知動作"
    except Exception as e:
        return f"❌ Git 失敗：{e}"


def execute_hardware():
    try:
        import psutil
        battery = psutil.sensors_battery()
        bat = f"{battery.percent:.0f}% {'充電中' if battery.power_plugged else '使用電池'}" if battery else "無電池"
        gpu_str = ""
        try:
            import GPUtil
            gpus = GPUtil.getGPUs()
            gpu_str = "\n".join(f"GPU {g.name}: 使用率 {g.load*100:.0f}% | 記憶體 {g.memoryUsed:.0f}/{g.memoryTotal:.0f}MB | 溫度 {g.temperature}°C" for g in gpus)
        except Exception:
            gpu_str = "GPU：未偵測到 NVIDIA GPU"
        return f"🔋 電池：{bat}\n{gpu_str}"
    except Exception as e:
        return f"❌ 硬體監控失敗：{e}"


def execute_report(title, data_json, output=""):
    try:
        import jinja2, json
        data = json.loads(data_json)
        out_path = output or str(Path.home() / "Desktop" / f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.html")
        tmpl = jinja2.Template("""<!DOCTYPE html><html><head><meta charset="utf-8">
<style>body{font-family:sans-serif;margin:40px}table{border-collapse:collapse;width:100%}
th,td{border:1px solid #ccc;padding:8px}th{background:#4472C4;color:white}
tr:nth-child(even){background:#f2f2f2}h1{color:#4472C4}</style></head>
<body><h1>{{ title }}</h1><p>生成時間：{{ time }}</p>
{% for section, rows in data.items() %}<h2>{{ section }}</h2>
{% if rows is iterable and rows is not string %}{% if rows[0] is mapping %}
<table><tr>{% for k in rows[0].keys() %}<th>{{ k }}</th>{% endfor %}</tr>
{% for row in rows %}<tr>{% for v in row.values() %}<td>{{ v }}</td>{% endfor %}</tr>{% endfor %}</table>
{% else %}<ul>{% for i in rows %}<li>{{ i }}</li>{% endfor %}</ul>{% endif %}
{% else %}<p>{{ rows }}</p>{% endif %}{% endfor %}</body></html>""")
        Path(out_path).write_text(tmpl.render(title=title, data=data, time=datetime.now().strftime("%Y-%m-%d %H:%M:%S")), encoding="utf-8")
        return f"✅ 報告已生成：{out_path}"
    except Exception as e:
        return f"❌ 報告生成失敗：{e}"


def execute_dropbox(action, local, remote, token=""):
    try:
        import dropbox as dbx
        tok = token or os.getenv("DROPBOX_TOKEN", "")
        if not tok: return "❌ 請設定 DROPBOX_TOKEN 環境變數"
        d = dbx.Dropbox(tok)
        if action == "upload":
            with open(local, "rb") as f:
                d.files_upload(f.read(), remote, mode=dbx.files.WriteMode.overwrite)
            return f"✅ 已上傳到 Dropbox：{remote}"
        elif action == "download":
            _, res = d.files_download(remote)
            Path(local).write_bytes(res.content)
            return f"✅ 已從 Dropbox 下載：{local}"
    except Exception as e:
        return f"❌ Dropbox 失敗：{e}"


def execute_docker(action, name=""):
    try:
        import docker as _docker
        client = _docker.from_env()
        if action == "list":
            return "\n".join(f"[{c.status}] {c.name} {c.image.tags}" for c in client.containers.list(all=True)) or "（無容器）"
        elif action == "start": client.containers.get(name).start(); return f"✅ {name} 已啟動"
        elif action == "stop": client.containers.get(name).stop(); return f"✅ {name} 已停止"
        elif action == "logs": return client.containers.get(name).logs(tail=50).decode(errors="replace")
        elif action == "images": return "\n".join(f"{img.tags} {img.short_id}" for img in client.images.list())
    except Exception as e:
        return f"❌ Docker 失敗：{e}"


def execute_pdf_to_image(path, output_dir="", dpi=150):
    try:
        import fitz
        doc = fitz.open(path)
        out = Path(output_dir) if output_dir else Path(path).parent / (Path(path).stem + "_imgs")
        out.mkdir(parents=True, exist_ok=True)
        for i, page in enumerate(doc):
            pix = page.get_pixmap(matrix=fitz.Matrix(dpi/72, dpi/72))
            pix.save(str(out / f"page_{i+1}.png"))
        return f"✅ 已轉換 {len(doc)} 頁到：{out}"
    except Exception as e:
        return f"❌ PDF 轉圖片失敗：{e}"


def execute_barcode(image_path=""):
    try:
        from pyzbar.pyzbar import decode
        from PIL import Image
        img = Image.open(image_path) if image_path else pyautogui.screenshot()
        results = decode(img)
        if not results: return "❌ 未偵測到條碼"
        return "\n".join(f"類型：{r.type}  內容：{r.data.decode('utf-8', errors='replace')}" for r in results)
    except Exception as e:
        return f"❌ 條碼掃描失敗：{e}"


def execute_nlp(action, text):
    try:
        import anthropic as _ant
        c = _ant.Anthropic(api_key=os.getenv("ANTHROPIC_API_KEY"))
        prompt = f"請用繁體中文摘要以下文字（100字以內）：\n\n{text}" if action == "summarize" else f"分析以下文字的情緒，只回覆：正面/負面/中性 + 一句說明：\n\n{text}"
        msg = c.messages.create(model="claude-haiku-4-5-20251001", max_tokens=256, messages=[{"role":"user","content":prompt}])
        return msg.content[0].text
    except Exception as e:
        return f"❌ NLP 失敗：{e}"


def execute_vpn(action, name="", user="", password=""):
    try:
        if action == "list":
            r = subprocess.run(["powershell.exe","-Command","Get-VpnConnection | Select-Object Name,ConnectionStatus | Format-Table"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout or "（未設定 VPN）"
        elif action == "connect":
            r = subprocess.run(["rasdial", name, user, password], capture_output=True, text=True, encoding="cp950", errors="replace")
            return r.stdout.strip()
        elif action == "disconnect":
            r = subprocess.run(["rasdial", name, "/disconnect"], capture_output=True, text=True, encoding="cp950", errors="replace")
            return r.stdout.strip()
    except Exception as e:
        return f"❌ VPN 失敗：{e}"


def execute_restore_point(action, description=""):
    try:
        if action == "create":
            ps = f"Checkpoint-Computer -Description '{description or 'Claude Auto Restore'}' -RestorePointType MODIFY_SETTINGS"
            r = subprocess.run(["powershell.exe","-Command",ps], capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout or "✅ 還原點已建立"
        elif action == "list":
            r = subprocess.run(["powershell.exe","-Command","Get-ComputerRestorePoint | Select-Object SequenceNumber,Description,CreationTime | Format-Table"], capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout or "（無還原點）"
    except Exception as e:
        return f"❌ 系統還原點失敗：{e}"


def execute_disk_analyze(path="C:/", top=10):
    try:
        import psutil
        usage = psutil.disk_usage(path)
        result = f"磁碟：{path}\n總容量：{usage.total/1024**3:.1f} GB | 已使用：{usage.used/1024**3:.1f} GB ({usage.percent}%) | 可用：{usage.free/1024**3:.1f} GB\n\n"
        sizes = []
        for item in Path(path).iterdir():
            try:
                size = sum(f.stat().st_size for f in item.rglob("*") if f.is_file()) if item.is_dir() else item.stat().st_size
                sizes.append((size, str(item)))
            except Exception: pass
        sizes.sort(reverse=True)
        result += "\n".join(f"{s/1024**3:.2f} GB  {n}" for s, n in sizes[:top])
        return result
    except Exception as e:
        return f"❌ 磁碟分析失敗：{e}"


def execute_face_detect(image_path="", output=""):
    try:
        import cv2, numpy as np
        img = cv2.cvtColor(np.array(pyautogui.screenshot()), cv2.COLOR_RGB2BGR) if not image_path else cv2.imread(image_path)
        cascade = cv2.CascadeClassifier(cv2.data.haarcascades + "haarcascade_frontalface_default.xml")
        faces = cascade.detectMultiScale(cv2.cvtColor(img, cv2.COLOR_BGR2GRAY), 1.1, 4)
        for (x, y, w, h) in faces:
            cv2.rectangle(img, (x, y), (x+w, y+h), (0, 255, 0), 2)
        out_path = output or str(Path.home() / "Desktop" / f"faces_{datetime.now().strftime('%H%M%S')}.jpg")
        cv2.imwrite(out_path, img)
        return f"✅ 偵測到 {len(faces)} 張人臉：{out_path}"
    except Exception as e:
        return f"❌ 人臉偵測失敗：{e}"


def execute_video_gif(path, start=0, duration=5.0, output="", fps=10):
    try:
        import imageio
        out = output or path.replace(".mp4", ".gif")
        reader = imageio.get_reader(path)
        video_fps = reader.get_meta_data().get("fps", 30)
        frames = [f for i, f in enumerate(reader) if int(start*video_fps) <= i < int((start+duration)*video_fps)]
        imageio.mimsave(out, frames, fps=fps)
        return f"✅ GIF 已生成：{out}（{len(frames)} 幀）"
    except Exception as e:
        return f"❌ 影片轉 GIF 失敗：{e}"


def execute_excel_chart(path, sheet, chart_type="bar", title=""):
    try:
        import openpyxl
        from openpyxl.chart import BarChart, LineChart, PieChart, Reference
        wb = openpyxl.load_workbook(path)
        ws = wb[sheet] if sheet in wb.sheetnames else wb.active
        chart = {"bar": BarChart, "line": LineChart, "pie": PieChart}.get(chart_type, BarChart)()
        chart.title = title or sheet; chart.style = 10
        data = Reference(ws, min_col=2, min_row=1, max_row=ws.max_row, max_col=ws.max_column)
        chart.add_data(data, titles_from_data=True)
        chart.set_categories(Reference(ws, min_col=1, min_row=2, max_row=ws.max_row))
        ws.add_chart(chart, "A" + str(ws.max_row + 2))
        wb.save(path); return f"✅ 圖表已加入：{path}"
    except Exception as e:
        return f"❌ Excel 圖表失敗：{e}"


def execute_speedtest():
    try:
        import speedtest as st
        s = st.Speedtest(); s.get_best_server()
        dl = s.download()/1_000_000; ul = s.upload()/1_000_000
        return f"📶 下載：{dl:.1f} Mbps | 上傳：{ul:.1f} Mbps | Ping：{s.results.ping:.0f} ms | 伺服器：{s.results.server.get('name','')}"
    except Exception as e:
        return f"❌ 速度測試失敗：{e}"


def execute_screenshot_compare(img1_path="", img2_path="", output=""):
    try:
        import cv2, numpy as np, time as t
        img1 = cv2.cvtColor(np.array(pyautogui.screenshot()), cv2.COLOR_RGB2BGR) if not img1_path else cv2.imread(img1_path)
        if not img2_path: t.sleep(2); img2 = cv2.cvtColor(np.array(pyautogui.screenshot()), cv2.COLOR_RGB2BGR)
        else: img2 = cv2.imread(img2_path)
        h = min(img1.shape[0], img2.shape[0]); w = min(img1.shape[1], img2.shape[1])
        diff = cv2.absdiff(img1[:h,:w], img2[:h,:w])
        _, thresh = cv2.threshold(cv2.cvtColor(diff, cv2.COLOR_BGR2GRAY), 30, 255, cv2.THRESH_BINARY)
        contours, _ = cv2.findContours(thresh, cv2.RETR_EXTERNAL, cv2.CHAIN_APPROX_SIMPLE)
        result = img2[:h,:w].copy(); cv2.drawContours(result, contours, -1, (0, 0, 255), 2)
        out = output or str(Path.home() / "Desktop" / f"diff_{datetime.now().strftime('%H%M%S')}.png")
        cv2.imwrite(out, result)
        pct = cv2.countNonZero(thresh) / (h * w) * 100
        return f"差異：{pct:.2f}%，標記圖：{out}"
    except Exception as e:
        return f"❌ 截圖比對失敗：{e}"


def execute_screen_record(action, duration=10.0, output=""):
    try:
        if action == "record":
            import mss, cv2, numpy as np, time as t
            out_path = output or str(Path.home() / "Desktop" / f"record_{datetime.now().strftime('%H%M%S')}.mp4")
            with mss.mss() as sct:
                mon = sct.monitors[1]
                w, h = mon["width"], mon["height"]
                writer = cv2.VideoWriter(out_path, cv2.VideoWriter_fourcc(*"mp4v"), 10, (w, h))
                end = t.time() + duration
                while t.time() < end:
                    frame = np.array(sct.grab(mon))
                    writer.write(cv2.cvtColor(frame, cv2.COLOR_BGRA2BGR))
                    t.sleep(0.1)
                writer.release()
            return f"✅ 錄影完成：{out_path}"
        elif action == "webcam":
            import cv2
            out_path = output or str(Path.home() / "Desktop" / f"webcam_{datetime.now().strftime('%H%M%S')}.jpg")
            cap = cv2.VideoCapture(0)
            ret, frame = cap.read()
            cap.release()
            if ret:
                cv2.imwrite(out_path, frame)
                return f"✅ 已拍照：{out_path}"
            return "❌ 無法存取攝影機"
    except Exception as e:
        return f"❌ 失敗：{e}"


def execute_translate(text, target="zh-TW", source="auto"):
    try:
        from deep_translator import GoogleTranslator
        return GoogleTranslator(source=source, target=target).translate(text)
    except Exception as e:
        return f"❌ 翻譯失敗：{e}"


def execute_chart(chart_type, data_json, title="", output=""):
    try:
        import matplotlib
        matplotlib.use("Agg")
        import matplotlib.pyplot as plt, json
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
        return out_path
    except Exception as e:
        return f"❌ 圖表生成失敗：{e}"


def execute_pptx(action, path, slides=""):
    try:
        from pptx import Presentation
        if action == "read":
            prs = Presentation(path)
            lines = []
            for i, slide in enumerate(prs.slides, 1):
                texts = [sh.text for sh in slide.shapes if sh.has_text_frame]
                lines.append(f"[投影片 {i}] " + " | ".join(t for t in texts if t.strip()))
            return "\n".join(lines) or "（簡報為空）"
        elif action == "create":
            import json
            from pptx.util import Pt
            data = json.loads(slides)
            prs = Presentation()
            for s in data:
                sl = prs.slides.add_slide(prs.slide_layouts[1])
                if "title" in s:
                    sl.shapes.title.text = s["title"]
                if "body" in s:
                    sl.placeholders[1].text = s["body"]
            prs.save(path)
            return f"✅ 已建立簡報：{path}"
    except Exception as e:
        return f"❌ PPT 操作失敗：{e}"


def execute_api_call(method, url, headers="{}", body="{}"):
    try:
        import json
        h = json.loads(headers) if headers else {}
        b = json.loads(body) if body else None
        resp = requests.request(method.upper(), url, headers=h, json=b, timeout=30)
        try:
            return json.dumps(resp.json(), ensure_ascii=False, indent=2)[:2000]
        except Exception:
            return resp.text[:2000]
    except Exception as e:
        return f"❌ API 呼叫失敗：{e}"


def execute_watchdog(process, script, duration=60.0):
    import psutil, time as t
    restarts = 0
    end = t.time() + duration
    while t.time() < end:
        running = any(p.name().lower() == process.lower() for p in psutil.process_iter())
        if not running:
            subprocess.Popen(["pythonw", script] if script.endswith(".py") else [script])
            restarts += 1
        t.sleep(5)
    return f"守護結束，共重啟 {restarts} 次"


def execute_ssh_sftp(action, host, user, password, command="", local="", remote="", port=22):
    try:
        import paramiko
        if action == "ssh_run":
            c = paramiko.SSHClient()
            c.set_missing_host_key_policy(paramiko.AutoAddPolicy())
            c.connect(host, port=port, username=user, password=password, timeout=15)
            _, stdout, stderr = c.exec_command(command)
            out = stdout.read().decode(errors="replace") + stderr.read().decode(errors="replace")
            c.close()
            return out.strip() or "（執行完畢，無輸出）"
        else:
            t = paramiko.Transport((host, port))
            t.connect(username=user, password=password)
            sftp = paramiko.SFTPClient.from_transport(t)
            if action == "sftp_upload":
                sftp.put(local, remote)
                result = f"✅ 已上傳：{local} → {remote}"
            else:
                sftp.get(remote, local)
                result = f"✅ 已下載：{remote} → {local}"
            sftp.close(); t.close()
            return result
    except Exception as e:
        return f"❌ SSH/SFTP 失敗：{e}"


def execute_network_diag(action, host, ports="22,80,443,3306,3389,8080"):
    try:
        if action == "ping":
            r = subprocess.run(["ping", "-n", "4", host], capture_output=True, text=True, encoding="cp950", errors="replace", timeout=20)
            return r.stdout.strip()
        elif action == "traceroute":
            r = subprocess.run(["tracert", host], capture_output=True, text=True, encoding="cp950", errors="replace", timeout=60)
            return r.stdout[:2000]
        elif action == "portscan":
            import socket
            results = []
            for p in [int(x) for x in ports.split(",")]:
                s = socket.socket(); s.settimeout(1)
                r = s.connect_ex((host, p))
                results.append(f"Port {p}: {'開放 ✅' if r == 0 else '關閉 ❌'}")
                s.close()
            return "\n".join(results)
    except Exception as e:
        return f"❌ 網路診斷失敗：{e}"


def execute_win_service(action, name=""):
    try:
        if action == "list":
            r = subprocess.run(["powershell.exe", "-Command",
                "Get-Service | Select-Object Name,Status | Format-Table -AutoSize"],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout[:2000]
        else:
            cmd = f"{'Start' if action=='start' else 'Stop'}-Service -Name '{name}' -Force"
            r = subprocess.run(["powershell.exe", "-Command", cmd],
                capture_output=True, text=True, encoding="utf-8", errors="replace")
            return r.stdout or r.stderr or f"✅ {action} {name}"
    except Exception as e:
        return f"❌ 服務操作失敗：{e}"


def execute_pdf_edit(action, path="", output="", paths="", text=""):
    try:
        import fitz, json
        if action == "merge":
            pdf_list = json.loads(paths)
            writer = fitz.open()
            for p in pdf_list:
                writer.insert_pdf(fitz.open(p))
            writer.save(output)
            return f"✅ 已合併 {len(pdf_list)} 個 PDF：{output}"
        elif action == "split":
            doc = fitz.open(path)
            out_dir = Path(output)
            out_dir.mkdir(parents=True, exist_ok=True)
            for i, _ in enumerate(doc):
                out = fitz.open()
                out.insert_pdf(doc, from_page=i, to_page=i)
                out.save(str(out_dir / f"page_{i+1}.pdf"))
            return f"✅ 已分割 {len(doc)} 頁到：{output}"
        elif action == "watermark":
            doc = fitz.open(path)
            out_path = output or path.replace(".pdf", "_wm.pdf")
            for page in doc:
                page.insert_text((page.rect.width/2-50, page.rect.height/2),
                    text, fontsize=40, color=(0.8,0.8,0.8), rotate=45)
            doc.save(out_path)
            return f"✅ 已加浮水印：{out_path}"
    except Exception as e:
        return f"❌ PDF 編輯失敗：{e}"


def execute_audio_process(action, input_path, output="", start_ms=0, end_ms=0):
    try:
        from pydub import AudioSegment
        if action == "convert":
            fmt = Path(output).suffix.lstrip(".")
            AudioSegment.from_file(input_path).export(output, format=fmt)
            return f"✅ 已轉換：{output}"
        elif action == "trim":
            audio = AudioSegment.from_file(input_path)[start_ms:end_ms]
            out = output or input_path.replace(".", "_trim.")
            audio.export(out, format=Path(out).suffix.lstrip("."))
            return f"✅ 已剪輯：{out}"
    except Exception as e:
        return f"❌ 音訊處理失敗：{e}"


def execute_push_notify(platform, message, webhook_or_token):
    try:
        if platform == "discord":
            resp = requests.post(webhook_or_token, json={"content": message}, timeout=10)
            return f"✅ Discord 已發送（{resp.status_code}）"
        elif platform == "line":
            resp = requests.post(
                "https://notify-api.line.me/api/notify",
                headers={"Authorization": f"Bearer {webhook_or_token}"},
                data={"message": message}, timeout=10
            )
            return f"✅ LINE 已發送（{resp.status_code}）"
    except Exception as e:
        return f"❌ 推播失敗：{e}"


def execute_disk_backup(action, src="", dest=""):
    try:
        import tempfile, shutil
        tmp = Path(tempfile.gettempdir())
        if action == "list_temp":
            files = list(tmp.rglob("*"))
            total = sum(f.stat().st_size for f in files if f.is_file())
            return f"暫存資料夾：{tmp}\n檔案數：{len(files)}\n佔用空間：{total/1024/1024:.1f} MB"
        elif action == "clean_temp":
            count = 0
            for f in tmp.iterdir():
                try:
                    if f.is_file(): f.unlink(); count += 1
                    elif f.is_dir(): shutil.rmtree(f, ignore_errors=True); count += 1
                except Exception: pass
            return f"✅ 已清理 {count} 個暫存項目"
        elif action == "backup":
            out = Path(dest) / f"{Path(src).name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
            shutil.make_archive(str(out), "zip", src)
            return f"✅ 備份完成：{out}.zip"
    except Exception as e:
        return f"❌ 磁碟操作失敗：{e}"


def execute_registry(action, key, value_name="", value=""):
    try:
        import winreg
        parts = key.split("\\", 1)
        roots = {"HKEY_LOCAL_MACHINE": winreg.HKEY_LOCAL_MACHINE,
                 "HKEY_CURRENT_USER": winreg.HKEY_CURRENT_USER,
                 "HKLM": winreg.HKEY_LOCAL_MACHINE, "HKCU": winreg.HKEY_CURRENT_USER}
        root = roots[parts[0]]
        if action == "read":
            with winreg.OpenKey(root, parts[1]) as k:
                if value_name:
                    val, _ = winreg.QueryValueEx(k, value_name)
                    return f"{value_name} = {val}"
                lines = []
                i = 0
                while True:
                    try:
                        n, v, _ = winreg.EnumValue(k, i); lines.append(f"{n} = {v}"); i += 1
                    except OSError: break
                return "\n".join(lines[:20])
        elif action == "write":
            with winreg.OpenKey(root, parts[1], 0, winreg.KEY_SET_VALUE) as k:
                winreg.SetValueEx(k, value_name, 0, winreg.REG_SZ, value)
            return f"✅ 已寫入：{value_name} = {value}"
    except Exception as e:
        return f"❌ 登錄檔操作失敗：{e}"


def execute_video_process(action, path, second=0, start=0, end=0, output=""):
    try:
        import cv2
        if action == "screenshot":
            cap = cv2.VideoCapture(path)
            cap.set(cv2.CAP_PROP_POS_MSEC, second * 1000)
            ret, frame = cap.read()
            cap.release()
            out = output or path.replace(".mp4", f"_frame{int(second)}s.jpg")
            if ret:
                cv2.imwrite(out, frame)
                return f"✅ 已擷取畫面：{out}"
            return "❌ 無法讀取影片"
        elif action == "trim":
            out = output or path.replace(".mp4", "_trim.mp4")
            subprocess.run(["ffmpeg", "-y", "-i", path, "-ss", str(start), "-to", str(end), "-c", "copy", out], capture_output=True)
            return f"✅ 已剪輯：{out}"
    except Exception as e:
        return f"❌ 影片處理失敗：{e}"


def execute_monitor_config():
    try:
        from screeninfo import get_monitors
        return "\n".join(
            f"{'主螢幕' if m.is_primary else '副螢幕'} {m.width}x{m.height} @({m.x},{m.y}) {m.name}"
            for m in get_monitors()
        )
    except Exception as e:
        return f"❌ 取得螢幕資訊失敗：{e}"


def execute_run_code(type_, code):
    try:
        if type_ == "python":
            import io as _io, traceback, contextlib
            buf = _io.StringIO()
            with contextlib.redirect_stdout(buf):
                try:
                    exec(compile(code, "<string>", "exec"), {})
                except Exception:
                    buf.write(traceback.format_exc())
            return buf.getvalue() or "（執行完畢，無輸出）"
        elif type_ == "shell":
            result = subprocess.run(
                ["powershell.exe", "-Command", code],
                capture_output=True, text=True, encoding="utf-8", errors="replace", timeout=30
            )
            return (result.stdout + result.stderr).strip() or "（執行完畢，無輸出）"
        return "未知類型"
    except Exception as e:
        return f"執行失敗：{e}"


def execute_document(action, path, content="", sheet=None):
    try:
        if action == "word_read":
            from docx import Document
            doc = Document(path)
            return "\n".join(p.text for p in doc.paragraphs if p.text.strip()) or "（文件為空）"
        elif action == "word_write":
            from docx import Document
            doc = Document() if not Path(path).exists() else Document(path)
            doc.add_paragraph(content)
            doc.save(path)
            return f"已寫入：{path}"
        elif action == "excel_read":
            import openpyxl
            wb = openpyxl.load_workbook(path)
            ws = wb[sheet] if sheet else wb.active
            rows = [[str(c.value or "") for c in row] for row in ws.iter_rows()]
            return "\n".join(["\t".join(r) for r in rows[:30]])
        elif action == "excel_write":
            import openpyxl, json
            wb = openpyxl.load_workbook(path) if Path(path).exists() else openpyxl.Workbook()
            ws = wb[sheet] if (sheet and sheet in wb.sheetnames) else wb.active
            data = json.loads(content)
            for row in data:
                ws.append(row)
            wb.save(path)
            return f"已寫入 {len(data)} 行到 {path}"
        elif action == "pdf_read":
            import fitz
            doc = fitz.open(path)
            text = "\n".join(page.get_text() for page in doc)
            return text[:3000] if text else "（PDF 無可讀文字）"
        return "未知動作"
    except Exception as e:
        return f"文件操作失敗：{e}"


def execute_web_scrape(action, url="", selector="body", interval=2.0, region="full"):
    try:
        if action == "scrape":
            from bs4 import BeautifulSoup
            res = requests.get(url, timeout=15, headers={"User-Agent": "Mozilla/5.0"})
            soup = BeautifulSoup(res.text, "html.parser")
            elements = soup.select(selector)
            return "\n".join(e.get_text(strip=True) for e in elements[:10]) or "（未找到內容）"
        elif action == "screen_diff":
            import time as t
            img1 = pyautogui.screenshot()
            t.sleep(interval)
            img2 = pyautogui.screenshot()
            import numpy as np
            a1, a2 = np.array(img1), np.array(img2)
            diff = np.abs(a1.astype(int) - a2.astype(int)).mean()
            if diff > 2.0:
                return f"⚠️ 螢幕有變化（差異度：{diff:.1f}）"
            return f"✅ 螢幕無明顯變化（差異度：{diff:.1f}）"
        return "未知動作"
    except Exception as e:
        return f"操作失敗：{e}"


def execute_image_edit(action, path, *params):
    try:
        from PIL import Image, ImageDraw, ImageFont
        img = Image.open(path)
        if action == "crop":
            x, y, w, h = int(params[0]), int(params[1]), int(params[2]), int(params[3])
            img = img.crop((x, y, x+w, y+h))
        elif action == "resize":
            w, h = int(params[0]), int(params[1])
            img = img.resize((w, h))
        elif action == "text":
            x, y = int(params[0]), int(params[1])
            text = " ".join(params[2:]) if len(params) > 2 else ""
            draw = ImageDraw.Draw(img)
            draw.text((x, y), text, fill="red")
        elif action == "merge":
            img2 = Image.open(params[0])
            merged = Image.new("RGB", (img.width + img2.width, max(img.height, img2.height)))
            merged.paste(img, (0, 0))
            merged.paste(img2, (img.width, 0))
            img = merged
        out = path.replace(".", f"_{action}.")
        img.save(out)
        return f"✅ 圖片已儲存：{out}"
    except Exception as e:
        return f"圖片編輯失敗：{e}"


def execute_cloud_storage(action, path, drive_id="root"):
    try:
        from google.oauth2.credentials import Credentials
        from googleapiclient.discovery import build
        from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
        import io as _io
        creds_path = Path("C:/Users/blue_/claude-telegram-bot/gdrive_token.json")
        if not creds_path.exists():
            return "❌ 未找到 Google Drive 憑證（gdrive_token.json）"
        creds = Credentials.from_authorized_user_file(str(creds_path))
        service = build("drive", "v3", credentials=creds)
        if action == "upload":
            media = MediaFileUpload(path)
            meta = {"name": Path(path).name, "parents": [drive_id]}
            f = service.files().create(body=meta, media_body=media, fields="id").execute()
            return f"✅ 已上傳，檔案 ID：{f.get('id')}"
        elif action == "download":
            req = service.files().get_media(fileId=drive_id)
            buf = _io.BytesIO()
            dl = MediaIoBaseDownload(buf, req)
            done = False
            while not done:
                _, done = dl.next_chunk()
            with open(path, "wb") as f:
                f.write(buf.getvalue())
            return f"✅ 已下載到：{path}"
        return "未知動作"
    except Exception as e:
        return f"雲端儲存失敗：{e}"


def execute_database(type_, db, sql, name=""):
    try:
        if type_ == "sqlite":
            conn = sqlite3.connect(db)
            cur = conn.cursor()
            cur.execute(sql)
            rows = cur.fetchall()
            conn.commit()
            conn.close()
            if rows:
                return "\n".join(str(r) for r in rows[:20])
            return "✅ 執行成功（無回傳資料）"
        elif type_ == "mysql":
            import pymysql
            conn = pymysql.connect(host=db, database=name, read_default_file="~/.my.cnf")
            cur = conn.cursor()
            cur.execute(sql)
            rows = cur.fetchall()
            conn.commit()
            conn.close()
            if rows:
                return "\n".join(str(r) for r in rows[:20])
            return "✅ 執行成功（無回傳資料）"
        return "未知類型"
    except Exception as e:
        return f"資料庫操作失敗：{e}"


def execute_encrypt_file(action, path, password):
    try:
        from cryptography.hazmat.primitives.kdf.pbkdf2 import PBKDF2HMAC
        from cryptography.hazmat.primitives import hashes
        from cryptography.fernet import Fernet
        import base64
        salt = b"xiaoniuma_salt_v1"
        kdf = PBKDF2HMAC(algorithm=hashes.SHA256(), length=32, salt=salt, iterations=100000)
        key = base64.urlsafe_b64encode(kdf.derive(password.encode()))
        f = Fernet(key)
        data = Path(path).read_bytes()
        if action == "encrypt":
            enc = f.encrypt(data)
            out = path + ".enc"
            Path(out).write_bytes(enc)
            return f"✅ 已加密：{out}"
        elif action == "decrypt":
            dec = f.decrypt(data)
            out = path.replace(".enc", ".dec")
            Path(out).write_bytes(dec)
            return f"✅ 已解密：{out}"
        return "未知動作"
    except Exception as e:
        return f"加密/解密失敗：{e}"


def execute_qr_code(action, content="", path="", duration=30.0):
    try:
        if action == "qr_gen":
            import qrcode as _qr
            img = _qr.make(content)
            save_path = path or str(Path.home() / "Desktop" / "qrcode.png")
            img.save(save_path)
            return f"✅ QR Code 已生成：{save_path}"
        elif action == "qr_scan":
            from pyzbar.pyzbar import decode
            from PIL import Image
            img = Image.open(path) if path else pyautogui.screenshot()
            results = decode(img)
            if not results:
                return "❌ 未偵測到 QR Code"
            return "\n".join(f"掃描結果：{r.data.decode('utf-8')}" for r in results)
        elif action == "clipboard_watch":
            import pyperclip, time as t
            last = pyperclip.paste()
            changes = []
            start = t.time()
            while t.time() - start < duration:
                cur = pyperclip.paste()
                if cur != last:
                    changes.append(f"[{datetime.now().strftime('%H:%M:%S')}] {cur[:100]}")
                    last = cur
                t.sleep(0.5)
            return "\n".join(changes) if changes else f"監控 {duration} 秒內無剪貼簿變化"
        return "未知動作"
    except Exception as e:
        return f"操作失敗：{e}"


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
        is_owner = update.effective_user.id == OWNER_ID
        sender_name = "于晏" if is_owner else (update.effective_user.first_name or str(update.effective_user.id))

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

        base_system = SYSTEM_PROMPT_OWNER if is_owner else SYSTEM_PROMPT_DEFAULT

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
                "ai_plan": lambda: execute_ai_plan(tool_use.input["goal"]),
                "drag": lambda: execute_drag(
                    tool_use.input["x1"], tool_use.input["y1"],
                    tool_use.input["x2"], tool_use.input["y2"],
                    tool_use.input.get("duration", 0.5)),
                "power_control": lambda: execute_power(tool_use.input["action"]),
                "virtual_desktop": lambda: execute_vdesktop(tool_use.input["action"]),
                "bluetooth": lambda: execute_bluetooth(
                    tool_use.input["action"], tool_use.input.get("mac","")),
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
                "run_code": lambda: execute_run_code(
                    tool_use.input["type"], tool_use.input["code"]),
                "document_control": lambda: execute_document(
                    tool_use.input["action"],
                    tool_use.input["path"],
                    tool_use.input.get("content",""),
                    tool_use.input.get("sheet")),
                "web_scrape": lambda: execute_web_scrape(
                    tool_use.input["action"],
                    tool_use.input.get("url",""),
                    tool_use.input.get("selector","body"),
                    tool_use.input.get("interval", 2.0),
                    tool_use.input.get("region","full")),
                "image_edit": lambda: execute_image_edit(
                    tool_use.input["action"],
                    tool_use.input["path"],
                    *tool_use.input.get("params","").split()),
                "cloud_storage": lambda: execute_cloud_storage(
                    tool_use.input["action"],
                    tool_use.input["path"],
                    tool_use.input.get("drive_id","root")),
                "database": lambda: execute_database(
                    tool_use.input["type"],
                    tool_use.input["db"],
                    tool_use.input["sql"],
                    tool_use.input.get("name","")),
                "encrypt_file": lambda: execute_encrypt_file(
                    tool_use.input["action"],
                    tool_use.input["path"],
                    tool_use.input["password"]),
                "qr_code": lambda: execute_qr_code(
                    tool_use.input["action"],
                    tool_use.input.get("content",""),
                    tool_use.input.get("path",""),
                    tool_use.input.get("duration", 30.0)),
                "screen_record": lambda: execute_screen_record(
                    tool_use.input["action"],
                    tool_use.input.get("duration", 10.0),
                    tool_use.input.get("output","")),
                "translate": lambda: execute_translate(
                    tool_use.input["text"],
                    tool_use.input.get("target","zh-TW"),
                    tool_use.input.get("source","auto")),
                "pptx_control": lambda: execute_pptx(
                    tool_use.input["action"],
                    tool_use.input["path"],
                    tool_use.input.get("slides","")),
                "api_call": lambda: execute_api_call(
                    tool_use.input["method"],
                    tool_use.input["url"],
                    tool_use.input.get("headers","{}"),
                    tool_use.input.get("body","{}")),
                "watchdog": lambda: execute_watchdog(
                    tool_use.input["process"],
                    tool_use.input["script"],
                    tool_use.input.get("duration", 60.0)),
                "ssh_sftp": lambda: execute_ssh_sftp(
                    tool_use.input["action"],
                    tool_use.input["host"],
                    tool_use.input["user"],
                    tool_use.input["password"],
                    tool_use.input.get("command",""),
                    tool_use.input.get("local",""),
                    tool_use.input.get("remote","")),
                "network_diag": lambda: execute_network_diag(
                    tool_use.input["action"],
                    tool_use.input["host"],
                    tool_use.input.get("ports","22,80,443,3306,3389,8080")),
                "win_service": lambda: execute_win_service(
                    tool_use.input["action"],
                    tool_use.input.get("name","")),
                "pdf_edit": lambda: execute_pdf_edit(
                    tool_use.input["action"],
                    tool_use.input.get("path",""),
                    tool_use.input.get("output",""),
                    tool_use.input.get("paths",""),
                    tool_use.input.get("text","")),
                "audio_process": lambda: execute_audio_process(
                    tool_use.input["action"],
                    tool_use.input["input"],
                    tool_use.input.get("output",""),
                    tool_use.input.get("start_ms",0),
                    tool_use.input.get("end_ms",0)),
                "push_notify": lambda: execute_push_notify(
                    tool_use.input["platform"],
                    tool_use.input["message"],
                    tool_use.input["webhook_or_token"]),
                "disk_backup": lambda: execute_disk_backup(
                    tool_use.input["action"],
                    tool_use.input.get("src",""),
                    tool_use.input.get("dest","")),
                "registry": lambda: execute_registry(
                    tool_use.input["action"],
                    tool_use.input["key"],
                    tool_use.input.get("value_name",""),
                    tool_use.input.get("value","")),
                "video_process": lambda: execute_video_process(
                    tool_use.input["action"],
                    tool_use.input["path"],
                    tool_use.input.get("second",0),
                    tool_use.input.get("start",0),
                    tool_use.input.get("end",0),
                    tool_use.input.get("output","")),
                "monitor_config": lambda: execute_monitor_config(),
                "email_control": lambda: execute_email_read(
                    tool_use.input["host"], tool_use.input["user"], tool_use.input["password"],
                    tool_use.input.get("folder","INBOX"), tool_use.input.get("count",5)),
                "calendar": lambda: execute_calendar(
                    tool_use.input["action"], tool_use.input.get("days",7),
                    tool_use.input.get("title",""), tool_use.input.get("start",""),
                    tool_use.input.get("end",""), tool_use.input.get("description","")),
                "global_hotkey": lambda: execute_global_hotkey(
                    tool_use.input["hotkey"], tool_use.input["command"],
                    tool_use.input.get("duration",60.0)),
                "git": lambda: execute_git(
                    tool_use.input["action"],
                    tool_use.input.get("repo","."),
                    tool_use.input.get("message",""),
                    tool_use.input.get("branch","master")),
                "hardware": lambda: execute_hardware(),
                "report": lambda: execute_report(
                    tool_use.input["title"], tool_use.input["data"],
                    tool_use.input.get("output","")),
                "dropbox": lambda: execute_dropbox(
                    tool_use.input["action"], tool_use.input["local"],
                    tool_use.input["remote"], tool_use.input.get("token","")),
                "docker": lambda: execute_docker(
                    tool_use.input["action"], tool_use.input.get("name","")),
                "pdf_image": lambda: execute_pdf_to_image(
                    tool_use.input["path"], tool_use.input.get("output_dir",""),
                    tool_use.input.get("dpi",150)),
                "barcode": lambda: execute_barcode(tool_use.input.get("image_path","")),
                "nlp": lambda: execute_nlp(
                    tool_use.input["action"], tool_use.input["text"]),
                "vpn": lambda: execute_vpn(
                    tool_use.input["action"], tool_use.input.get("name",""),
                    tool_use.input.get("user",""), tool_use.input.get("password","")),
                "restore_point": lambda: execute_restore_point(
                    tool_use.input["action"], tool_use.input.get("description","")),
                "disk_analyze": lambda: execute_disk_analyze(
                    tool_use.input.get("path","C:/"), tool_use.input.get("top",10)),
                "face_detect": lambda: execute_face_detect(
                    tool_use.input.get("image_path",""), tool_use.input.get("output","")),
                "video_gif": lambda: execute_video_gif(
                    tool_use.input["path"], tool_use.input.get("start",0),
                    tool_use.input.get("duration",5.0), tool_use.input.get("output",""),
                    tool_use.input.get("fps",10)),
                "excel_chart": lambda: execute_excel_chart(
                    tool_use.input["path"], tool_use.input["sheet"],
                    tool_use.input.get("type","bar"), tool_use.input.get("title","")),
                "speedtest": lambda: execute_speedtest(),
                "screenshot_compare": lambda: execute_screenshot_compare(
                    tool_use.input.get("img1",""), tool_use.input.get("img2",""),
                    tool_use.input.get("output","")),
                "reminder": lambda: execute_reminder(
                    tool_use.input["time"], tool_use.input["message"]),
                "webpage_shot": lambda: execute_webpage_shot(
                    tool_use.input["action"], tool_use.input["url"],
                    tool_use.input.get("selector","body"),
                    tool_use.input.get("interval",60.0), tool_use.input.get("duration",3600.0)),
                "file_tools": lambda: execute_file_tools(
                    tool_use.input["action"], tool_use.input["path"],
                    tool_use.input.get("dest",""), tool_use.input.get("pattern",""),
                    tool_use.input.get("replacement",""), tool_use.input.get("ext","")),
                "image_tools": lambda: execute_image_tools(
                    tool_use.input["action"], tool_use.input.get("path",""),
                    tool_use.input.get("quality",75), tool_use.input.get("width",0),
                    tool_use.input.get("height",0), tool_use.input.get("target_lang","zh-TW")),
                "lookup": lambda: execute_lookup(
                    tool_use.input["action"], tool_use.input.get("ip",""),
                    tool_use.input.get("amount",1.0), tool_use.input.get("from_cur","USD"),
                    tool_use.input.get("to_cur","TWD")),
                "system_tools": lambda: execute_system_tools(
                    tool_use.input["action"], **{k:v for k,v in tool_use.input.items() if k!="action"}),
                "tts_advanced": lambda: execute_tts_advanced(
                    tool_use.input["action"], tool_use.input.get("text",""),
                    tool_use.input.get("voice","zh-TW-YunJheNeural")),
                "todo_list": lambda: execute_todo(
                    tool_use.input["action"], tool_use.input.get("task",""),
                    tool_use.input.get("id",0)),
                "password_mgr": lambda: execute_password_mgr(
                    tool_use.input["action"], tool_use.input["site"],
                    tool_use.input["master"], tool_use.input.get("username",""),
                    tool_use.input.get("password","")),
                "clipboard_image": lambda: execute_clipboard_image(
                    tool_use.input["action"], tool_use.input.get("path","")),
            }

            if tool_use.name == "send_voice":
                import asyncio
                loop = asyncio.get_running_loop()
                text = tool_use.input["text"]
                voice = tool_use.input.get("voice", "zh-TW-YunJheNeural")
                await update.message.reply_chat_action("record_voice")
                try:
                    ogg_data = await loop.run_in_executor(None, generate_voice_ogg, text, voice)
                    import io as _io
                    await update.message.reply_voice(voice=_io.BytesIO(ogg_data))
                    reply = f"🔊 語音訊息已傳送"
                except Exception as e:
                    reply = f"語音生成失敗：{e}\n\n{text}"
                    await update.message.reply_text(reply)
                save_message(chat_id, "assistant", text)
                return

            elif tool_use.name == "sysres_chart":
                import asyncio
                loop = asyncio.get_running_loop()
                duration = tool_use.input.get("duration", 10)
                await update.message.reply_text(f"⏳ 監控 {duration} 秒中...")
                out_path = await loop.run_in_executor(None, execute_sysres_chart, duration)
                if out_path and Path(out_path).exists():
                    with open(out_path, "rb") as f:
                        await update.message.reply_photo(photo=f, caption="系統資源使用率")
                    reply = f"資源圖表已生成"
                else:
                    reply = out_path
                save_message(chat_id, "assistant", reply)
                await update.message.reply_text(reply)
                return

            elif tool_use.name == "chart":
                import asyncio
                loop = asyncio.get_running_loop()
                out_path = await loop.run_in_executor(None, execute_chart,
                    tool_use.input["type"],
                    tool_use.input["data"],
                    tool_use.input.get("title",""),
                    tool_use.input.get("output",""))
                if out_path and Path(out_path).exists():
                    with open(out_path, "rb") as f:
                        await update.message.reply_photo(photo=f, caption=tool_use.input.get("title","圖表"))
                    reply = f"圖表已生成：{out_path}"
                else:
                    reply = out_path
                save_message(chat_id, "assistant", reply)
                await update.message.reply_text(reply)
                return

            elif tool_use.name == "screen_stream":
                duration = tool_use.input.get("duration", 10)
                interval = tool_use.input.get("interval", 2)
                import asyncio
                loop = asyncio.get_running_loop()
                await update.message.reply_text(f"📹 開始串流 {duration} 秒...")
                screenshots = await loop.run_in_executor(None, execute_screen_stream, duration, interval)
                for i, img_bytes in enumerate(screenshots, 1):
                    await update.message.reply_photo(photo=img_bytes, caption=f"📸 {i}/{len(screenshots)}")
                reply = f"串流完成，共傳送 {len(screenshots)} 張截圖"
                save_message(chat_id, "assistant", reply)
                await update.message.reply_text(reply)
                return

            elif tool_use.name in simple_tools:
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
