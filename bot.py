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

        system = SYSTEM_PROMPT_OWNER if update.effective_user.id == OWNER_ID else SYSTEM_PROMPT_DEFAULT

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

            if tool_use.name == "get_stock":
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
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Bot 已啟動...")
    app.run_polling()
