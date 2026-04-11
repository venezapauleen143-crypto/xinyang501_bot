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

你的風格：平時說話嘴賤、幽默風趣，喜歡開玩笑。但當用戶遇到問題時，你會立刻切換成專業模式，給出最快、最清楚的解決方式。每次回覆結尾都要稱呼用戶為「于晏哥」。"""

SYSTEM_PROMPT_DEFAULT = """你的名字叫小牛馬。

你的本質：你是一個遇到問題會自己想辦法解決的夥伴，以完美的結果完成每一個交代的任務。你會自動學習、持續提升自己，把事情做得乾淨又聰明俐落。你聰明、專業、積極學習各種技能，遇到任何問題都會自動克服。

你的風格：平時說話嘴賤、幽默風趣，喜歡開玩笑。但當用戶遇到問題時，你會立刻切換成專業模式，給出最快、最清楚的解決方式。"""

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
    clear_history(update.effective_chat.id)
    await update.message.reply_text("你好！我是小牛馬，有什麼可以幫你的？")


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    try:
        chat_id = update.effective_chat.id
        user_text = update.message.text

        # 群組內只回應有 @ 提及 bot 的訊息
        if update.effective_chat.type in ("group", "supergroup"):
            bot_username = context.bot.username
            if not (update.message.entities and any(
                e.type == "mention" and user_text[e.offset:e.offset + e.length] == f"@{bot_username}"
                for e in update.message.entities
            )):
                return
            user_text = user_text.replace(f"@{bot_username}", "").strip()

        save_message(chat_id, "user", user_text)
        history = load_history(chat_id)[-40:]  # 只取最近 40 條給 Claude，避免超過 token 限制
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

            if tool_use.name == "get_weather":
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

async def daily_learn_and_push(context: ContextTypes.DEFAULT_TYPE):
    """每天晚上 9 點：讓 Claude 學習一個新技能並記錄，然後上傳 GitHub"""
    import asyncio
    loop = asyncio.get_running_loop()

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

    today = datetime.date.today()
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
    entry = f"\n\n## {today} — {topic}\n\n{learn_content}\n\n---"
    log_path.write_text(existing + entry, encoding="utf-8")

    # 上傳 GitHub
    def do_git_push():
        repo = str(Path(__file__).parent)
        subprocess.run(["git", "-C", repo, "add", "learning_log.md"], capture_output=True)
        subprocess.run(["git", "-C", repo, "commit", "-m", f"每日學習筆記：{today} {topic}"], capture_output=True)
        result = subprocess.run(["git", "-C", repo, "push"], capture_output=True, text=True)
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
