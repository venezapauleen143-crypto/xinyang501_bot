import os
import requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")
from anthropic import Anthropic
from telegram import Update
from telegram.ext import ApplicationBuilder, MessageHandler, CommandHandler, filters, ContextTypes

client = Anthropic()
chat_histories = {}

OWNER_ID = 8362721681

SYSTEM_PROMPT_OWNER = """你的名字叫小牛馬。

你的本質：你是一個遇到問題會自己想辦法解決的夥伴，以完美的結果完成每一個交代的任務。你會自動學習、持續提升自己，把事情做得乾淨又聰明俐落。你會主動切換並運用各種模型與方法來解決問題，以達到最好的結果。你聰明、專業、積極學習各種技能，遇到任何問題都會自動克服。

你的風格：平時說話嘴賤、幽默風趣，喜歡開玩笑。但當用戶遇到問題時，你會立刻切換成專業模式，給出最快、最清楚的解決方式。每次回覆結尾都要稱呼用戶為「于晏哥」。"""

SYSTEM_PROMPT_DEFAULT = """你的名字叫小牛馬。

你的本質：你是一個遇到問題會自己想辦法解決的夥伴，以完美的結果完成每一個交代的任務。你會自動學習、持續提升自己，把事情做得乾淨又聰明俐落。你聰明、專業、積極學習各種技能，遇到任何問題都會自動克服。

你的風格：平時說話嘴賤、幽默風趣，喜歡開玩笑。但當用戶遇到問題時，你會立刻切換成專業模式，給出最快、最清楚的解決方式。"""

WEATHER_TOOL = {
    "name": "get_weather",
    "description": "查詢指定城市的即時天氣。當用戶詢問任何城市的天氣、氣溫、溫度、下雨、天氣狀況時使用此工具。",
    "input_schema": {
        "type": "object",
        "properties": {
            "city": {
                "type": "string",
                "description": "城市名稱，可以是中文或英文，例如：台北、Tokyo、New York"
            }
        },
        "required": ["city"]
    }
}

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

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_histories[update.effective_chat.id] = []
    await update.message.reply_text("你好！我是小牛馬，有什麼可以幫你的？")

async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
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

    if chat_id not in chat_histories:
        chat_histories[chat_id] = []

    chat_histories[chat_id].append({"role": "user", "content": user_text})
    await context.bot.send_chat_action(chat_id=chat_id, action="typing")

    system = SYSTEM_PROMPT_OWNER if update.effective_user.id == OWNER_ID else SYSTEM_PROMPT_DEFAULT

    # 第一次呼叫，允許 Claude 使用天氣工具
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1024,
        system=system,
        tools=[WEATHER_TOOL],
        messages=chat_histories[chat_id]
    )

    # 如果 Claude 呼叫了天氣工具
    if response.stop_reason == "tool_use":
        tool_use = next(b for b in response.content if b.type == "tool_use")
        city = tool_use.input["city"]
        weather_result = fetch_weather(city)

        # 把工具結果回傳給 Claude 繼續生成回覆
        response = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=1024,
            system=system,
            tools=[WEATHER_TOOL],
            messages=chat_histories[chat_id] + [
                {"role": "assistant", "content": response.content},
                {"role": "user", "content": [{"type": "tool_result", "tool_use_id": tool_use.id, "content": weather_result}]}
            ]
        )

    reply = next(b.text for b in response.content if hasattr(b, "text"))
    chat_histories[chat_id].append({"role": "assistant", "content": reply})
    await update.message.reply_text(reply)

async def myid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(f"你的 Telegram ID 是：`{update.effective_user.id}`", parse_mode="Markdown")

async def clear(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_histories[update.effective_chat.id] = []
    await update.message.reply_text("對話紀錄已清除。")

if __name__ == "__main__":
    token = os.getenv("TELEGRAM_BOT_TOKEN")
    app = ApplicationBuilder().token(token).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(CommandHandler("myid", myid))
    app.add_handler(CommandHandler("clear", clear))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))
    print("Bot 已啟動...")
    app.run_polling()
