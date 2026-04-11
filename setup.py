import getpass

print("=== Telegram Bot 設定 ===")
telegram_token = input("請輸入 Telegram Bot Token: ").strip()
anthropic_key = input("請輸入 Anthropic API Key: ").strip()

with open(".env", "w", encoding="utf-8") as f:
    f.write(f"TELEGRAM_BOT_TOKEN={telegram_token}\n")
    f.write(f"ANTHROPIC_API_KEY={anthropic_key}\n")

print("✓ .env 檔案已儲存！")
