"""
小牛馬終端機雙向同步
- 即時顯示 Telegram 上的所有對話
- 輸入訊息直接發送到 Telegram（以機器人身份）
"""
import os
import sys
import time
import threading
import requests
from pathlib import Path
from dotenv import load_dotenv

load_dotenv(Path(__file__).parent / ".env")

TOKEN = os.getenv("TELEGRAM_BOT_TOKEN")
OWNER_ID = 8362721681
LOG_FILE = Path(__file__).parent / "messages.log"
API_URL = f"https://api.telegram.org/bot{TOKEN}"

# 目標 chat_id（預設私訊于晏哥，可在執行時指定）
target_chat_id = OWNER_ID
if len(sys.argv) > 1:
    try:
        target_chat_id = int(sys.argv[1])
    except ValueError:
        pass


def tail_log():
    """即時追蹤 messages.log，有新訊息就印出"""
    if not LOG_FILE.exists():
        LOG_FILE.touch()
    with open(LOG_FILE, "r", encoding="utf-8") as f:
        f.seek(0, 2)  # 跳到檔案結尾，只看新訊息
        while True:
            line = f.readline()
            if line:
                sys.stdout.write("\r" + line)
                sys.stdout.write("你: ")
                sys.stdout.flush()
            else:
                time.sleep(0.1)


def send_message(chat_id: int, text: str):
    """透過 Bot API 發送訊息到 Telegram"""
    try:
        res = requests.post(
            f"{API_URL}/sendMessage",
            json={"chat_id": chat_id, "text": text},
            timeout=10
        )
        if not res.ok:
            print(f"[發送失敗] {res.status_code}: {res.text}")
    except Exception as e:
        print(f"[發送錯誤] {e}")


def main():
    global target_chat_id
    print("=" * 50)
    print("  小牛馬終端機  ──  雙向同步")
    print(f"  發送目標：chat_id {target_chat_id}")
    print("  輸入訊息後按 Enter 發送到 Telegram")
    print("  輸入 /quit 離開 | /id <chat_id> 切換對象")
    print("=" * 50)
    print()

    # 背景執行 log 監聽
    t = threading.Thread(target=tail_log, daemon=True)
    t.start()
    while True:
        try:
            sys.stdout.write("你: ")
            sys.stdout.flush()
            msg = input()

            if not msg.strip():
                continue
            if msg.lower() == "/quit":
                print("掰掰！")
                break
            if msg.lower().startswith("/id "):
                try:
                    target_chat_id = int(msg.split()[1])
                    print(f"[已切換到 chat_id: {target_chat_id}]")
                except ValueError:
                    print("[格式錯誤，例：/id -1001234567890]")
                continue

            send_message(target_chat_id, msg)

        except (KeyboardInterrupt, EOFError):
            print("\n掰掰！")
            break


if __name__ == "__main__":
    main()
