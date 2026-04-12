"""
Claude Code 每日自學腳本
每天晚上 9 點與小牛馬同步學習相同主題
"""
import os
import sys
import io
import datetime
import subprocess
from pathlib import Path
from dotenv import load_dotenv

# 強制 UTF-8
if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))

from anthropic import Anthropic

client = Anthropic()

TOPICS = [
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

LOG_PATH = Path("C:/Users/blue_/xinyang501_bot/learning_log.md")
REPO_PATH = Path("C:/Users/blue_/xinyang501_bot")

def main():
    today = datetime.date.today()
    today_str = str(today)
    topic = TOPICS[today.toordinal() % len(TOPICS)]

    print(f"[{today_str}] 學習主題：{topic}")

    # 用 Claude API 生成學習筆記（從 Claude Code 工具助理的視角）
    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1500,
        system="你是 Claude Code，一個整合在開發環境中的 AI 助理。每天晚上你會主動學習一個新技能並做成筆記，重點放在如何在實際開發工作中應用這個技能。",
        messages=[{"role": "user", "content": f"今天的學習主題是：{topic}。請整理出重點知識、實用範例程式碼（如果適用）、以及你學到的心得，特別說明這個技能如何幫助開發工作，用繁體中文條列式整理，格式清晰。"}]
    )
    learn_content = response.content[0].text

    # 儲存到 learning_log.md（與小牛馬共用同一份）
    existing = LOG_PATH.read_text(encoding="utf-8") if LOG_PATH.exists() else ""
    entry = f"\n\n## {today_str} — {topic}（Claude Code 筆記）\n\n{learn_content}\n\n---"
    LOG_PATH.write_text(existing + entry, encoding="utf-8")
    print(f"筆記已儲存到 {LOG_PATH}")

    # push 到 GitHub
    subprocess.run(["git", "-C", str(REPO_PATH), "add", "learning_log.md"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    subprocess.run(["git", "-C", str(REPO_PATH), "commit", "-m", f"Claude Code 學習筆記：{today_str} {topic}"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    result = subprocess.run(["git", "-C", str(REPO_PATH), "push"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.PIPE, text=True, encoding="utf-8")
    status = "已上傳 GitHub" if result.returncode == 0 else "GitHub 上傳失敗"
    print(f"[{status}]")

if __name__ == "__main__":
    main()
