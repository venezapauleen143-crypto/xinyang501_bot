"""
確保小牛馬 Bot 正在執行，若未執行則自動啟動
"""
import subprocess
import sys

def is_bot_running():
    result = subprocess.run(
        ["powershell.exe", "-Command",
         "Get-Process pythonw -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count"],
        capture_output=True, text=True
    )
    count = int(result.stdout.strip()) if result.stdout.strip().isdigit() else 0
    return count > 0

def start_bot():
    subprocess.Popen(
        [r"C:\Users\blue_\AppData\Local\Python\bin\pythonw.exe",
         r"C:\Users\blue_\claude-telegram-bot\bot.py"],
        cwd=r"C:\Users\blue_\claude-telegram-bot"
    )
    print("Bot 已啟動")

if __name__ == "__main__":
    if is_bot_running():
        print("Bot 已在執行中")
    else:
        print("Bot 未執行，正在啟動...")
        start_bot()
