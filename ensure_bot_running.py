"""
確保小牛馬 Bot 正在執行，若未執行則自動啟動
支援精確比對 bot.py 腳本、清除殭屍/重複程序、記錄 log
"""
import subprocess
import sys
import io
import datetime
from pathlib import Path

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

PYTHONW   = r"C:\Users\blue_\AppData\Local\Microsoft\WindowsApps\pythonw3.12.exe"
BOT_SCRIPT = r"C:\Users\blue_\claude-telegram-bot\bot.py"
BOT_DIR    = r"C:\Users\blue_\claude-telegram-bot"
LOG_PATH   = Path(r"C:\Users\blue_\xinyang501_bot\bot_restart.log")

def log(msg: str):
    ts = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"[{ts}] {msg}"
    print(line)
    try:
        with open(LOG_PATH, "a", encoding="utf-8") as f:
            f.write(line + "\n")
    except Exception:
        pass

def get_bot_pids() -> list[int]:
    """取得所有執行 bot.py 的 pythonw PID"""
    result = subprocess.run(
        ["powershell.exe", "-Command",
         "Get-WmiObject Win32_Process | Where-Object { $_.Name -like 'pythonw*' } "
         "| Select-Object ProcessId, CommandLine | ConvertTo-Json -Compress"],
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    pids = []
    try:
        import json
        data = json.loads(result.stdout.strip())
        if isinstance(data, dict):
            data = [data]
        for proc in data:
            cmd = proc.get("CommandLine") or ""
            if "bot.py" in cmd:
                pids.append(int(proc["ProcessId"]))
    except Exception:
        pass
    return pids

def kill_pids(pids: list[int]):
    for pid in pids:
        subprocess.run(
            ["powershell.exe", "-Command", f"Stop-Process -Id {pid} -Force -ErrorAction SilentlyContinue"],
            capture_output=True
        )

def start_bot():
    subprocess.Popen(
        [PYTHONW, BOT_SCRIPT],
        cwd=BOT_DIR,
        creationflags=0x00000008  # DETACHED_PROCESS
    )
    log("✅ Bot 已啟動")

if __name__ == "__main__":
    pids = get_bot_pids()

    if len(pids) > 1:
        log(f"⚠️ 發現 {len(pids)} 個 bot 程序（重複），清除後重啟...")
        kill_pids(pids)
        start_bot()
    elif len(pids) == 1:
        log(f"✅ Bot 已在執行中（PID {pids[0]}）")
    else:
        log("⚠️ Bot 未執行，正在啟動...")
        start_bot()
