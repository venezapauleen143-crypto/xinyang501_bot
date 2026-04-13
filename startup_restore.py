"""
啟動還原腳本
確保開機或重開終端機後，所有設定與服務自動恢復
"""
import subprocess
import sys
import io
from pathlib import Path

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

SCHTASKS = "C:\\Windows\\System32\\schtasks.exe"
PYTHONW = r"C:\Users\blue_\AppData\Local\Python\bin\pythonw.exe"
BOT_SCRIPT = r"C:\Users\blue_\claude-telegram-bot\bot.py"
BOT_DIR = r"C:\Users\blue_\claude-telegram-bot"
REPO_PATH = r"C:\Users\blue_\xinyang501_bot"

REQUIRED_TASKS = {
    "ClaudeDailyLearn":   ("21:00", r"C:\Users\blue_\claude_daily_learn.py"),
    "WakeForGoodMorning": ("10:55", r"C:\Users\blue_\ensure_bot_running.py"),
    "WakeForDailyLearn":  ("20:55", r"C:\Users\blue_\ensure_bot_running.py"),
    "WakeForGoodNight":   ("22:25", r"C:\Users\blue_\ensure_bot_running.py"),
}

def log(msg):
    print(msg)

def is_bot_running():
    result = subprocess.run(
        ["powershell.exe", "-Command",
         "Get-Process pythonw -ErrorAction SilentlyContinue | Measure-Object | Select-Object -ExpandProperty Count"],
        capture_output=True, text=True
    )
    count = result.stdout.strip()
    return count.isdigit() and int(count) > 0

def start_bot():
    subprocess.Popen(["pythonw", BOT_SCRIPT], cwd=BOT_DIR)
    log("✅ Bot 已啟動")

def get_existing_tasks():
    result = subprocess.run(
        [SCHTASKS, "/Query", "/FO", "CSV", "/NH"],
        capture_output=True, text=True, encoding="cp950", errors="replace"
    )
    tasks = set()
    for line in result.stdout.strip().splitlines():
        parts = line.strip('"').split('","')
        if parts:
            tasks.add(parts[0].replace("\\", "").strip())
    return tasks

def create_task(name, time_hhmm, script):
    ps = (
        f"$action = New-ScheduledTaskAction -Execute '{PYTHONW}' -Argument '{script}';"
        f"$trigger = New-ScheduledTaskTrigger -Daily -At '{time_hhmm}';"
        f"$settings = New-ScheduledTaskSettingsSet -WakeToRun -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries;"
        f"Register-ScheduledTask -TaskName '{name}' -Action $action -Trigger $trigger -Settings $settings -Force | Out-Null"
    )
    subprocess.run(["powershell.exe", "-Command", ps], capture_output=True)
    log(f"✅ 排程 [{name}] 已建立（每天 {time_hhmm}，可喚醒）")

def pull_latest():
    result = subprocess.run(
        ["git", "-C", REPO_PATH, "pull"],
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    if "Already up to date" in result.stdout:
        log("✅ GitHub 已是最新版本")
    elif result.returncode == 0:
        log(f"✅ GitHub 已更新：{result.stdout.strip()}")
    else:
        log(f"⚠️ GitHub pull 失敗：{result.stderr.strip()}")

def main():
    log("=" * 45)
    log("  啟動還原中...")
    log("=" * 45)

    # 1. 拉取最新程式碼
    log("\n[1] 從 GitHub 拉取最新版本...")
    pull_latest()

    # 2. 確保 bot 在跑
    log("\n[2] 確認 Bot 狀態...")
    if is_bot_running():
        log("✅ Bot 已在執行中")
    else:
        log("⚠️ Bot 未執行，正在啟動...")
        start_bot()

    # 3. 確認所有排程存在
    log("\n[3] 確認排程任務...")
    existing = get_existing_tasks()
    for name, (time_hhmm, script) in REQUIRED_TASKS.items():
        if name in existing:
            log(f"✅ [{name}] 已存在")
        else:
            log(f"⚠️ [{name}] 不存在，正在建立...")
            create_task(name, time_hhmm, script)

    log("\n✅ 所有服務已就緒！")
    log("=" * 45)

if __name__ == "__main__":
    main()
