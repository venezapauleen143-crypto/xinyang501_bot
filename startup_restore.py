"""
啟動還原腳本
確保開機或重開終端機後，所有設定與服務自動恢復：
  1. 從 GitHub 拉取最新程式碼
  2. 自動安裝所有必要套件
  3. 確保 Bot 在跑（含殭屍清除）
  4. 確認並建立所有排程任務
"""
import subprocess
import sys
import io
import json
from pathlib import Path

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

PYTHONW    = r"C:\Users\blue_\AppData\Local\Microsoft\WindowsApps\pythonw3.12.exe"
PYTHON     = r"C:\Users\blue_\AppData\Local\Microsoft\WindowsApps\python3.12.exe"
BOT_SCRIPT = r"C:\Users\blue_\claude-telegram-bot\bot.py"
BOT_DIR    = r"C:\Users\blue_\claude-telegram-bot"
REPO_PATH  = r"C:\Users\blue_\xinyang501_bot"
ENSURE_BOT = r"C:\Users\blue_\ensure_bot_running.py"

# ── 所有排程任務 ──────────────────────────────────────────────────────────────
# 格式：名稱 → (觸發類型, 時間/事件, 腳本路徑, 說明)
# 觸發類型：Daily=每天固定時間  Logon=登入時  Hourly=每小時
REQUIRED_TASKS = {
    # 每天固定時間（WakeToRun = 可喚醒電腦）
    "ClaudeDailyLearn":   ("Daily",  "21:00", r"C:\Users\blue_\claude_daily_learn.py",  "每日自學"),
    "WakeForGoodMorning": ("Daily",  "10:55", ENSURE_BOT,                               "早安前確認 Bot"),
    "WakeForDailyLearn":  ("Daily",  "20:55", ENSURE_BOT,                               "自學前確認 Bot"),
    "WakeForGoodNight":   ("Daily",  "22:25", ENSURE_BOT,                               "晚安前確認 Bot"),
    "WakeForNoonCheck":   ("Daily",  "12:55", ENSURE_BOT,                               "午間確認 Bot"),
    "WakeForEveningCheck":("Daily",  "17:55", ENSURE_BOT,                               "傍晚確認 Bot"),
    # 登入時自動啟動 Bot（不需要喚醒，用 Logon 觸發）
    "BotStartOnLogin":    ("Logon",  None,    ENSURE_BOT,                               "登入時啟動 Bot"),
    # 每小時健康檢查（確保 Bot 沒有崩潰）
    "BotHourlyHealthCheck":("Hourly", None,   ENSURE_BOT,                               "每小時確認 Bot 存活"),
}

# ── 必要 Python 套件 ──────────────────────────────────────────────────────────
REQUIRED_PACKAGES = [
    # Bot 核心
    "python-telegram-bot[job-queue]", "anthropic", "python-dotenv",
    # TTS / STT
    "edge-tts", "SpeechRecognition", "sounddevice", "pydub",
    # 桌面自動化
    "pyautogui", "pywinauto", "pynput", "keyboard", "mss",
    # 系統資訊
    "psutil", "GPUtil", "speedtest-cli", "wmi",
    # 檔案/版本控制
    "watchdog", "gitpython", "imapclient",
    # 媒體/圖像
    "pillow", "imageio[ffmpeg]", "opencv-python",
    # 音訊控制
    "pycaw", "comtypes",
    # 網路/雲端
    "dropbox", "docker", "requests",
    # 通知
    "win10toast",
    # 模板
    "jinja2",
]

# 這些套件較重或有特殊安裝步驟，單獨處理
OPTIONAL_PACKAGES = [
    "easyocr",  # 需要 PyTorch，較大
]

def log(msg: str):
    print(msg)

def is_bot_running() -> bool:
    """精確比對 bot.py 是否在執行"""
    result = subprocess.run(
        ["powershell.exe", "-Command",
         "Get-WmiObject Win32_Process | Where-Object { $_.Name -like 'pythonw*' } "
         "| Select-Object CommandLine | ConvertTo-Json -Compress"],
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    try:
        data = json.loads(result.stdout.strip())
        if isinstance(data, dict):
            data = [data]
        return any("bot.py" in (p.get("CommandLine") or "") for p in data)
    except Exception:
        return False

def start_bot():
    subprocess.Popen(
        [PYTHONW, BOT_SCRIPT],
        cwd=BOT_DIR,
        creationflags=0x00000008
    )
    log("✅ Bot 已啟動")

def get_existing_tasks() -> set:
    result = subprocess.run(
        ["powershell.exe", "-Command",
         "Get-ScheduledTask | Select-Object TaskName | ConvertTo-Json -Compress"],
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    tasks = set()
    try:
        data = json.loads(result.stdout.strip())
        if isinstance(data, dict):
            data = [data]
        for t in data:
            tasks.add(t.get("TaskName", "").strip())
    except Exception:
        pass
    return tasks

def create_task(name: str, trigger_type: str, time_hhmm, script: str, desc: str):
    # 所有排程統一最強設定：WakeToRun + StartWhenAvailable + 失敗重試3次
    _SETTINGS = ("New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries "
                 "-StartWhenAvailable -WakeToRun -ExecutionTimeLimit (New-TimeSpan -Days 0) "
                 "-RestartCount 3 -RestartInterval (New-TimeSpan -Minutes 1)")
    if trigger_type == "Daily":
        trigger_ps = f"New-ScheduledTaskTrigger -Daily -At '{time_hhmm}'"
    elif trigger_type == "Logon":
        trigger_ps = "New-ScheduledTaskTrigger -AtLogOn -User $env:USERNAME"
    elif trigger_type == "Hourly":
        trigger_ps = ("New-ScheduledTaskTrigger -Once -At '00:00' "
                      "-RepetitionInterval (New-TimeSpan -Hours 1) "
                      "-RepetitionDuration ([TimeSpan]::MaxValue)")
    else:
        return
    settings_ps = _SETTINGS

    ps = (
        f"$action = New-ScheduledTaskAction -Execute '{PYTHONW}' -Argument '{script}';"
        f"$trigger = {trigger_ps};"
        f"$settings = {settings_ps};"
        f"Register-ScheduledTask -TaskName '{name}' -Action $action "
        f"-Trigger $trigger -Settings $settings -Description '{desc}' -Force | Out-Null"
    )
    result = subprocess.run(
        ["powershell.exe", "-Command", ps],
        capture_output=True, text=True, encoding="utf-8", errors="replace"
    )
    if result.returncode == 0:
        label = f"每天 {time_hhmm}" if trigger_type == "Daily" else trigger_type
        log(f"✅ 排程 [{name}] 已建立（{label}）")
    else:
        log(f"⚠️ 排程 [{name}] 建立失敗：{result.stderr.strip()[:200]}")

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

def install_packages():
    log("正在檢查並安裝必要套件...")
    pkgs = " ".join(f'"{p}"' for p in REQUIRED_PACKAGES)
    result = subprocess.run(
        ["powershell.exe", "-Command",
         f"& '{PYTHON}' -m pip install {pkgs} --quiet --exists-action i 2>&1"],
        capture_output=True, text=True, encoding="utf-8", errors="replace",
        timeout=300
    )
    if result.returncode == 0:
        log("✅ 所有必要套件已安裝")
    else:
        log(f"⚠️ 部分套件安裝有警告（通常無礙）")

    # imageio ffmpeg backend
    subprocess.run(
        [PYTHON, "-c", "import imageio; imageio.plugins.freeimage.download()"],
        capture_output=True, timeout=60
    )

def main():
    log("=" * 50)
    log("  啟動還原中...")
    log("=" * 50)

    # 1. 拉取最新程式碼
    log("\n[1] 從 GitHub 拉取最新版本...")
    pull_latest()

    # 2. 安裝套件
    log("\n[2] 檢查必要套件...")
    install_packages()

    # 3. 確保 bot 在跑
    log("\n[3] 確認 Bot 狀態...")
    if is_bot_running():
        log("✅ Bot 已在執行中")
    else:
        log("⚠️ Bot 未執行，正在啟動...")
        start_bot()

    # 4. 確認所有排程存在
    log("\n[4] 確認排程任務...")
    existing = get_existing_tasks()
    for name, (trigger_type, time_hhmm, script, desc) in REQUIRED_TASKS.items():
        if name in existing:
            log(f"✅ [{name}] 已存在")
        else:
            log(f"⚠️ [{name}] 不存在，正在建立...")
            create_task(name, trigger_type, time_hhmm, script, desc)

    log("\n✅ 所有服務已就緒！")
    log("=" * 50)

if __name__ == "__main__":
    main()
