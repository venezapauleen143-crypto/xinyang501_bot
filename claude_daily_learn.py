"""
Claude Code 每日自學腳本
每天晚上 9 點與小牛馬同步學習相同主題
涵蓋所有技能：Wave 1-13 + 進階主題
"""
import os
import sys
import io
import datetime
import subprocess
from pathlib import Path
from dotenv import load_dotenv

if sys.stdout.encoding != 'utf-8':
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8', errors='replace')

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))

from anthropic import Anthropic

client = Anthropic()

TOPICS = [
    # ── Python 基礎與進階 ─────────────────────────────────────────────
    "Python asyncio 進階用法與最佳實踐",
    "Python 錯誤處理與 logging 最佳實踐",
    "Python 型別提示（Type Hints）與 mypy 靜態檢查",
    "Python dataclass 與 pydantic 資料模型",
    "Python 效能分析與優化：cProfile、line_profiler",
    "Python 多執行緒與多程序：threading vs multiprocessing",
    "Python subprocess 進階：管道、串流、逾時控制",
    "Python pathlib 與檔案系統操作最佳實踐",
    "Python 正則表達式進階技巧",
    "Python 裝飾器（Decorator）深入解析",

    # ── Telegram Bot ─────────────────────────────────────────────────
    "Telegram Bot API 進階功能：InlineKeyboard、CallbackQuery",
    "Telegram Bot 訊息限速與 flood control 處理",
    "Telegram Bot 群組管理與權限控制",
    "python-telegram-bot JobQueue 排程任務管理",
    "Telegram Bot 語音訊息：STT 辨識與回覆",
    "Telegram Bot 檔案收發：圖片、影片、文件",
    "Telegram Bot 狀態機（ConversationHandler）設計",

    # ── Claude / AI API ──────────────────────────────────────────────
    "Claude API Tool Use（工具使用）進階設計模式",
    "Claude API 串流回應（Streaming）實作",
    "Claude API System Prompt 工程技巧",
    "Claude API 多輪對話記憶管理策略",
    "大型語言模型 Token 計算與成本優化",

    # ── 桌面自動化 ───────────────────────────────────────────────────
    "pyautogui 桌面自動化：滑鼠、鍵盤、截圖",
    "pywinauto Windows GUI 自動化進階技巧",
    "pynput 全域快捷鍵與滑鼠監聽",
    "keyboard 模組：快捷鍵錄製與重播",
    "mss 高效螢幕截圖與區域擷取",

    # ── 視覺辨識 / OCR ───────────────────────────────────────────────
    "OpenCV 圖像處理：模板匹配與特徵偵測",
    "easyOCR 文字辨識：中英文混合場景",
    "Pillow 圖像操作：裁切、比較、像素分析",
    "螢幕像素監控：顏色比對與變化偵測",

    # ── 音訊處理 ─────────────────────────────────────────────────────
    "edge-tts 語音合成：聲音選擇與參數調整",
    "SpeechRecognition + sounddevice 語音輸入",
    "pycaw Windows 音量控制與音訊裝置管理",
    "pydub 音訊格式轉換與剪輯",

    # ── 系統管理 ─────────────────────────────────────────────────────
    "psutil 系統資源監控：CPU、記憶體、磁碟、網路",
    "Windows 防火牆規則管理（netsh advfirewall）",
    "Windows 程序管理：啟動、終止、優先級",
    "Windows 電源計畫管理（powercfg）",
    "Windows 事件日誌查詢（Get-WinEvent）",
    "Windows 排程任務管理（Task Scheduler）進階",
    "Windows Defender 設定與排除清單管理",
    "Windows 系統資訊查詢：WMI 與 PowerShell",

    # ── 檔案系統 ─────────────────────────────────────────────────────
    "watchdog 檔案系統監控：即時偵測變化",
    "difflib 文字差異比較與 diff 輸出",
    "Python zipfile、tarfile 壓縮解壓縮",
    "檔案雜湊計算：MD5、SHA256 完整性驗證",

    # ── 網路 ─────────────────────────────────────────────────────────
    "requests 進階：Session、重試、代理設定",
    "imapclient IMAP 收信與郵件解析",
    "speedtest-cli 網路速度測試自動化",
    "Wake-on-LAN 網路喚醒實作",
    "FTP 自動化：ftplib 上傳下載",
    "Socket 程式設計：TCP/UDP 基礎到進階",

    # ── 瀏覽器自動化 ─────────────────────────────────────────────────
    "Playwright 瀏覽器自動化：頁面操作、截圖、PDF",
    "Playwright 網頁爬蟲：動態內容擷取",
    "Playwright 多標籤頁與視窗管理",
    "Chrome DevTools Protocol 進階技巧",

    # ── 雲端 / 容器 ──────────────────────────────────────────────────
    "Dropbox API：檔案上傳、下載、分享連結",
    "Docker SDK for Python：容器管理自動化",
    "OneDrive / SharePoint REST API 整合",
    "Git 自動化：gitpython 操作倉庫",

    # ── 硬體 / 周邊 ──────────────────────────────────────────────────
    "OpenCV 攝影機控制：截圖、錄影、即時串流",
    "GPUtil GPU 監控：溫度、使用率、記憶體",
    "多螢幕管理：分辨率、排列、設定",
    "印表機管理：列印、查詢狀態",
    "USB 裝置偵測與管理",
    "Wi-Fi 網路掃描與熱點建立",
    "ADB Android 設備遠端控制",

    # ── Windows 子系統 ───────────────────────────────────────────────
    "WSL2 與 Python 互通：執行 Linux 指令",
    "Hyper-V 虛擬機管理自動化",
    "Windows 遠端桌面（RDP）連接自動化",
    "Windows 代理伺服器設定（系統 / 瀏覽器）",

    # ── 資料處理 ─────────────────────────────────────────────────────
    "SQLite 效能優化與索引設計",
    "pandas 資料清洗與分析技巧",
    "json / csv 大檔案串流處理",
    "Jinja2 模板引擎進階用法",

    # ── 安全性 / 加密 ────────────────────────────────────────────────
    "Python hashlib 與 secrets 安全隨機數",
    "Python cryptography 套件：對稱/非對稱加密",
    "Windows 憑證管理與 DPAPI",

    # ── 監控 / 告警 ──────────────────────────────────────────────────
    "自動化告警系統：條件觸發 → Telegram 通知",
    "間隔排程自動化：APScheduler vs asyncio",
    "螢幕文字等待：OCR + 輪詢策略",
    "像素變化監控：遊戲/應用程式自動化",
]

LOG_PATH  = Path("C:/Users/blue_/xinyang501_bot/learning_log.md")
REPO_PATH = Path("C:/Users/blue_/xinyang501_bot")

def main():
    today     = datetime.date.today()
    today_str = str(today)
    topic     = TOPICS[today.toordinal() % len(TOPICS)]

    print(f"[{today_str}] 學習主題：{topic}")

    response = client.messages.create(
        model="claude-sonnet-4-6",
        max_tokens=1500,
        system=(
            "你是 Claude Code，一個整合在 Windows 開發環境中的 AI 助理。"
            "你擁有超過 100 個自動化技能，涵蓋桌面控制、語音、視覺、網路、雲端、系統管理等。"
            "每天晚上你會主動學習一個新技能並做成筆記，重點放在如何在實際自動化工作中應用這個技能。"
        ),
        messages=[{
            "role": "user",
            "content": (
                f"今天的學習主題是：{topic}。\n"
                "請整理出：\n"
                "1. 核心概念與工作原理\n"
                "2. 實用程式碼範例（可直接使用）\n"
                "3. 常見坑與注意事項\n"
                "4. 與其他已掌握技能的整合應用場景\n"
                "用繁體中文條列式整理，格式清晰易查閱。"
            )
        }]
    )
    learn_content = response.content[0].text

    existing = LOG_PATH.read_text(encoding="utf-8") if LOG_PATH.exists() else ""
    entry = f"\n\n## {today_str} — {topic}（Claude Code 筆記）\n\n{learn_content}\n\n---"
    LOG_PATH.write_text(existing + entry, encoding="utf-8")
    print(f"筆記已儲存到 {LOG_PATH}")

    subprocess.run(["git", "-C", str(REPO_PATH), "add", "learning_log.md"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    subprocess.run(["git", "-C", str(REPO_PATH), "commit", "-m",
                    f"Claude Code 學習筆記：{today_str} {topic}"],
                   stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    result = subprocess.run(["git", "-C", str(REPO_PATH), "push"],
                            stdout=subprocess.DEVNULL, stderr=subprocess.PIPE,
                            text=True, encoding="utf-8")
    status = "已上傳 GitHub" if result.returncode == 0 else "GitHub 上傳失敗"
    print(f"[{status}]")

if __name__ == "__main__":
    main()
