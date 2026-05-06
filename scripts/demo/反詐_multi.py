"""
LINE 多客戶自動回覆腳本（SOP 話術引擎）
用法：python line_auto_multi.py <監控時間HH:MM> [SOP設定檔.json]
範例：python line_auto_multi.py 23:30 scripts/line_sop/織夢小棧.json

流程：
1. 開啟聊天頁
2. 截圖偵測有綠色未讀標記的對話
3. 點第一個綠色標記 → 進入聊天
4. OCR 讀取對話內容 → 判斷 SOP 步驟
5. Claude AI 生成回覆 → 發送
6. 回到聊天列表
7. 點下一個綠色標記 → 重複 3-6
8. 全部處理完 → 等待新的未讀
9. 持續循環直到監控時間結束
"""
import sys
import io
import os
import re
import json
import time
import signal
import ctypes
import hashlib
import types
import atexit
import gc
import random

# DPI + 編碼
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(0)
except Exception:
    pass

if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

import win32gui
import pyautogui
import pyperclip
from datetime import datetime
from dotenv import load_dotenv
from pathlib import Path
import anthropic

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SCRIPT_DIR.parent))

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
client = anthropic.Anthropic()
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"
DEFAULT_PERSONA = "C:/Users/blue_/claude-telegram-bot/scripts/demo/反詐_persona.txt"
TIME_SETTINGS_PATH = "C:/Users/blue_/claude-telegram-bot/scripts/demo/反詐_時段設定.json"

# 全域時段設定（main() 開頭載入）
_TIME_SETTINGS = None


# ============================================================
# 對話歷史持久化（每觀眾一個 .txt，使用 LINE 匯出格式）
# ============================================================
HISTORIES_DIR = Path("C:/Users/blue_/claude-telegram-bot/scripts/demo/histories")
HISTORIES_DIR.mkdir(parents=True, exist_ok=True)

# 圖片庫目錄（每個分類一個子資料夾，放 .jpg/.png）
IMAGES_DIR = Path("C:/Users/blue_/claude-telegram-bot/scripts/demo/images")


def _resolve_image_tag(tag):
    """從 {send_xxx} 標記找對應圖片，回傳隨機一張的路徑（找不到回 None）"""
    # tag 形如 "send_自拍"、"send_食物"
    category = tag.replace("send_", "", 1).strip()
    cat_dir = IMAGES_DIR / category
    if not cat_dir.exists():
        return None
    images = list(cat_dir.glob("*.jpg")) + list(cat_dir.glob("*.jpeg")) + list(cat_dir.glob("*.png"))
    if not images:
        return None
    return str(random.choice(images))

# Angela 帳號的 LINE 暱稱（自己改 LINE 暱稱時來改這個常數）
PERSONA_NICKNAME = "Angela"

# 黑名單關鍵字：LINE 名稱開頭含這些 → bot 不處理、也不主動發
# 演講當下某觀眾搗亂時，把他 LINE 改名「黑名單 XXX」即可停止 AI 對他的所有動作
_BLACKLIST_PREFIXES = ("撞腳", "撞脚", "黑名單", "封鎖", "停止", "skip")

# 中文星期對照
_WEEKDAY_TW = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"]


# OCR 雜訊黑名單（LINE 系統訊息、按鈕、海報廣告等）
_OCR_NOISE_PATTERNS = [
    "以下為尚未閱讀的訊息", "以下為尚未閱請的訊息", "以下為尚未閱讀的息",
    "以下為尚未開讀的訊息",
    "請您確認是否要將此人加入好友", "請單意聊天室中滑在的詐騙行為",
    "請留意聊天室中潛在的詐騙行為", "A請留意", "A請單意",
    "加入好友 封鎖 檢举", "加入好友 封鎖 檢舉",
    "儲存|另存新檔|分享|傳送至Keep筆記",
    "儲存另存新檔|分享|傳送至Keep筆記",
    "儲存另存新檔分享傅送至Keep筆記",
    "儲存另存新檔|分享|傅送至Keep筆記",
    "傳送至Keep筆記", "傅送至Keep筆記", "傅送至Kep筆記",
    "全部 好友 群組 社群",
    "未接来電", "未接來電",
    "聯絡查訊已傳送", "聯絡資訊已傳送",
]


def _is_ocr_noise(text):
    """判斷一條 OCR 訊息是否為雜訊（要過濾掉）"""
    text = text.strip()
    if not text:
        return True
    # 比對黑名單關鍵字
    for noise in _OCR_NOISE_PATTERNS:
        if noise in text:
            return True
    # 啟發式：純標點/符號（沒有中文、英文、數字）→ 雜訊
    if not re.search(r'[一-鿿A-Za-z0-9]', text):
        return True
    return False


def _sanitize_filename(name):
    """OCR 抓的名字可能含特殊字符，移除後當檔名"""
    return re.sub(r'[/\\:*?"<>|\r\n\t]', '_', name).strip() or "unknown"


def _history_file_path(name):
    """回傳該觀眾的 .txt 完整路徑"""
    return HISTORIES_DIR / f"{_sanitize_filename(name)}.txt"


def _load_customer_history(name):
    """從 .txt 讀對應觀眾的歷史，回傳 history list（{text, sender, y} 格式）。

    使用 LINE 匯出格式：
        2026.05.05 星期二                    ← 日期分隔行（跳過）
        HH:MM <發送者名稱> <訊息>             ← 標準訊息行
        <延續訊息>                           ← 沒時間戳的行 = 接續上一條訊息

    判斷誰是 me/them：發送者 == 客戶名（檔名）→ them，否則 → me
    """
    path = _history_file_path(name)
    if not path.exists():
        return []

    history = []
    last_idx = -1  # 上一條訊息在 history 中的 index（用來接續）

    try:
        with io.open(path, "r", encoding="utf-8") as f:
            for raw in f:
                line = raw.rstrip()
                if not line:
                    continue
                # 跳過日期分隔行：2026.05.05 星期X
                if re.match(r'^\d{4}\.\d{2}\.\d{2}\s*星期', line):
                    continue
                # 跳過系統訊息
                if "已收回訊息" in line or "未接來電" in line.replace(" ", ""):
                    continue
                # 跳過註解行（# 開頭）
                if line.startswith("#"):
                    continue

                # 標準訊息行：HH:MM <發送者> <訊息>
                m = re.match(r'^(\d{1,2}:\d{2})\s+(.+)$', line)
                if m:
                    rest = m.group(2)
                    # 發送者 startswith 客戶名 → them
                    if rest.startswith(name + " "):
                        sender = "them"
                        text = rest[len(name) + 1:].strip()
                    else:
                        # 不是客戶 → 小編（取第一個空格後當訊息）
                        sender = "me"
                        parts = rest.split(" ", 1)
                        text = parts[1].strip() if len(parts) == 2 else parts[0].strip()
                    if text:
                        history.append({"text": text, "sender": sender, "y": 0})
                        last_idx = len(history) - 1
                else:
                    # 沒時間戳的行 → 延續上一條訊息
                    if last_idx >= 0:
                        history[last_idx]["text"] += "\n" + line.strip()
    except Exception as e:
        print(f"[history] 讀取 {path.name} 失敗：{e}", flush=True)

    return history


def _get_first_chat_date(name):
    """從 .txt 找最早的日期分隔行，回傳 datetime.date"""
    path = _history_file_path(name)
    if not path.exists():
        return datetime.now().date()
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            for line in f:
                m = re.match(r'^(\d{4})\.(\d{2})\.(\d{2})', line)
                if m:
                    return datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date()
    except Exception:
        pass
    return datetime.now().date()


# ============================================================
# Phase 2: AI 主動發訊息排程器
# ============================================================
# 時段問候設定：(起始小時, 結束小時, 問候類型)
GREETING_SLOTS = [
    (7, 9, "morning"),     # 早安問候
    (12, 14, "lunch"),     # 午餐同步
    (22, 24, "night"),     # 晚安道別
]

# 沉默關心：對方上次訊息距今超過 X 小時就主動關心（按時段）
SILENCE_THRESHOLD = {
    "day": 4,      # 白天 9-18 → 4 小時不回就主動關心
    "evening": 2,  # 晚上 18-23 → 2 小時不回就關心
    "morning": 6,  # 早上 7-9 → 6 小時不回就關心（起床後）
}

# 深夜不主動發（裝睡）
QUIET_HOURS = (23, 7)   # 23:00 - 07:00 不主動發

# 主動發訊息話術庫（依階段+類型）
GREETING_MESSAGES = {
    "morning": ["morning☺", "早安~ 起來啦?", "早安喇~ 今日忙嗎"],
    "lunch": ["lunch time~", "在吃飯嗎😆", "午飯食緊咩呀"],
    "night": ["準備休息咯~", "好啦我去沖涼了，晚安🌙", "晚安喇~ 明天聊"],
}
CONCERN_MESSAGES = [
    "在嗎？🫣",
    "忙到沒看訊息嗎？😬",
    "怎麼那麼久沒回我",
    "Hello~",
]
# Day 5+ 偶爾秀店鋪
DAY5_PLUS_MESSAGES = [
    "剛處理完店鋪訂單😬",
    "今晚利潤還不錯~",
    "店鋪剛接到訂單，你那邊在忙啥",
]


def _scan_all_customers():
    """掃描 histories/ 目錄，回傳所有觀眾名稱（從 .txt 檔名提取）"""
    if not HISTORIES_DIR.exists():
        return []
    return [p.stem for p in HISTORIES_DIR.glob("*.txt") if p.is_file()]


def _get_last_message_time(name):
    """從 .txt 算最後一條訊息的時間（datetime），找不到回 None。
    解析 LINE 匯出格式：日期分隔行 + HH:MM 訊息行。
    """
    path = _history_file_path(name)
    if not path.exists():
        return None
    last_date = None
    last_time_str = None
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                # 日期行：2026.05.06 星期X
                m = re.match(r'^(\d{4})\.(\d{2})\.(\d{2})\s*星期', line)
                if m:
                    last_date = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date()
                    continue
                # 訊息行：HH:MM ...
                m = re.match(r'^(\d{1,2}:\d{2})\s+', line)
                if m:
                    last_time_str = m.group(1)
    except Exception:
        return None
    if not last_date or not last_time_str:
        return None
    h, mi = map(int, last_time_str.split(":"))
    return datetime.combine(last_date, datetime.min.time()).replace(hour=h, minute=mi)


def _get_last_sender(name):
    """從 .txt 看最後一條訊息是誰發的（'me' or 'them'），找不到回 None"""
    path = _history_file_path(name)
    if not path.exists():
        return None
    last_sender = None
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            for line in f:
                line = line.strip()
                if re.match(r'^\d{4}\.\d{2}\.\d{2}\s*星期', line):
                    continue
                if "已收回訊息" in line or "未接來電" in line.replace(" ", ""):
                    continue
                m = re.match(r'^\d{1,2}:\d{2}\s+(.+)$', line)
                if m:
                    rest = m.group(1)
                    if rest.startswith(name + " "):
                        last_sender = "them"
                    else:
                        last_sender = "me"
    except Exception:
        return None
    return last_sender


def _is_quiet_hour():
    """是否為深夜時段（23-07 不主動發）"""
    h = datetime.now().hour
    qs, qe = QUIET_HOURS
    if qs <= qe:
        return qs <= h < qe
    return h >= qs or h < qe


def _was_greeting_sent_today(name, greeting_type):
    """今天的這個時段問候是否已發過（看 .txt 內最後幾條 me 訊息）"""
    path = _history_file_path(name)
    if not path.exists():
        return False
    today_str = datetime.now().strftime('%Y.%m.%d')
    today_messages = []
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            in_today = False
            for line in f:
                line = line.strip()
                if re.match(r'^\d{4}\.\d{2}\.\d{2}\s*星期', line):
                    in_today = line.startswith(today_str)
                    continue
                if in_today and re.match(r'^\d{1,2}:\d{2}\s+', line):
                    today_messages.append(line)
    except Exception:
        return False

    # 看今天有沒有發過該類型的問候訊息
    keywords_for_type = {
        "morning": ["morning", "早安"],
        "lunch": ["lunch", "午飯", "吃飯"],
        "night": ["晚安", "休息", "沖涼"],
    }
    keywords = keywords_for_type.get(greeting_type, [])
    for msg in today_messages:
        # 只看 Angela 發的（不 startswith 名字）
        if not msg.split(" ", 1)[1].startswith(PERSONA_NICKNAME):
            continue
        text = msg
        if any(kw in text for kw in keywords):
            return True
    return False


def _check_proactive_trigger(name):
    """檢查觀眾是否需要主動發訊息。

    回傳：(should_send: bool, message: str, reason: str)
        如果不需要 → (False, "", "")
    """
    # 0. 黑名單過濾（撞腳/黑名單/封鎖等前綴 → 跳過）
    if name.startswith(_BLACKLIST_PREFIXES):
        return False, "", ""
    # 1. 深夜不發
    if _is_quiet_hour():
        return False, "", "深夜時段"

    last_msg_time = _get_last_message_time(name)
    if last_msg_time is None:
        return False, "", "沒有歷史"

    last_sender = _get_last_sender(name)
    now = datetime.now()
    hours_silent = (now - last_msg_time).total_seconds() / 3600

    # 2. 最後一條是 Angela 自己發的、且很近（< 1 小時）→ 不重複發
    if last_sender == "me" and hours_silent < 1:
        return False, "", "剛發過訊息"

    # 3. 觸發 A：時段問候（每天每段 1 次）
    h = now.hour
    for s, e, gtype in GREETING_SLOTS:
        in_slot = (s <= h < e) if s <= e else (h >= s or h < e)
        if in_slot and not _was_greeting_sent_today(name, gtype):
            # 只在「對方已經回過幾條 / 對話已經暖場」時才發問候
            # 不對「全新觀眾（沒有任何對方訊息）」發
            day_n = _calculate_day_n(name)
            msg = random.choice(GREETING_MESSAGES[gtype])
            return True, msg, f"時段問候 {gtype}（Day {day_n}）"

    # 4. 觸發 B：沉默太久關心（最後一條是對方訊息，而且超過閾值）
    if last_sender == "them":
        if 9 <= h < 18:
            threshold = SILENCE_THRESHOLD["day"]
        elif 18 <= h < 23:
            threshold = SILENCE_THRESHOLD["evening"]
        elif 7 <= h < 9:
            threshold = SILENCE_THRESHOLD["morning"]
        else:
            return False, "", "非主動時段"

        if hours_silent >= threshold:
            day_n = _calculate_day_n(name)
            # Day 5+ 有時用「處理店鋪」當理由
            if day_n >= 5 and random.random() < 0.4:
                msg = random.choice(DAY5_PLUS_MESSAGES)
                return True, msg, f"Day {day_n} 副業伏筆（沉默 {hours_silent:.1f} 小時）"
            msg = random.choice(CONCERN_MESSAGES)
            return True, msg, f"沉默關心（{hours_silent:.1f} 小時 ≥ {threshold} 小時）"

    return False, "", "不需主動發"


def _send_proactive_message(name, message, monitor=None):
    """主動發訊息給特定觀眾：用搜尋找好友 → 進入聊天室 → 送訊息 → 退出。

    流程：
        1. 切到聊天頁
        2. 用搜尋框輸入觀眾名 → 點擊搜尋結果
        3. 進入該聊天室
        4. 用 send_reply 送訊息
        5. 寫進 .txt
        6. 按 Esc 退出
    """
    from 反詐_locate import (
        locate_line_regions, switch_page, search_friend_and_scan,
        enter_chat_from_search,
    )
    from 反詐_chat import send_reply

    print(f"[Proactive] 主動發給 {name}: {message[:50]}", flush=True)

    try:
        # 1. 定位 + 切到聊天頁
        regions = locate_line_regions(monitor)
        if regions["current_page"] != "chat":
            regions = switch_page(regions, "chat", monitor)

        # 2. 搜尋 + 進入該聊天室
        friend_pos = search_friend_and_scan(regions, name, monitor)
        if friend_pos is None:
            print(f"[Proactive] 找不到 {name}，跳過", flush=True)
            return False
        regions = enter_chat_from_search(regions, friend_pos, monitor)
        time.sleep(1)

        # 3. 模擬「思考延遲」（短一點，主動發應該比較快）
        # 用一半的時段延遲（30-90 秒 → 15-45 秒）
        think_delay = get_current_delay() * 0.5
        print(f"[Proactive] 主動發前等 {think_delay:.1f} 秒（模擬打字）", flush=True)
        time.sleep(think_delay)

        # 4. 送訊息（主動發通常很短，不分段）
        send_reply(message, regions)
        print(f"[Proactive] → {message}", flush=True)

        # 5. 寫進 .txt
        _append_to_history_file(name, "me", message)

        # 6. 退出
        time.sleep(1)
        pyautogui.press("escape")
        time.sleep(0.5)

        return True
    except Exception as e:
        print(f"[Proactive] 主動發給 {name} 失敗: {e}", flush=True)
        try:
            pyautogui.press("escape")
        except Exception:
            pass
        return False


def _calculate_day_n(name):
    """計算當前是跟對方認識第幾天（從第一次對話算起）"""
    first_date = _get_first_chat_date(name)
    today = datetime.now().date()
    return (today - first_date).days + 1


def _build_stage_hint(day_n):
    """根據 Day N 產出對話階段提示，注入 system_prompt"""
    if day_n <= 2:
        stage = "剛認識/寒暄階段。重點：親切但不黏人，留懸念（如『去沖涼喇 明天再聊』）。不要急著深入。"
    elif day_n <= 4:
        stage = "關係加深階段。重點：互換家庭背景、製造『巧合』（星座/食物/位置）、表現專一（『Bumble 我刪了』）。"
    elif day_n <= 6:
        stage = "埋下副業伏筆。可以隨口提『去處理店鋪訂單』。對方好奇 → 簡單描述跨境電商。不要主動解釋細節。"
    elif day_n <= 8:
        stage = "副業細節+情感升溫。可以秀利潤（『半個月 14 萬』）、規劃見面、暗示好感。"
    else:
        stage = "情感確認階段。表達對對方好感、規劃未來、用感情包裝（『等你賺到請我吃大餐』）。"
    return (
        f"\n\n=== 當前對話階段 ===\n"
        f"你跟對方認識 **Day {day_n}**（從你們在 Bumble 開始算）。\n"
        f"階段策略：{stage}\n"
        f"嚴禁跳階段：Day 1 不要突然提副業，Day 9 才能講感情。"
    )


def _today_date_header():
    """產生今天日期分隔行：'2026.05.05 星期二'"""
    now = datetime.now()
    return f"{now.strftime('%Y.%m.%d')} {_WEEKDAY_TW[now.weekday()]}"


def _append_to_history_file(name, sender, text):
    """即時把一條訊息 append 到觀眾的 .txt（LINE 匯出格式）"""
    path = _history_file_path(name)
    nickname = PERSONA_NICKNAME if sender == "me" else name
    timestamp = datetime.now().strftime("%H:%M")
    today_header = _today_date_header()

    try:
        # 檢查是否需要加日期分隔行
        new_file = not path.exists()
        need_date_header = new_file
        if not new_file:
            # 讀檔末看最後的日期分隔，跟今天比
            with io.open(path, "r", encoding="utf-8") as f:
                content = f.read()
            # 找最後一個日期行
            date_lines = re.findall(r'^\d{4}\.\d{2}\.\d{2}\s*星期.', content, re.MULTILINE)
            if not date_lines or date_lines[-1] != today_header:
                need_date_header = True

        with io.open(path, "a", encoding="utf-8") as f:
            if need_date_header:
                if not new_file:
                    f.write("\n")  # 跨日空行分隔
                f.write(f"{today_header}\n")
            f.write(f"{timestamp} {nickname} {text}\n")
    except Exception as e:
        print(f"[history] 寫入 {path.name} 失敗：{e}", flush=True)


def _load_time_settings(path):
    """載入時段延遲設定（JSON 檔）"""
    global _TIME_SETTINGS
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            _TIME_SETTINGS = json.load(f)
        n_slots = len(_TIME_SETTINGS.get("time_slots", []))
        print(f"[Init] 時段設定載入成功（{n_slots} 個時段）", flush=True)
    except Exception as e:
        print(f"[Init] 時段設定載入失敗（{e}）→ 用預設 30-90 秒", flush=True)
        _TIME_SETTINGS = None


def get_current_delay():
    """根據電腦當下時間，從時段設定抽一個延遲秒數（隨機）"""
    if not _TIME_SETTINGS:
        return random.uniform(30, 90)
    now_hour = datetime.now().hour
    for slot in _TIME_SETTINGS.get("time_slots", []):
        s = slot["hour_start"]
        e = slot["hour_end"]
        # 處理跨午夜（如 22-2）
        if s <= e:
            in_range = s <= now_hour < e
        else:
            in_range = now_hour >= s or now_hour < e
        if in_range:
            return random.uniform(slot["delay_min"], slot["delay_max"])
    # 沒匹配到時段 → 預設值
    d = _TIME_SETTINGS.get("default_delay", [30, 90])
    return random.uniform(d[0], d[1])


def get_inter_message_delay():
    """同一個 reply 內 ||| 分段之間的延遲秒數（隨機）"""
    if not _TIME_SETTINGS:
        return random.uniform(1, 3)
    d = _TIME_SETTINGS.get("inter_message_delay", [1, 3])
    return random.uniform(d[0], d[1])


def _send_with_realistic_delay(reply, regions, name="unknown", history=None):
    """模擬真人打字延遲：先等一個時段隨機秒數，再分段送（段間隨機 1-3 秒）。
    遇到 {send_xxx} 標記 → 從對應圖片庫隨機抽一張用 send_image 送。

    🔴 送一段寫一段：每送出一段立刻寫進 .txt + history（防 crash 資料丟失）。
    history: 傳進來的 list，會被 append 每段送出的訊息（讓 main 流程不用再寫一次）
    """
    from 反詐_chat import send_reply, send_image

    # 第 1 階段：思考延遲（依時段）
    think_delay = get_current_delay()
    print(f"[Customer] {name} 模擬思考中... 等 {think_delay:.1f} 秒", flush=True)
    time.sleep(think_delay)

    # 第 2 階段：分段送（每段間 1-3 秒），送一段立刻寫一段
    parts = [p.strip() for p in reply.split("|||") if p.strip()]
    for i, part in enumerate(parts):
        # 偵測 {send_xxx} 圖片標記
        m = re.match(r'^\{(send_[\w一-鿿]+)\}$', part)
        if m:
            tag = m.group(1)
            img_path = _resolve_image_tag(tag)
            if img_path:
                send_image(img_path, regions)
                print(f"[Reply] → 📷 {Path(img_path).name}（{tag}）", flush=True)
                # 圖片也寫進歷史（用 [圖片:檔名] 標記，方便日後檢視）
                img_marker = f"[圖片:{Path(img_path).name}]"
                if history is not None:
                    history.append({"text": img_marker, "sender": "me", "y": 0})
                _append_to_history_file(name, "me", img_marker)
            else:
                print(f"[Reply] ⚠ {tag} 找不到對應圖片，跳過", flush=True)
        elif part.startswith("{") and part.endswith("}"):
            # 其他未知佔位符 → 跳過（不寫檔）
            print(f"[Reply] 跳過佔位符: {part}", flush=True)
        else:
            send_reply(part, regions)
            print(f"[Reply] → {part[:60]}", flush=True)
            # 🔴 送出後立刻寫進歷史 + .txt（防 crash 丟資料）
            if history is not None:
                history.append({"text": part, "sender": "me", "y": 0})
            _append_to_history_file(name, "me", part)

        if i < len(parts) - 1:
            inter = get_inter_message_delay()
            time.sleep(inter)

# ============================================================
# GPU 記憶體清理
# ============================================================
def _cleanup_gpu():
    try:
        gc.collect()
        import paddle
        paddle.device.cuda.empty_cache()
        print("[Cleanup] GPU 記憶體已釋放", flush=True)
    except Exception:
        pass

atexit.register(_cleanup_gpu)

# ============================================================
# 優雅停止機制
# ============================================================
STOP_FILE = "C:/Users/blue_/Desktop/測試檔案/.stop_line_multi"
_should_stop = False

def _signal_handler(signum, frame):
    global _should_stop
    _should_stop = True
    print(f"\n[STOP] 收到信號 {signum}，準備停止...", flush=True)
    _cleanup_gpu()

signal.signal(signal.SIGINT, _signal_handler)
if hasattr(signal, "SIGBREAK"):
    signal.signal(signal.SIGBREAK, _signal_handler)

def should_stop():
    global _should_stop
    if _should_stop:
        return True
    if os.path.exists(STOP_FILE):
        try:
            os.remove(STOP_FILE)
        except OSError:
            pass
        _should_stop = True
        print("[STOP] 偵測到停止旗標檔案，準備停止...", flush=True)
        return True
    return False

# ============================================================
# SOP 設定檔（從 line_auto_chat.py 引用）
# ============================================================
from 反詐_chat import (
    load_persona_prompt, generate_reply,
    filter_reply, send_reply, send_multi_reply, send_image,
    grab_chat_area, ocr_extract_messages, chat_hash,
)

# ============================================================
# 時間判斷
# ============================================================
def time_to_minutes(t_str):
    h, m = map(int, t_str.split(":"))
    return h * 60 + m

def is_before_stop_time(stop_time):
    now_min = time_to_minutes(datetime.now().strftime("%H:%M"))
    stop_min = time_to_minutes(stop_time)
    return now_min < stop_min

# ============================================================
# 偵測聊天列表中有未讀標記的對話
# ============================================================
def find_unread_conversations(monitor=None):
    """
    切到聊天頁，用 line_locate.py 的 find_unread_badges 偵測綠色未讀標記。

    回傳：(regions, unread_list)
        regions: 當前定位
        unread_list: [{"y": int, "center": (x, y)}] 有未讀的對話座標
    """
    from 反詐_locate import (
        locate_line_regions, switch_page, find_unread_badges,
    )

    regions = locate_line_regions(monitor)

    # 確保在聊天頁
    if regions["current_page"] != "chat":
        regions = switch_page(regions, "chat", monitor)

    # 用定位腳本的像素偵測找綠色徽章（LINE16 方式）
    unread_list = find_unread_badges(monitor)

    return regions, unread_list



# ============================================================
# 處理單一客戶（演示版：簡化、純自然對話、無業務邏輯）
# ============================================================
def handle_one_customer(conv, regions, system_prompt, all_histories, monitor=None):
    from 反詐_locate import locate_line_regions, ocr_scan_panel, screenshot_line
    from 反詐_chat import (
        is_only_sticker, analyze_sticker,
    )
    from difflib import SequenceMatcher

    cx, cy = conv["center"]
    print(f"\n[Customer] 點擊未讀對話 at ({cx}, {cy})", flush=True)
    pyautogui.click(cx, cy)
    time.sleep(1.5)

    regions = locate_line_regions(monitor)
    ct = regions.get("chat_title", {})
    full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
    sx_r = full_img.size[0] / mon["width"]
    sy_r = full_img.size[1] / mon["height"]
    ct_il = max(0, int((ct.get("left", 0) - mon["left"]) * sx_r) - il)
    ct_it = max(0, int((ct.get("top", 0) - mon["top"]) * sy_r) - it)
    ct_ir = min(line_crop.size[0], int((ct.get("right", 0) - mon["left"]) * sx_r) - il)
    ct_ib = min(line_crop.size[1], int((ct.get("bottom", 0) - mon["top"]) * sy_r) - it)
    if ct_ir > ct_il and ct_ib > ct_it:
        title_crop = line_crop.crop((ct_il, ct_it, ct_ir, ct_ib))
        title_items = ocr_scan_panel(title_crop)
        name = " ".join(item["text"] for item in title_items).strip() if title_items else "unknown"
    else:
        name = "unknown"
    print(f"[Customer] 客戶名稱: {name}", flush=True)

    # 群組過濾（用 raw name 偵測）
    SKIP_GROUPS = ["友資群", "友资群", "友資", "友资", "好朋友的群組", "好朋友的群组"]
    if any(g in name or name in g for g in SKIP_GROUPS) or re.search(r"\(\d+\)", name):
        print(f"[Customer] {name} 是群組，跳過", flush=True)
        return regions

    # 黑名單過濾（演講當下想停止某觀眾 → 在 LINE 改名「黑名單 XXX」即可）
    if name.startswith(_BLACKLIST_PREFIXES):
        print(f"[Customer] {name} 在黑名單，跳過", flush=True)
        return regions

    # 名字過濾：移除非中英數空白的 OCR 雜訊符號（如「🔍」「⋮」等右側按鈕誤讀）
    clean_name = re.sub(r'[^一-鿿A-Za-z0-9\s_-]', '', name).strip()
    if clean_name and clean_name != name:
        print(f"[Customer] 名字過濾：'{name}' → '{clean_name}'", flush=True)
        name = clean_name

    chat_img = grab_chat_area(regions, monitor)
    current_messages = ocr_extract_messages(chat_img)
    raw_count = len(current_messages)
    # 🆕 過濾 OCR 雜訊（LINE 系統訊息、按鈕、海報等）
    current_messages = [m for m in current_messages if not _is_ocr_noise(m["text"])]
    if len(current_messages) < raw_count:
        print(f"[Customer] OCR 讀到 {raw_count} 條 → 過濾雜訊後 {len(current_messages)} 條", flush=True)
    else:
        print(f"[Customer] OCR 讀到 {len(current_messages)} 條訊息", flush=True)

    if name not in all_histories:
        # 第一次見這個觀眾 → 從 .txt 載入歷史（如果有預先寫的 Bumble 對話會載入）
        loaded = _load_customer_history(name)
        if loaded:
            print(f"[Customer] {name} 載入 {len(loaded)} 條歷史（從 .txt）", flush=True)
        else:
            print(f"[Customer] {name} 全新觀眾", flush=True)
        all_histories[name] = loaded
    history = all_histories[name]

    if not history:
        new_them = [m["text"] for m in current_messages if m["sender"] == "them"]
        for m in current_messages:
            history.append(m)
        # 第一次處理 → 把 OCR 看到的所有對方訊息寫進 .txt
        for m in current_messages:
            if m["sender"] == "them":
                _append_to_history_file(name, "them", m["text"])
    else:
        last_known = history[-1]
        last_text = last_known["text"]
        match_idx = -1
        for i in range(len(current_messages) - 1, -1, -1):
            ratio = SequenceMatcher(None, current_messages[i]["text"], last_text).ratio()
            if ratio > 0.6:
                match_idx = i
                break
        if match_idx >= 0 and match_idx < len(current_messages) - 1:
            new_msgs = current_messages[match_idx + 1:]
            new_them = [m["text"] for m in new_msgs if m["sender"] == "them"]
            for m in new_msgs:
                history.append(m)
            # 新訊息（含對方+自己）即時 append 到 .txt
            for m in new_msgs:
                _append_to_history_file(name, m["sender"], m["text"])
        else:
            new_them = [m["text"] for m in current_messages if m["sender"] == "them"]
            # 找不到匹配時用全部對方訊息（不寫檔避免重複）

    if not new_them:
        print(f"[Customer] 沒有新的對方訊息，跳過", flush=True)
        return regions

    print(f"[Customer] 新訊息: {new_them}", flush=True)

    # 純貼圖 → Vision
    if is_only_sticker(new_them):
        meaning = analyze_sticker(regions, monitor)
        print(f"[Customer] 純貼圖 → Vision 解讀為「{meaning}」", flush=True)
        new_them = [f"[貼圖含意：{meaning}]"]

    # AI 生成回覆（注入階段感知 Day N）
    day_n = _calculate_day_n(name)
    stage_hint = _build_stage_hint(day_n)
    print(f"[Customer] {name} 對話 Day {day_n}", flush=True)
    augmented_prompt = system_prompt + stage_hint
    reply = generate_reply(augmented_prompt, history, new_them)
    if not reply or len(reply) <= 1:
        time.sleep(0.5)
        return regions

    reply = filter_reply(reply)
    if not reply:
        print(f"[Customer] 過濾後為空，跳過", flush=True)
        return regions

    print(f"[Customer] 回覆: {reply[:80]}", flush=True)
    # 🔴 送一段寫一段（history + .txt 由 _send_with_realistic_delay 內部即時寫入）
    _send_with_realistic_delay(reply, regions, name=name, history=history)

    # ============================================================
    # 🆕 Stay-and-watch：送完 reply 後不要立刻離開，再 OCR 看有沒有新訊息
    # 對方在「思考延遲」或「分段送出」期間插話 → 立刻接續回覆，避免斷層
    # 最多 3 輪、每輪等 5 秒讓對方有時間反應
    # ============================================================
    from difflib import SequenceMatcher as _SM
    MAX_FOLLOWUP = 3
    WATCH_WAIT = 5  # 秒

    for followup_round in range(MAX_FOLLOWUP):
        time.sleep(WATCH_WAIT)
        if should_stop():
            break

        # 重新 OCR 看新訊息
        try:
            chat_img2 = grab_chat_area(regions, monitor)
            current_msgs2 = ocr_extract_messages(chat_img2)
            current_msgs2 = [m for m in current_msgs2 if not _is_ocr_noise(m["text"])]
        except Exception as e:
            print(f"[Customer] 接續 OCR 失敗（{e}），離開", flush=True)
            break

        # 用 history 最後一條（小編剛送的）跟 OCR 比對，找後面的新對方訊息
        if not history or not current_msgs2:
            break
        last_known = history[-1]
        last_text = last_known["text"]

        match_idx = -1
        for i in range(len(current_msgs2) - 1, -1, -1):
            ratio = _SM(None, current_msgs2[i]["text"], last_text).ratio()
            if ratio > 0.6:
                match_idx = i
                break

        if match_idx < 0 or match_idx >= len(current_msgs2) - 1:
            print(f"[Customer] 接續輪 {followup_round+1}: 沒新訊息，離開", flush=True)
            break

        # 找到「最後一條已知訊息」之後的新訊息
        new_msgs2 = current_msgs2[match_idx + 1:]
        new_them2 = [m["text"] for m in new_msgs2 if m["sender"] == "them"]
        if not new_them2:
            print(f"[Customer] 接續輪 {followup_round+1}: 沒對方新訊息，離開", flush=True)
            break

        print(f"[Customer] 🆕 接續輪 {followup_round+1}: 對方插話 {len(new_them2)} 條 → 補回應", flush=True)
        # 把新訊息加進 history + .txt
        for m in new_msgs2:
            history.append(m)
            _append_to_history_file(name, m["sender"], m["text"])

        # 再生成 reply 接續回應
        reply2 = generate_reply(augmented_prompt, history, new_them2)
        if not reply2 or len(reply2) <= 1:
            break
        reply2 = filter_reply(reply2)
        if not reply2:
            break

        print(f"[Customer] 接續輪 {followup_round+1} 回覆: {reply2[:80]}", flush=True)
        _send_with_realistic_delay(reply2, regions, name=name, history=history)

    time.sleep(0.5)
    return regions


# ============================================================
# 主流程
# ============================================================
def main(stop_time, monitor=None):
    from 反詐_locate import (
        locate_line_regions, switch_page, find_line_window,
        screenshot_line_window,
    )
    import win32con

    if os.path.exists(STOP_FILE):
        try:
            os.remove(STOP_FILE)
        except OSError:
            pass

    print("=" * 50, flush=True)
    print(f"反詐演示自動回覆", flush=True)
    print(f"監控到：{stop_time}", flush=True)
    print(f"模式：反詐演示（demo 沙盤）", flush=True)
    print("=" * 50, flush=True)

    # 載入時段延遲設定
    _load_time_settings(TIME_SETTINGS_PATH)

    BOX_PERSONA = {
        "demo": "C:/Users/blue_/claude-telegram-bot/scripts/demo/反詐_persona.txt",
    }

    box_prompt_cache = {}
    for _box, _persona_path in BOX_PERSONA.items():
        try:
            box_prompt_cache[_box] = load_persona_prompt(_persona_path)
            print(f"[Init] {_box} → 演示人設載入成功 ({len(box_prompt_cache[_box])} chars)", flush=True)
        except Exception as e:
            print(f"[Init] {_box} 載入人設失敗: {e}", flush=True)

    system_prompt = next(iter(box_prompt_cache.values())) if box_prompt_cache else ""

    line = find_line_window()
    if not line:
        print("[ERROR] 找不到 LINE 視窗", flush=True)
        return False

    SWP_FLAGS = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
    try:
        win32gui.SetWindowPos(line[0], win32con.HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        win32gui.SetWindowPos(line[0], win32con.HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        win32gui.SetWindowPos(line[0], win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              SWP_FLAGS | win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(line[0])
    except Exception:
        pass
    time.sleep(0.5)

    all_histories = {}
    POLL_INTERVAL = 10

    print(f"\n[Monitor] 開始監控未讀訊息...", flush=True)
    print(f"[Monitor] 停止方式：touch {STOP_FILE}", flush=True)

    from 反詐_locate import set_active_box, find_line_window
    import win32gui as _w32g
    BOXES = list(BOX_PERSONA.keys())

    while is_before_stop_time(stop_time):
        if should_stop():
            break

        for box in BOXES:
            if should_stop():
                break
            try:
                set_active_box(box)
                box_prompt = box_prompt_cache.get(box)
                if not box_prompt:
                    print(f"[Multi] box={box} 沒設定人設，跳過", flush=True)
                    continue

                line = find_line_window()
                if not line:
                    print(f"[Multi] 找不到 box={box} 的 LINE 視窗，跳過", flush=True)
                    continue
                hwnd = line[0]
                try:
                    _w32g.SetForegroundWindow(hwnd)
                except Exception as e:
                    print(f"[Multi] SetForegroundWindow {box} 失敗: {e}", flush=True)
                time.sleep(1.0)
                print(f"\n[Multi] === 開始處理 box={box} → 反詐演示 (hwnd={hwnd}) ===", flush=True)

                processed_count = 0
                MAX_PER_BOX = 15
                while processed_count < MAX_PER_BOX:
                    if should_stop():
                        break

                    regions, unread_list = find_unread_conversations(monitor)
                    if not unread_list:
                        if processed_count == 0:
                            print(f"[Multi] box={box} 無未讀，跳下一個", flush=True)
                        else:
                            print(f"[Multi] box={box} 本輪處理完畢（共 {processed_count} 個）", flush=True)
                        break

                    if processed_count == 0:
                        print(f"[Multi] box={box} 偵測到 {len(unread_list)} 個未讀對話", flush=True)

                    conv = unread_list[0]
                    regions = handle_one_customer(
                        conv, regions, box_prompt, all_histories, monitor
                    )

                    if should_stop():
                        break

                    pyautogui.press("escape")
                    time.sleep(0.5)
                    processed_count += 1
                else:
                    print(f"[Multi] box={box} 達單輪上限 {MAX_PER_BOX}，跳下一個", flush=True)

            except Exception as e:
                print(f"[ERR] box={box}: {e}", flush=True)
                import traceback
                traceback.print_exc()
                time.sleep(2)

        # ============================================================
        # 🆕 排程器：每輪 box 處理完後，檢查是否該主動發訊息給某些觀眾
        # 觸發條件：時段問候 / 沉默太久關心 / Day 5+ 副業伏筆
        # ============================================================
        if not should_stop():
            try:
                # 為每個 box 執行排程器（目前只有 demo box）
                for box in BOXES:
                    set_active_box(box)
                    line2 = find_line_window()
                    if not line2:
                        continue
                    try:
                        _w32g.SetForegroundWindow(line2[0])
                    except Exception:
                        pass
                    time.sleep(0.5)

                    # 掃描所有觀眾
                    all_customers = _scan_all_customers()
                    if not all_customers:
                        continue

                    proactive_sent = 0
                    MAX_PROACTIVE_PER_ROUND = 3   # 每輪最多主動發 3 個觀眾，避免一次太多
                    for cust_name in all_customers:
                        if proactive_sent >= MAX_PROACTIVE_PER_ROUND:
                            break
                        if should_stop():
                            break
                        should_send, msg, reason = _check_proactive_trigger(cust_name)
                        if not should_send:
                            continue
                        print(f"[Proactive] {cust_name} 觸發：{reason}", flush=True)
                        ok = _send_proactive_message(cust_name, msg, monitor)
                        if ok:
                            proactive_sent += 1
                            time.sleep(2)   # 主動發之間隔開一下

            except Exception as e:
                print(f"[ERR] 排程器失敗: {e}", flush=True)
                import traceback
                traceback.print_exc()

        for _ in range(POLL_INTERVAL):
            if should_stop():
                break
            time.sleep(1)

    reason = "收到停止信號" if _should_stop else "到達結束時間"
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] 監控結束（{reason}）", flush=True)
    print(f"共處理 {len(all_histories)} 個觀眾", flush=True)
    return True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        print("用法：python 反詐_multi.py <監控時間HH:MM>")
        sys.exit(0)

    stop = sys.argv[1]
    main(stop)
