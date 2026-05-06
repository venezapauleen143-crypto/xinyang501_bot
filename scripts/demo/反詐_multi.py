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

    # 🔴 log rotation: 把 stdout/stderr 同步寫到 logs/反詐_multi.log（100MB × 5 份）
    import logging
    from logging.handlers import RotatingFileHandler
    LOGS_DIR = Path("C:/Users/blue_/claude-telegram-bot/scripts/demo/logs")
    LOGS_DIR.mkdir(parents=True, exist_ok=True)
    _log_handler = RotatingFileHandler(
        str(LOGS_DIR / "反詐_multi.log"),
        maxBytes=100 * 1024 * 1024,  # 100MB
        backupCount=5,
        encoding="utf-8",
    )
    _log_handler.setLevel(logging.INFO)
    _log_handler.setFormatter(logging.Formatter("%(asctime)s %(message)s", datefmt="%Y-%m-%d %H:%M:%S"))

    class _TeeStream:
        """同時把寫入分流到原 stdout 和 log file"""
        def __init__(self, original, handler):
            self.original = original
            self.handler = handler
            self.buffer = ""
        def write(self, s):
            self.original.write(s)
            self.buffer += s
            while "\n" in self.buffer:
                line, self.buffer = self.buffer.split("\n", 1)
                if line.strip():
                    record = logging.LogRecord(
                        name="multi", level=logging.INFO, pathname="", lineno=0,
                        msg=line, args=(), exc_info=None,
                    )
                    self.handler.emit(record)
        def flush(self):
            self.original.flush()
        def __getattr__(self, name):
            return getattr(self.original, name)

    sys.stdout = _TeeStream(sys.stdout, _log_handler)
    sys.stderr = _TeeStream(sys.stderr, _log_handler)

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

# ============================================================
# 即時資訊工具（從 claude_tools.py 借）
# ============================================================
sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")

# 全域 context 快取（30 分鐘 TTL，避免每次對話都查 API）
_WORLD_CONTEXT_CACHE = {}
_WORLD_CONTEXT_TTL = 1800  # 30 分鐘


def _safe_call_tool(tool_name, *args, **kwargs):
    """安全呼叫 claude_tools 函式，失敗回空字串"""
    try:
        import claude_tools
        fn = getattr(claude_tools, tool_name, None)
        if fn is None:
            return ""
        result = fn(*args, **kwargs)
        return str(result) if result else ""
    except Exception as e:
        print(f"[tool] {tool_name} 呼叫失敗: {e}", flush=True)
        return ""


# 台灣常見地名 → 對應查詢城市
_TW_AREA_MAP = {
    "台北": "Taipei", "新北": "New Taipei", "桃園": "Taoyuan",
    "台中": "Taichung", "台南": "Tainan", "高雄": "Kaohsiung",
    "新竹": "Hsinchu", "苗栗": "Miaoli", "彰化": "Changhua",
    "南投": "Nantou", "雲林": "Yunlin", "嘉義": "Chiayi",
    "屏東": "Pingtung", "宜蘭": "Yilan", "花蓮": "Hualien",
    "台東": "Taitung", "基隆": "Keelung",
}


def _detect_opponent_location(name, history):
    """從對方 LINE 名稱 + history 推測所在地（回傳城市名供查天氣）"""
    text_to_scan = name + " " + " ".join(
        m.get("text", "") for m in history if m.get("sender") == "them"
    )
    # 找台灣縣市
    for tw_name, en_name in _TW_AREA_MAP.items():
        if tw_name in text_to_scan or tw_name.replace("台", "臺") in text_to_scan:
            return tw_name
    # 找其他地區
    other_areas = ["香港", "上海", "北京", "廣州", "深圳", "東京", "首爾", "新加坡"]
    for area in other_areas:
        if area in text_to_scan:
            return area
    return None


def _get_world_context(opponent_location=None):
    """組今天的世界資訊：日期、Angela 香港天氣、對方所在地天氣、香港當日新聞。

    用 30 分鐘快取避免每次對話都重複查 API。
    """
    cache_key = (opponent_location or "_none_", datetime.now().strftime("%Y%m%d-%H"))
    if cache_key in _WORLD_CONTEXT_CACHE:
        cached, ts = _WORLD_CONTEXT_CACHE[cache_key]
        if (time.time() - ts) < _WORLD_CONTEXT_TTL:
            return cached

    parts = []
    now = datetime.now()
    weekday_zh = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"][now.weekday()]
    parts.append(
        f"📅 今天：{now.strftime('%Y-%m-%d')} {weekday_zh}，**現在時間 {now.strftime('%H:%M')}**\n"
        f"⚠️ 對方訊息中提到的時間（如「8點半下班」「12點開會」）是對方的『行程時間』，"
        f"不是「現在時間」。要分清楚。"
    )

    # Angela 所在地（香港）天氣
    hk_weather = _safe_call_tool("fetch_weather", "Hong Kong")
    if hk_weather:
        parts.append(f"🌤 Angela 所在地（香港）天氣：\n{hk_weather[:300]}")

    # 對方所在地天氣（如果偵測到）
    if opponent_location:
        opp_weather = _safe_call_tool("fetch_weather", opponent_location)
        if opp_weather:
            parts.append(f"🌤 對方所在地（{opponent_location}）天氣：\n{opp_weather[:300]}")

    # 全球當日新聞 top 3
    global_news = _safe_call_tool("fetch_global_news", count=3)
    if global_news:
        parts.append(f"📰 今日要聞：\n{global_news[:600]}")

    context = "\n\n".join(parts)
    _WORLD_CONTEXT_CACHE[cache_key] = (context, time.time())
    return context


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


def _today_subdir():
    """今日的子目錄路徑（histories/YYYY-MM-DD/）"""
    today = datetime.now().strftime("%Y-%m-%d")
    p = HISTORIES_DIR / today
    p.mkdir(parents=True, exist_ok=True)
    return p


def _history_file_path(name):
    """回傳該觀眾的 .txt 完整路徑（新檔放今日子目錄）"""
    return _today_subdir() / f"{_sanitize_filename(name)}.txt"


def _resolve_history_file(name):
    """找 name 對應的 .txt 檔。

    搜尋順序：
      1. 今日子目錄精確
      2. 平級舊檔精確（向後相容）
      3. 全目錄遞迴 fuzzy match（忽略空格繁簡）
      4. 都沒有 → 回今日子目錄路徑（會建新檔）
    """
    today_path = _history_file_path(name)
    if today_path.exists():
        return today_path

    # 平級舊檔（向後相容 5/6 之前的檔案）
    legacy = HISTORIES_DIR / f"{_sanitize_filename(name)}.txt"
    if legacy.exists():
        return legacy

    # 正規化 fuzzy match
    def _normalize(s):
        s = re.sub(r'[\s　]+', '', s)
        return s.lower()

    target = _normalize(_sanitize_filename(name))
    if not target:
        return today_path

    # 遞迴掃所有子目錄的 .txt
    for f in HISTORIES_DIR.rglob("*.txt"):
        if _normalize(f.stem) == target:
            print(f"[history] 模糊匹配：'{name}' → 沿用既有檔 '{f.relative_to(HISTORIES_DIR)}'", flush=True)
            return f

    # 找不到 → 用今日子目錄（建新檔）
    return today_path


def _load_customer_history(name):
    """從 .txt 讀對應觀眾的歷史，回傳 history list（{text, sender, y} 格式）。

    使用 LINE 匯出格式：
        2026.05.05 星期二                    ← 日期分隔行（跳過）
        HH:MM <發送者名稱> <訊息>             ← 標準訊息行
        <延續訊息>                           ← 沒時間戳的行 = 接續上一條訊息

    判斷誰是 me/them：發送者 == 客戶名（檔名）→ them，否則 → me
    """
    path = _resolve_history_file(name)
    if not path.exists():
        return []

    # 🔴 Bug A 修復：用 path.stem 當 actual_name，並準備 normalize 比對
    # 原因：OCR 抓的 name（如「仁輝JAMES」無空格）可能與 .txt 內格式（「仁輝 JAMES」有空格）不同
    actual_name = path.stem  # 從檔名拿 fuzzy match 後的版本
    name_variants = list({name, actual_name, name.replace(" ", ""), actual_name.replace(" ", "")})

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
                    # 發送者 startswith 客戶名 → them（嘗試所有 name 變體）
                    matched_variant = None
                    for variant in name_variants:
                        if variant and rest.startswith(variant + " "):
                            matched_variant = variant
                            break
                    if matched_variant:
                        sender = "them"
                        text = rest[len(matched_variant) + 1:].strip()
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


# ============================================================
# Profile Memory（雙層記憶架構，保證 100% 不漏記）
# 架構：raw .txt（episodic）+ profile.json（semantic 結構化事實）
# 業界主流 2026：MemMachine / Memori 都用這架構，可達 80%+ token 節省
# ============================================================

def _profile_file_path(name):
    """回傳 profile.json 完整路徑（跟著 .txt 走 — 同目錄）"""
    # 跟對應的 .txt 同目錄存
    txt_path = _resolve_history_file(name)
    return txt_path.parent / f"{_sanitize_filename(name)}.profile.json"


def _resolve_profile_file(name):
    """找 name 對應的 .profile.json（同 history fuzzy match）"""
    expected = _profile_file_path(name)
    if expected.exists():
        return expected

    # 平級舊檔
    legacy = HISTORIES_DIR / f"{_sanitize_filename(name)}.profile.json"
    if legacy.exists():
        return legacy

    def _normalize(s):
        s = re.sub(r'[\s　]+', '', s)
        return s.lower()

    target = _normalize(_sanitize_filename(name))
    if not target:
        return expected

    suffix = ".profile.json"
    for f in HISTORIES_DIR.rglob(f"*{suffix}"):
        bare = f.name[:-len(suffix)]
        if _normalize(bare) == target:
            return f
    return expected


def _load_profile(name):
    """載入 profile.json，沒有回 None"""
    path = _resolve_profile_file(name)
    if not path.exists():
        return None
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[profile] 讀取 {path.name} 失敗：{e}", flush=True)
        return None


def _save_profile(name, profile):
    """存 profile.json"""
    path = _resolve_profile_file(name)
    try:
        with io.open(path, "w", encoding="utf-8") as f:
            json.dump(profile, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"[profile] 寫入 {path.name} 失敗：{e}", flush=True)
        return False


def _strip_json_codeblock(text):
    """剝掉 LLM 偶爾包的 markdown code block"""
    text = text.strip()
    if text.startswith("```"):
        text = re.sub(r'^```(?:json)?\s*\n', '', text)
        text = re.sub(r'\n?```\s*$', '', text)
    return text.strip()


# ============================================================
# Profile schema 固定化（避免 LLM 抽取時欄位漂移，方便之後 aggregate）
# ============================================================
PROFILE_SCHEMA_DEFAULTS = {
    "name": "",
    "core_facts": {
        "occupation": None,
        "location": None,
        "schedule": None,
        "age": None,
        "gender": None,
        "marital_status": None,
        "personality_traits": [],
    },
    "shared_disclosures": [],     # [{speaker, fact, timestamp?}]
    "interests": [],
    "family_relationships": [],
    "milestones": [],
    "current_stage": "",
    "topic_hooks_remaining": [],
    "trust_score": 0,             # 0-10，由話題深度推測
    "ai_suspicion_flags": [],     # 對方曾質疑 AI 的時刻 [{timestamp, quote}]
    "first_seen": "",             # ISO timestamp 首次見面
    "last_updated": "",           # ISO timestamp 最近更新
    "total_turns": 0,             # 累積對話輪數
}


def _validate_and_normalize_profile(profile, name=""):
    """補齊預設欄位 + type 校驗，避免欄位漂移。

    如果 LLM 多輸出未知欄位，保留（不刪）但放在末尾。
    """
    if not isinstance(profile, dict):
        return None

    result = {}
    # 1. 依 schema 順序補齊欄位
    for key, default in PROFILE_SCHEMA_DEFAULTS.items():
        val = profile.get(key, default)

        if isinstance(default, dict):
            # 內層 dict（如 core_facts）也要補齊
            normalized = dict(default)
            if isinstance(val, dict):
                for sub_key, sub_default in default.items():
                    sub_val = val.get(sub_key, sub_default)
                    # type check
                    if isinstance(sub_default, list) and not isinstance(sub_val, list):
                        sub_val = [sub_val] if sub_val else []
                    normalized[sub_key] = sub_val
                # 保留 LLM 多輸出的 sub key
                for sub_key, sub_val in val.items():
                    if sub_key not in default:
                        normalized[sub_key] = sub_val
            result[key] = normalized
        elif isinstance(default, list) and not isinstance(val, list):
            result[key] = [val] if val else []
        elif isinstance(default, int) and not isinstance(val, (int, float)):
            try:
                result[key] = int(val) if val else 0
            except (TypeError, ValueError):
                result[key] = 0
        else:
            result[key] = val

    # 2. 自動補 metadata
    now_iso = datetime.now().isoformat(timespec="seconds")
    if not result["first_seen"]:
        result["first_seen"] = now_iso
    result["last_updated"] = now_iso
    if name and not result["name"]:
        result["name"] = name

    # 3. 保留 LLM 多輸出的未知欄位（不刪）
    for key, val in profile.items():
        if key not in PROFILE_SCHEMA_DEFAULTS and key not in result:
            result[key] = val

    return result


def _extract_profile_from_history(name, history):
    """用 Haiku 4.5 從整個 history 抽結構化 profile（首次見觀眾時用）"""
    if not history:
        return None

    history_text_lines = []
    for msg in history:
        label = "[對方]" if msg["sender"] == "them" else "[Angela]"
        history_text_lines.append(f"{label} {msg['text']}")
    history_text = "\n".join(history_text_lines)

    prompt = (
        f"你是對話分析師。以下是 Angela 跟對方（{name}）的完整聊天記錄。"
        f"請從中**精確抽出**對方揭露的所有事實，輸出嚴格的 JSON。\n\n"
        f"<對話記錄>\n{history_text}\n</對話記錄>\n\n"
        f"<抽取規則>\n"
        f"1. **只抽對方真的說過的事**，Angela 說的不算對方的事實\n"
        f"2. **不要推測**：對方沒說職業就不要寫 occupation\n"
        f"3. **每筆 disclosure** 格式：speaker=\"them\"、fact=「對方說過的事實」\n"
        f"4. 興趣/家人/工作/作息/個性等任何揭露都列進去\n"
        f"5. shared_disclosures 是時序列表，依對話順序\n"
        f"</抽取規則>\n\n"
        f"<JSON 格式>\n"
        f"{{\n"
        f'  "name": "{name}",\n'
        f'  "core_facts": {{\n'
        f'    "occupation": "牙醫 或 null",\n'
        f'    "location": "城市 或 未知",\n'
        f'    "schedule": "作息描述 或 null",\n'
        f'    "personality_traits": ["責任感強", "..."]\n'
        f'  }},\n'
        f'  "shared_disclosures": [\n'
        f'    {{"speaker": "them", "fact": "牙醫"}},\n'
        f'    {{"speaker": "them", "fact": "晚上 8:30 下班"}}\n'
        f'  ],\n'
        f'  "interests": ["..."],\n'
        f'  "family_relationships": ["..."],\n'
        f'  "milestones": ["Day 1: 互換職業"],\n'
        f'  "current_stage": "Day N，描述",\n'
        f'  "topic_hooks_remaining": ["副業", "..."]\n'
        f"}}\n"
        f"</JSON 格式>\n\n"
        f"請直接輸出 JSON，不要任何解釋、不要 markdown code block。"
    )

    try:
        client = anthropic.Anthropic()
        resp = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}]
        )
        text = _strip_json_codeblock(resp.content[0].text)
        profile = json.loads(text)
        # 🔴 schema 校驗：補齊缺漏欄位、補 metadata
        profile = _validate_and_normalize_profile(profile, name=name)
        print(f"[profile] {name} 首次抽取完成（{len(profile.get('shared_disclosures') or [])} 筆 disclosures）", flush=True)
        return profile
    except Exception as e:
        print(f"[profile] 首次抽取 {name} 失敗：{e}", flush=True)
        return None


def _update_profile_incrementally(name, old_profile, new_messages):
    """每輪 reply 後增量更新 profile（只送舊 profile + 最新對話 → Haiku）

    含 critic 驗證：舊 profile 的 disclosures / core_facts 不能消失
    """
    if not old_profile:
        return None
    if not new_messages:
        return old_profile

    new_text_lines = []
    for msg in new_messages:
        label = "[對方]" if msg["sender"] == "them" else "[Angela]"
        new_text_lines.append(f"{label} {msg['text']}")
    new_text = "\n".join(new_text_lines)

    old_profile_json = json.dumps(old_profile, ensure_ascii=False, indent=2)

    prompt = (
        f"你是對話分析師。Angela 跟對方（{name}）有最新一輪對話。請**更新** profile（只加新事實，舊的全保留）。\n\n"
        f"<舊 profile>\n{old_profile_json}\n</舊 profile>\n\n"
        f"<最新對話>\n{new_text}\n</最新對話>\n\n"
        f"<更新規則>\n"
        f"1. **舊 profile 內所有事實必須保留**（除非對方明確改口/糾正）\n"
        f"2. 只**新增**對方剛揭露的事實到對應欄位\n"
        f"3. shared_disclosures 加新筆，舊筆不動\n"
        f"4. milestones 出現新里程碑就加\n"
        f"5. **不要刪舊資料**\n"
        f"</更新規則>\n\n"
        f"請直接輸出更新後的完整 JSON（保留所有舊資料 + 新增新揭露），不要解釋、不要 markdown。"
    )

    try:
        client = anthropic.Anthropic()
        resp = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}]
        )
        text = _strip_json_codeblock(resp.content[0].text)
        new_profile = json.loads(text)

        # 🔴 critic 驗證 1：disclosures 不能消失
        old_disclosures = old_profile.get("shared_disclosures") or []
        new_disclosures = new_profile.get("shared_disclosures") or []
        old_facts = {d.get("fact") for d in old_disclosures if d.get("fact")}
        new_facts = {d.get("fact") for d in new_disclosures if d.get("fact")}
        missing = old_facts - new_facts
        if missing:
            print(f"[profile] ⚠️ critic 偵測到漏掉舊 disclosures：{missing}，merge 回去", flush=True)
            existing = new_facts
            for d in old_disclosures:
                if d.get("fact") and d.get("fact") not in existing:
                    new_disclosures.append(d)
            new_profile["shared_disclosures"] = new_disclosures

        # 🔴 critic 驗證 2：core_facts 已知值不能變空
        old_core = old_profile.get("core_facts") or {}
        new_core = new_profile.get("core_facts") or {}
        for key in ("occupation", "location", "schedule"):
            old_val = old_core.get(key)
            new_val = new_core.get(key)
            if old_val and old_val not in (None, "", "未知", "null") and \
               (not new_val or new_val in (None, "", "未知", "null")):
                print(f"[profile] ⚠️ critic 偵測到 core_facts.{key} 從 '{old_val}' 變空，還原", flush=True)
                new_core[key] = old_val
        new_profile["core_facts"] = new_core

        # 🔴 schema 校驗 + 累計輪數
        new_profile = _validate_and_normalize_profile(new_profile, name=name)
        # total_turns 累加（此次更新算 1 輪）
        new_profile["total_turns"] = (old_profile.get("total_turns") or 0) + 1
        # 保留 first_seen
        if old_profile.get("first_seen"):
            new_profile["first_seen"] = old_profile["first_seen"]

        return new_profile
    except Exception as e:
        print(f"[profile] 增量更新 {name} 失敗：{e}，沿用舊 profile", flush=True)
        return old_profile


def _format_profile_for_prompt(profile):
    """把 profile 渲染成餵給 AI 的純文字區塊（system_prompt 用）"""
    if not profile:
        return ""
    parts = ["📄 **對方檔案**（你已經知道的事實，必須記得，不要假裝不知道）"]

    core = profile.get("core_facts") or {}
    if core.get("occupation"):
        parts.append(f"- 職業：{core['occupation']}")
    if core.get("location") and core["location"] not in ("未知", "", None):
        parts.append(f"- 所在地：{core['location']}")
    if core.get("schedule"):
        parts.append(f"- 作息：{core['schedule']}")
    traits = core.get("personality_traits") or []
    if traits:
        parts.append(f"- 個性：{'、'.join(traits)}")

    interests = profile.get("interests") or []
    if interests:
        parts.append(f"- 興趣愛好：{'、'.join(interests)}")

    family = profile.get("family_relationships") or []
    if family:
        parts.append(f"- 家庭/關係：{'、'.join(family)}")

    disclosures = profile.get("shared_disclosures") or []
    them_d = [d for d in disclosures if d.get("speaker") == "them"]
    if them_d:
        parts.append("\n**對方自爆過的事實（時序）：**")
        for d in them_d[-15:]:
            fact = d.get("fact", "")
            if fact:
                parts.append(f"  - {fact}")

    milestones = profile.get("milestones") or []
    if milestones:
        parts.append("\n**關係里程碑：**")
        for m in milestones[-5:]:
            parts.append(f"  - {m}")

    stage = profile.get("current_stage")
    if stage:
        parts.append(f"\n**當前階段：** {stage}")

    return "\n".join(parts)


def _ensure_profile(name, history):
    """保證該觀眾有 profile：沒有就抽，並存檔。回傳 profile dict 或 None。"""
    profile = _load_profile(name)
    if profile is None and history:
        print(f"[profile] {name} 首次見，從 {len(history)} 條 history 抽 profile...", flush=True)
        profile = _extract_profile_from_history(name, history)
        if profile:
            _save_profile(name, profile)
    return profile


# ============================================================
# Profile 更新 queue + worker（高並發 + rate limit 友善）
# ============================================================
import queue as _queue_mod
import threading as _threading

# 全新觀眾累積到多少條 history 才觸發首次抽 profile
MIN_TURNS_FOR_FIRST_PROFILE = 5

_PROFILE_TASK_QUEUE = _queue_mod.Queue(maxsize=200)
_PROFILE_WORKER_COUNT = 2
_PROFILE_WORKER_STARTED = False
_PROFILE_WORKER_LOCK = _threading.Lock()


def _profile_worker_loop():
    """worker thread：從 queue 取任務 → 抽/更新 profile → critic 驗證 → 存檔"""
    while True:
        try:
            task = _PROFILE_TASK_QUEUE.get(timeout=60)
        except _queue_mod.Empty:
            continue
        if task is None:  # poison pill
            _PROFILE_TASK_QUEUE.task_done()
            break

        name = task["name"]
        history_snapshot = task["history"]
        round_msgs = task["round_msgs"]
        all_profiles_ref = task["all_profiles"]

        try:
            current_profile = all_profiles_ref.get(name)

            if current_profile:
                # 已有 profile → 增量更新
                updated = _update_profile_incrementally(name, current_profile, round_msgs)
                action = "增量更新"
            elif len(history_snapshot) >= MIN_TURNS_FOR_FIRST_PROFILE:
                # 全新觀眾累積夠了 → 首次抽
                print(f"[profile-worker] {name} 累積 {len(history_snapshot)} 條，首次抽 profile...", flush=True)
                updated = _extract_profile_from_history(name, history_snapshot)
                action = "首次抽取"
            else:
                # 資料太少，跳過
                _PROFILE_TASK_QUEUE.task_done()
                continue

            if updated:
                all_profiles_ref[name] = updated
                _save_profile(name, updated)
                disc_n = len((updated.get("shared_disclosures") or []))
                print(f"[profile-worker] {name} {action}完成（{disc_n} 筆 disclosures）", flush=True)
                _emit_event(
                    "profile_first_extract" if action == "首次抽取" else "profile_updated",
                    customer=name,
                    data={"disclosures": disc_n, "action": action},
                )
        except Exception as e:
            err = str(e)
            # rate limit retry：簡單退避（worker 級，不卡其他 task）
            if "rate" in err.lower() or "429" in err:
                print(f"[profile-worker] {name} 撞 rate limit，退避 30 秒後重排：{err[:100]}", flush=True)
                _emit_event("rate_limit", customer=name, data={"err": err[:200]})
                time.sleep(30)
                try:
                    _PROFILE_TASK_QUEUE.put_nowait(task)
                except _queue_mod.Full:
                    print(f"[profile-worker] queue 滿，丟棄 {name} 的 retry", flush=True)
            else:
                print(f"[profile-worker] {name} 處理失敗：{err[:200]}", flush=True)
        finally:
            _PROFILE_TASK_QUEUE.task_done()


def _ensure_profile_workers_started():
    """lazy 啟動 worker（避免 module import 時就跑）"""
    global _PROFILE_WORKER_STARTED
    with _PROFILE_WORKER_LOCK:
        if _PROFILE_WORKER_STARTED:
            return
        for i in range(_PROFILE_WORKER_COUNT):
            t = _threading.Thread(target=_profile_worker_loop, daemon=True, name=f"profile-worker-{i}")
            t.start()
        _PROFILE_WORKER_STARTED = True
        print(f"[profile-worker] {_PROFILE_WORKER_COUNT} 個 worker 啟動", flush=True)


# ============================================================
# 識破偵測：對方說「你是 AI / 機器人 / 不是真人」→ 標記 + Telegram alert
# ============================================================
_AI_SUSPICION_KEYWORDS = (
    "你是ai", "你是 ai", "你是AI", "你是 AI",
    "機器人", "机器人",
    "不是真人", "不是真的人",
    "你是程式", "你是机器", "你是機器",
    "chatgpt", "ChatGPT", "claude",
    "你是AI吧", "你是机器吧", "你是機器吧",
    "假的吧", "你假的", "ai回的",
    "ai 写的", "AI 寫的", "AI寫的", "ai写的",
    "回應太快", "回应太快",
)


def _detect_ai_suspicion(messages):
    """偵測對方訊息中是否有「懷疑 AI」的關鍵字。

    回傳：(suspected_bool, matched_quote 或 None)
    """
    for msg_text in messages:
        low = msg_text.lower().replace(" ", "")
        for kw in _AI_SUSPICION_KEYWORDS:
            kw_norm = kw.lower().replace(" ", "")
            if kw_norm in low:
                return True, msg_text
    return False, None


def _send_telegram_alert(text):
    """透過 Telegram bot 發送 alert 給用戶（venezapauleen143 user id 8362721681）"""
    try:
        import urllib.request
        import urllib.parse
        # 從 .env 拿 token（小牛馬 bot 共用）
        token = os.environ.get("TELEGRAM_BOT_TOKEN")
        if not token:
            return False
        chat_id = "8362721681"  # 于晏哥
        url = f"https://api.telegram.org/bot{token}/sendMessage"
        data = urllib.parse.urlencode({
            "chat_id": chat_id,
            "text": text,
            "parse_mode": "HTML",
        }).encode("utf-8")
        req = urllib.request.Request(url, data=data, method="POST")
        with urllib.request.urlopen(req, timeout=5) as resp:
            return resp.status == 200
    except Exception as e:
        print(f"[alert] Telegram 失敗: {e}", flush=True)
        return False


def _emit_event(event_type, customer="", data=None):
    """關鍵事件寫進 events.jsonl，供之後 aggregate 分析

    欄位：timestamp / event_type / customer / data
    Event types:
      - customer_seen: 首次見到該客戶
      - new_messages: 偵測到新對方訊息
      - reply_sent: AI 已回覆
      - profile_first_extract: 首次抽 profile 完成
      - profile_updated: 增量更新 profile 完成
      - ocr_failed: OCR 異常
      - ai_suspicion: 對方疑似識破 AI
      - rate_limit: 撞 rate limit
    """
    try:
        events_dir = Path("C:/Users/blue_/claude-telegram-bot/scripts/demo/logs")
        events_dir.mkdir(parents=True, exist_ok=True)
        # 按日切檔，每日一個 events_YYYY-MM-DD.jsonl
        today = datetime.now().strftime("%Y-%m-%d")
        events_file = events_dir / f"events_{today}.jsonl"
        record = {
            "ts": datetime.now().isoformat(timespec="seconds"),
            "type": event_type,
            "customer": customer,
            "data": data or {},
        }
        with io.open(events_file, "a", encoding="utf-8") as f:
            f.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception as e:
        print(f"[event] write failed: {e}", flush=True)


def _enqueue_profile_task(name, history, round_msgs, all_profiles):
    """把 profile 任務丟進 queue（非阻塞，queue 滿就丟棄並 log）"""
    _ensure_profile_workers_started()
    # snapshot history（避免 worker 處理時 history 已經變動）
    history_snapshot = list(history)
    task = {
        "name": name,
        "history": history_snapshot,
        "round_msgs": round_msgs,
        "all_profiles": all_profiles,
    }
    try:
        _PROFILE_TASK_QUEUE.put_nowait(task)
    except _queue_mod.Full:
        print(f"[profile-worker] ⚠️ queue 滿（>200），丟棄 {name} 的更新", flush=True)


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
    path = _resolve_history_file(name)
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
    """分段送 reply（段間隨機 1-3 秒），送一段立刻寫一段（防 crash）。

    ⚠️ 思考延遲已**移到 handle_one_customer 開頭**（點擊前），
    避免「秒讀延遲回」破綻。這裡只保留「段間延遲」讓多條訊息看起來像真人一條一條打。

    遇到 {send_xxx} 標記 → 從對應圖片庫隨機抽一張用 send_image 送。
    history: 傳進來的 list，會被 append 每段送出的訊息（讓 main 流程不用再寫一次）
    """
    from 反詐_chat import send_reply, send_image

    # 分段送（每段間 1-3 秒），送一段立刻寫一段
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
            # 🔴 段間延遲 = 基礎延遲 + 「下一段字數比例」打字時間
            # 模擬真人：邊想（基礎）+ 邊打字（字數越多打越久）
            base = get_inter_message_delay()
            next_part = parts[i + 1] if i + 1 < len(parts) else ""
            # 排除圖片標記（{send_xxx}）的字數計算
            if not (next_part.startswith("{") and next_part.endswith("}")):
                # 每字 0.3-0.5 秒（手機打字速度）
                typing_time = len(next_part) * random.uniform(0.3, 0.5)
                # 上限 30 秒（避免長段卡死）
                typing_time = min(typing_time, 30)
            else:
                typing_time = 0
            total_delay = base + typing_time
            print(f"[Reply] 段間延遲 {total_delay:.1f}s（基礎 {base:.1f}s + 打字 {typing_time:.1f}s）", flush=True)
            time.sleep(total_delay)

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


def _drain_profile_queue(timeout_sec=30):
    """關閉前等 profile queue 處理完（最多 timeout 秒）"""
    if not _PROFILE_WORKER_STARTED:
        return
    start = time.time()
    while not _PROFILE_TASK_QUEUE.empty() and (time.time() - start) < timeout_sec:
        remaining = _PROFILE_TASK_QUEUE.qsize()
        if remaining > 0:
            print(f"[Cleanup] 等 profile queue 清空（剩 {remaining} 個任務）...", flush=True)
        time.sleep(2)
    if _PROFILE_TASK_QUEUE.empty():
        print("[Cleanup] profile queue 已清空", flush=True)
    else:
        print(f"[Cleanup] profile queue 還有 {_PROFILE_TASK_QUEUE.qsize()} 個任務未處理（timeout）", flush=True)


def _graceful_shutdown():
    """關閉時：先 drain queue，再清 GPU"""
    _drain_profile_queue()
    _cleanup_gpu()


atexit.register(_graceful_shutdown)


# ============================================================
# Startup recovery: 啟動時掃描既有觀眾資料，印 status
# ============================================================
def _print_startup_recovery():
    """印 startup state（從 histories/ 推算累積 active sessions）"""
    try:
        if not HISTORIES_DIR.exists():
            return
        txt_files = list(HISTORIES_DIR.rglob("*.txt"))
        json_files = list(HISTORIES_DIR.rglob("*.profile.json"))
        print(f"\n[Startup] 累積觀眾資料：", flush=True)
        print(f"  - 對話 .txt: {len(txt_files)} 個", flush=True)
        print(f"  - profile .json: {len(json_files)} 個", flush=True)
        # 統計各 stage 分佈（從 profile 取 current_stage）
        stage_count = {}
        for jf in json_files:
            try:
                with io.open(jf, "r", encoding="utf-8") as f:
                    p = json.load(f)
                stage = (p.get("current_stage") or "未知")[:30]
                stage_count[stage] = stage_count.get(stage, 0) + 1
            except Exception:
                continue
        if stage_count:
            print(f"  - stage 分佈：", flush=True)
            for stage, n in sorted(stage_count.items(), key=lambda x: -x[1])[:10]:
                print(f"      {n}x {stage}", flush=True)
    except Exception as e:
        print(f"[Startup] recovery 摘要失敗：{e}", flush=True)

# ============================================================
# 優雅停止機制
# ============================================================
STOP_FILE = "C:/Users/blue_/Desktop/測試檔案/.stop_line_multi"
_should_stop = False

def _signal_handler(signum, frame):
    global _should_stop
    _should_stop = True
    print(f"\n[STOP] 收到信號 {signum}，準備停止（drain queue 後）...", flush=True)
    _graceful_shutdown()

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
def handle_one_customer(conv, regions, system_prompt, all_histories, all_profiles=None, monitor=None):
    if all_profiles is None:
        all_profiles = {}
    from 反詐_locate import locate_line_regions, ocr_scan_panel, screenshot_line
    from 反詐_chat import (
        is_only_sticker, analyze_sticker,
    )
    from difflib import SequenceMatcher

    cx, cy = conv["center"]

    # 🔴 思考延遲（在點進去之前，避免「秒讀延遲回」破綻）
    # LINE 點進對話會立刻已讀 → 對方看到「秒讀」+「90 秒後才回」會覺得故意冷處理
    # 改成在外面 sleep → 對方看到「過幾分鐘才已讀 + 立刻回覆」 → 自然像真人
    think_delay = get_current_delay()
    print(f"[Customer] 偵測到未讀，等 {think_delay:.1f} 秒再點進去（避免秒讀）...", flush=True)
    time.sleep(think_delay)
    if should_stop():
        return regions

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

    # 🔴 OCR 異常容錯：失敗 retry 1 次 + 仍失敗就 skip 該客戶（避免 crash）
    current_messages = None
    for ocr_attempt in range(2):
        try:
            chat_img = grab_chat_area(regions, monitor)
            current_messages = ocr_extract_messages(chat_img)
            break
        except Exception as e:
            err_str = str(e)[:200]
            if ocr_attempt == 0:
                print(f"[Customer] OCR 失敗（嘗試 1/2），retry：{err_str}", flush=True)
                time.sleep(1.0)
            else:
                print(f"[Customer] ⚠️ OCR 連續失敗，跳過此客戶：{err_str}", flush=True)
                _emit_event("ocr_failed", customer=name, data={"err": err_str})
                return regions
    if current_messages is None:
        return regions

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
        _emit_event("customer_seen", customer=name, data={"prior_history_len": len(loaded)})
        all_histories[name] = loaded
        # 🔴 lazy 載入 profile：cache 命中直接用，沒 profile 不阻塞主流程，丟進背景 queue 抽
        cached_profile = _load_profile(name) if loaded else None
        if cached_profile:
            all_profiles[name] = cached_profile
            print(f"[Customer] {name} profile 已 cache 命中（{len(cached_profile.get('shared_disclosures') or [])} 筆）", flush=True)
        else:
            all_profiles[name] = None
            # 有 history（≥ MIN_TURNS_FOR_FIRST_PROFILE）就背景觸發首次抽
            if loaded and len(loaded) >= MIN_TURNS_FOR_FIRST_PROFILE:
                print(f"[Customer] {name} 無 profile，背景排程首次抽（不阻塞）", flush=True)
                _enqueue_profile_task(name, loaded, [], all_profiles)
    history = all_histories[name]
    profile = all_profiles.get(name)

    if not history:
        new_them = [m["text"] for m in current_messages if m["sender"] == "them"]
        for m in current_messages:
            history.append(m)
        # 第一次處理 → 把 OCR 看到的所有對方訊息寫進 .txt
        for m in current_messages:
            if m["sender"] == "them":
                _append_to_history_file(name, "them", m["text"])
    else:
        # 🔴 fallback 升級：先用 history[-1] 找 cutoff，找不到再往前 history[-5:] 重試
        # 解決：LINE 視窗滾動時 OCR 抓不到 history[-1] → 不會 fallback 全部當新訊息（重複回舊）
        match_idx = -1
        matched_via = None
        for hist_offset in range(1, min(6, len(history) + 1)):
            cand_text = history[-hist_offset]["text"]
            for i in range(len(current_messages) - 1, -1, -1):
                ratio = SequenceMatcher(None, current_messages[i]["text"], cand_text).ratio()
                if ratio > 0.6:
                    match_idx = i
                    matched_via = hist_offset
                    break
            if match_idx >= 0:
                break

        if match_idx >= 0 and match_idx < len(current_messages) - 1:
            if matched_via and matched_via > 1:
                print(f"[Customer] cutoff 用 history[-{matched_via}] 對到（[-1] 在視窗外）", flush=True)
            new_msgs = current_messages[match_idx + 1:]
            new_them = [m["text"] for m in new_msgs if m["sender"] == "them"]
            for m in new_msgs:
                history.append(m)
            # 新訊息（含對方+自己）即時 append 到 .txt
            for m in new_msgs:
                _append_to_history_file(name, m["sender"], m["text"])
        else:
            # history[-5:] 都對不上 → 真的找不到 → 只取最後 1 條當新訊息（保守）
            # 比之前「全部當新訊息」安全，避免重複回舊訊息
            print(f"[Customer] ⚠️ history[-5:] 都對不上 OCR，保守取最後 1 條對方訊息", flush=True)
            them_msgs = [m for m in current_messages if m["sender"] == "them"]
            new_them = [them_msgs[-1]["text"]] if them_msgs else []

    if not new_them:
        print(f"[Customer] 沒有新的對方訊息，跳過", flush=True)
        return regions

    print(f"[Customer] 新訊息: {new_them}", flush=True)

    # 🔴 識破偵測：對方說「你是 AI / 機器人」→ 標記 profile + Telegram alert
    suspected, suspect_quote = _detect_ai_suspicion(new_them)
    if suspected:
        print(f"[Customer] ⚠️⚠️ {name} 疑似識破 AI！對方說：「{suspect_quote}」", flush=True)
        _emit_event("ai_suspicion", customer=name, data={"quote": suspect_quote})
        # 標記 profile
        cur_profile = all_profiles.get(name)
        if cur_profile:
            flags = cur_profile.get("ai_suspicion_flags") or []
            flags.append({
                "timestamp": datetime.now().isoformat(timespec="seconds"),
                "quote": suspect_quote,
            })
            cur_profile["ai_suspicion_flags"] = flags
            _save_profile(name, cur_profile)
        # Telegram alert
        alert_text = f"⚠️ <b>{name}</b> 疑似識破 AI\n\n對方說：「{suspect_quote}」\n\n時間：{datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
        _send_telegram_alert(alert_text)

    # 純貼圖 → Vision
    if is_only_sticker(new_them):
        meaning = analyze_sticker(regions, monitor)
        print(f"[Customer] 純貼圖 → Vision 解讀為「{meaning}」", flush=True)
        new_them = [f"[貼圖含意：{meaning}]"]
    else:
        # 🆕 偵測對方是否傳了「真實照片」（不是貼圖、不是文字）
        # 用 Haiku Vision 看 chat_area 底部，回傳照片描述
        from 反詐_chat import analyze_recent_photo
        photo_desc = analyze_recent_photo(regions, monitor)
        if photo_desc:
            print(f"[Customer] 偵測到對方傳照片 → Vision 描述: {photo_desc}", flush=True)
            # 把照片描述加進 new_them 讓 AI 看到
            new_them = list(new_them) + [f"[對方傳了一張照片：{photo_desc}]"]

    # AI 生成回覆（注入階段感知 Day N + 今天的世界資訊）
    day_n = _calculate_day_n(name)
    stage_hint = _build_stage_hint(day_n)
    opponent_location = _detect_opponent_location(name, history)
    world_context = _get_world_context(opponent_location)
    print(f"[Customer] {name} 對話 Day {day_n}, 對方所在地: {opponent_location or '未知'}", flush=True)
    augmented_prompt = (
        system_prompt + stage_hint +
        "\n\n=== 今天的世界資訊（給你即時參考，避免破綻）===\n" + world_context
    )
    # 🔴 渲染 profile 給 AI 看（保證記得對方所有事實）
    profile_text = _format_profile_for_prompt(profile) if profile else ""
    if profile_text:
        print(f"[Customer] {name} profile 注入（{len((profile or {}).get('shared_disclosures') or [])} 筆事實）", flush=True)

    reply = generate_reply(augmented_prompt, history, new_them, profile_text=profile_text)
    if not reply or len(reply) <= 1:
        time.sleep(0.5)
        return regions

    reply = filter_reply(reply)
    if not reply:
        print(f"[Customer] 過濾後為空，跳過", flush=True)
        return regions

    print(f"[Customer] 回覆: {reply[:80]}", flush=True)
    _emit_event("new_messages", customer=name, data={"new_them_count": len(new_them), "day_n": day_n})
    # 🔴 送一段寫一段（history + .txt 由 _send_with_realistic_delay 內部即時寫入）
    _send_with_realistic_delay(reply, regions, name=name, history=history)
    _emit_event("reply_sent", customer=name, data={"reply_len": len(reply), "segments": reply.count("|||") + 1, "day_n": day_n})

    # 🔴 profile 處理（背景 queue，不阻塞主流程）
    # 邏輯：
    #   - 已有 profile → 增量更新
    #   - 沒 profile 但 history >= MIN_TURNS_FOR_FIRST_PROFILE → 首次抽
    #   - history < 閾值 → 跳過（資料太少抽不出實質事實）
    round_msgs = [{"sender": "them", "text": t} for t in new_them]
    round_msgs.append({"sender": "me", "text": reply})
    _enqueue_profile_task(name, history, round_msgs, all_profiles)

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

        # 再生成 reply 接續回應（也帶 profile）
        # 用 cache 裡最新的 profile（背景 thread 可能已經更新完）
        latest_profile = all_profiles.get(name) or profile
        latest_profile_text = _format_profile_for_prompt(latest_profile) if latest_profile else ""
        reply2 = generate_reply(augmented_prompt, history, new_them2, profile_text=latest_profile_text)
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
    all_profiles = {}  # 🔴 結構化 profile cache（保證 100% 記憶）
    POLL_INTERVAL = 10

    # 🔴 startup recovery 印 status
    _print_startup_recovery()

    print(f"\n[Monitor] 開始監控未讀訊息...", flush=True)
    print(f"[Monitor] 停止方式：touch {STOP_FILE}", flush=True)

    from 反詐_locate import set_active_box, find_line_window
    import win32gui as _w32g
    BOXES = list(BOX_PERSONA.keys())

    # 🔴 GPU 記憶體週期釋放（每 GPU_CLEAN_INTERVAL 個主迴圈跑一次）
    GPU_CLEAN_INTERVAL = 50  # ~ 50 × 10 秒 POLL = 約每 8 分鐘清一次
    STATUS_PRINT_INTERVAL = 6  # ~ 6 × 10 秒 = 約每 1 分鐘印 status
    _loop_counter = 0

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
                        conv, regions, box_prompt, all_histories, all_profiles, monitor
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

        # 🔴 status snapshot 每分鐘印一次
        _loop_counter += 1
        if _loop_counter % STATUS_PRINT_INTERVAL == 0:
            try:
                active_n = len(all_histories)
                profile_n = sum(1 for p in all_profiles.values() if p)
                # day 分佈
                day_dist = {}
                for cn in all_histories:
                    try:
                        dn = _calculate_day_n(cn)
                        day_dist[dn] = day_dist.get(dn, 0) + 1
                    except Exception:
                        continue
                # AI 識破累計
                susp_n = sum(
                    len((p.get("ai_suspicion_flags") or []))
                    for p in all_profiles.values() if p
                )
                queue_n = _PROFILE_TASK_QUEUE.qsize() if _PROFILE_WORKER_STARTED else 0
                day_str = " ".join(f"day{d}:{n}" for d, n in sorted(day_dist.items()))
                print(f"[Status] active={active_n} profile={profile_n} suspicion={susp_n} queue={queue_n} | {day_str}", flush=True)
            except Exception as e:
                print(f"[Status] 失敗：{e}", flush=True)

        # 🔴 GPU 記憶體週期釋放（避免 PaddleOCR 長時間 leak）
        if _loop_counter % GPU_CLEAN_INTERVAL == 0:
            try:
                gc.collect()
                import paddle
                paddle.device.cuda.empty_cache()
                print(f"[GPU] 週期釋放（loop #{_loop_counter}）", flush=True)
            except Exception:
                pass

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
