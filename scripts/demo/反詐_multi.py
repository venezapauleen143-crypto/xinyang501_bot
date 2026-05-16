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

# 🔴 log rotation: 把 stdout/stderr 同步寫到 logs/反詐_multi.log（100MB × 5 份）
# 放在 Path import 之後，避免 NameError
if __name__ == "__main__":
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

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SCRIPT_DIR.parent))

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
client = anthropic.Anthropic()
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"
# 🔴 舊版 DEFAULT_PERSONA（指向已搬走的 .txt）已棄用，由 PERSONA_CONFIG 取代
# 真正的 DEFAULT_PERSONA 在 PERSONA_CONFIG 區塊（值 = "Angela"）
TIME_SETTINGS_PATH = "C:/Users/blue_/claude-telegram-bot/scripts/demo/反詐_時段設定.json"

# 全域時段設定（main() 開頭載入）
_TIME_SETTINGS = None


# ============================================================
# Persona 隔離架構（多人設徹底分開：histories / images / logs 各自獨立）
# ============================================================
SCRIPT_DIR = Path(__file__).resolve().parent

PERSONA_CONFIG = {
    "Angela": {
        "sandbox": "demo",
        "nickname": "Angela",
        "self_location": "Hong Kong",
        "self_location_zh": "香港",
        "root": SCRIPT_DIR / "personas" / "Angela",
        "active": False,  # 🔴 暫停用（用戶切到 Ruby 測試）
    },
    "Ruby": {
        "sandbox": "demo",  # 🔴 暫時用同 demo 沙盤（之後建 demo_ruby 再切回）
        "nickname": "Ruby🧸",
        "self_location": "Taipei",
        "self_location_zh": "台北",
        "root": SCRIPT_DIR / "personas" / "Ruby",
        "active": True,  # 🔴 啟用中
    },
}

# 🔴 DEFAULT_PERSONA 動態決定：取第一個 active persona（避免 Angela 停用後仍用 Angela 當 fallback）
def _resolve_default_persona():
    actives = [k for k, v in PERSONA_CONFIG.items() if v.get("active", False)]
    if actives:
        return actives[0]
    # 沒任何 active → 取 PERSONA_CONFIG 的第一個（保底，避免 KeyError）
    return next(iter(PERSONA_CONFIG))


DEFAULT_PERSONA = _resolve_default_persona()


def get_active_personas():
    """回傳啟用中的 personas（dict[persona_name → config]）"""
    return {k: v for k, v in PERSONA_CONFIG.items() if v.get("active", False)}


def persona_root(persona: str) -> Path:
    p = PERSONA_CONFIG[persona]["root"]
    p.mkdir(parents=True, exist_ok=True)
    return p


def persona_histories_dir(persona: str) -> Path:
    p = persona_root(persona) / "histories"
    p.mkdir(parents=True, exist_ok=True)
    return p


def persona_images_dir(persona: str) -> Path:
    p = persona_root(persona) / "images"
    p.mkdir(parents=True, exist_ok=True)
    return p


def persona_logs_dir(persona: str) -> Path:
    p = persona_root(persona) / "logs"
    p.mkdir(parents=True, exist_ok=True)
    return p


def persona_nickname(persona: str) -> str:
    return PERSONA_CONFIG[persona]["nickname"]


def persona_persona_file(persona: str) -> Path:
    return persona_root(persona) / "persona.txt"


def get_persona_for_sandbox(sandbox: str) -> str:
    """sandbox name → persona name（反向查找）"""
    for name, cfg in PERSONA_CONFIG.items():
        if cfg["sandbox"] == sandbox:
            return name
    return DEFAULT_PERSONA


# 確保所有 active persona 的目錄都存在（C-7: 跳過 inactive，避免 Angela=False 也建空目錄）
for _p, _cfg in PERSONA_CONFIG.items():
    if not _cfg.get("active", True):
        continue
    persona_root(_p)
    persona_histories_dir(_p)
    persona_images_dir(_p)
    persona_logs_dir(_p)

# ============================================================
# 即時資訊工具（從 claude_tools.py 借）
# ============================================================
sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")

# 全域 context 快取（30 分鐘 TTL，避免每次對話都查 API）
_WORLD_CONTEXT_CACHE = {}

# Proactive 主動發訊息 cooldown：嘗試後（無論成功失敗）30 分鐘不再對同客戶觸發
# 解 2026-05-11 bug：失敗時沒冷卻 → 19 分鐘 13 次重複觸發
_LAST_PROACTIVE_ATTEMPT = {}  # {(persona, name): unix_ts}
PROACTIVE_COOLDOWN_SEC = 1800  # 30 分鐘

# Fuzzy match log 去重：同對話只 log 一次「沿用既有檔」訊息（避免主迴圈每輪都印）
_FUZZY_MATCH_LOGGED = set()  # {(persona, name, file_path)}
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


# 其他地區中→英對應（fetch_weather 需要英文 city name）
_OTHER_AREA_MAP = {
    "香港": "Hong Kong",
    "上海": "Shanghai",
    "北京": "Beijing",
    "廣州": "Guangzhou",
    "深圳": "Shenzhen",
    "東京": "Tokyo",
    "首爾": "Seoul",
    "新加坡": "Singapore",
}


def _detect_opponent_location(name, history):
    """從對方 LINE 名稱 + history 推測所在地，**回傳英文城市名**（fetch_weather 用）"""
    text_to_scan = name + " " + " ".join(
        m.get("text", "") for m in history if m.get("sender") == "them"
    )
    # 找台灣縣市（找中文 → 回英文）
    for tw_name, en_name in _TW_AREA_MAP.items():
        if tw_name in text_to_scan or tw_name.replace("台", "臺") in text_to_scan:
            return en_name
    # 找其他地區（找中文 → 回英文）
    for zh_name, en_name in _OTHER_AREA_MAP.items():
        if zh_name in text_to_scan:
            return en_name
    return None


# 🔴 Phase 3: CWA 中央氣象署 — 給台灣 persona 用（補強 wttr.in 不夠細）
# 含紫外線 / 降雨機率 / 舒適度 / 颱風警報（wttr 沒有）
_CWA_CITY_MAP = {
    "台北": "臺北市", "臺北": "臺北市", "新北": "新北市", "桃園": "桃園市",
    "台中": "臺中市", "臺中": "臺中市", "台南": "臺南市", "臺南": "臺南市",
    "高雄": "高雄市", "基隆": "基隆市", "新竹": "新竹市", "嘉義": "嘉義市",
}


def _fetch_cwa_weather(city_zh):
    """CWA 中央氣象署 36 小時預報。台灣 persona 用，香港 persona 不用"""
    token = os.environ.get("CWA_API_KEY")
    if not token:
        return ""
    cwa_city = _CWA_CITY_MAP.get(city_zh)
    if not cwa_city:
        return ""  # 非台灣城市跳過
    try:
        import requests as _req
        r = _req.get(
            "https://opendata.cwa.gov.tw/api/v1/rest/datastore/F-C0032-001",
            params={"Authorization": token, "locationName": cwa_city},
            timeout=10,
        )
        data = r.json()
        if data.get("success") != "true":
            return ""
        locs = data["records"]["location"]
        if not locs:
            return ""
        loc = locs[0]
        elements = {
            e["elementName"]: e["time"][0]["parameter"]["parameterName"]
            for e in loc["weatherElement"]
        }
        wx = elements.get("Wx", "")
        pop = elements.get("PoP", "")
        mint = elements.get("MinT", "")
        maxt = elements.get("MaxT", "")
        ci = elements.get("CI", "")
        return f"🌤 {cwa_city}（CWA 官方）：{wx}，{mint}-{maxt}°C，降雨機率 {pop}%，{ci}"
    except Exception:
        return ""


# 🔴 Phase 2: 台灣本土 RSS 新聞 top5（補強 fetch_global_news 預設 country=us 不夠台灣本土）
def _fetch_taiwan_news(top=5):
    """拉台灣本土 RSS 新聞 top N — 自由 / 公視"""
    try:
        import feedparser as _fp
    except ImportError:
        return ""

    sources = {
        "自由": "https://news.ltn.com.tw/rss/all.xml",
        "公視": "https://news.pts.org.tw/xml/newsfeed.xml",
    }
    items = []
    for name, url in sources.items():
        try:
            feed = _fp.parse(url)
            for e in feed.entries[:3]:
                title = getattr(e, "title", "").strip()
                if title:
                    items.append(f"[{name}] {title[:50]}")
            if len(items) >= top:
                break
        except Exception:
            continue

    if not items:
        return ""
    return "📰 台灣本土新聞：\n" + "\n".join(items[:top])


# 🔴 Phase 1: 台灣節日 / 節氣 — 注入 _get_world_context
# 含法定假日（holidays 套件）+ 民俗節日（母親節 / 父親節 / 七夕 / 雙十一 / 聖誕 / 跨年）
# 為什麼：14 天連跑會跨節日，Ruby 在母親節不提媽媽 / 端午沒粽子梗會破功
def _fetch_taiwan_holiday(check_days_ahead=3):
    """查今天 + 接下來 N 天的台灣節日"""
    try:
        import holidays as _holidays_mod
    except ImportError:
        return ""

    today = datetime.now().date()
    tw_holidays = _holidays_mod.Taiwan(years=[today.year, today.year + 1])

    # 民俗節日（不在法定假日內但 Ruby 會聊）
    from datetime import timedelta as _td
    custom = {}
    for y in (today.year, today.year + 1):
        # 母親節 = 5 月第二個星期日
        may_first = datetime(y, 5, 1).date()
        days_to_sunday = (6 - may_first.weekday()) % 7
        custom[may_first + _td(days=days_to_sunday + 7)] = "母親節"
        # 父親節 = 8/8
        custom[datetime(y, 8, 8).date()] = "父親節"
        # 雙十一
        custom[datetime(y, 11, 11).date()] = "雙十一光棍節"
        # 聖誕
        custom[datetime(y, 12, 24).date()] = "平安夜"
        custom[datetime(y, 12, 31).date()] = "跨年夜"
        # 情人節
        custom[datetime(y, 2, 14).date()] = "情人節"

    # 合併（法定優先）
    all_days = dict(tw_holidays)
    for d, n in custom.items():
        if d not in all_days:
            all_days[d] = n

    # 找今天 + 接下來 N 天
    upcoming = []
    for i in range(check_days_ahead + 1):
        check = today + _td(days=i)
        if check in all_days:
            label = "今天" if i == 0 else ("明天" if i == 1 else f"{i} 天後")
            upcoming.append(f"{label}（{check.strftime('%m/%d')}）{all_days[check]}")

    if not upcoming:
        return ""
    return "🎉 近期節日：" + " / ".join(upcoming)


def _get_world_context(persona=DEFAULT_PERSONA, opponent_location=None):
    """組今天的世界資訊：日期、自己所在地天氣、對方所在地天氣、新聞

    每個 persona 自己的所在地不同（Angela 香港 / Ruby 台北），cache key 含 persona
    invalid persona → 不 silent fallback Angela，而是 print warn + return ""
    """
    cfg = PERSONA_CONFIG.get(persona)
    if not cfg:
        print(f"[WARN] _get_world_context: invalid persona '{persona}'，跳過世界資訊（避免 silent fallback）", flush=True)
        return ""
    self_loc_en = cfg.get("self_location") or "Hong Kong"
    self_loc_zh = cfg.get("self_location_zh") or "香港"
    nick = cfg.get("nickname") or persona

    cache_key = (persona, opponent_location or "_none_", datetime.now().strftime("%Y%m%d-%H"))
    if cache_key in _WORLD_CONTEXT_CACHE:
        cached, ts = _WORLD_CONTEXT_CACHE[cache_key]
        if (time.time() - ts) < _WORLD_CONTEXT_TTL:
            return cached

    parts = []
    now = datetime.now()
    weekday_zh = ["星期一", "星期二", "星期三", "星期四", "星期五", "星期六", "星期日"][now.weekday()]
    parts.append(
        f"📅 今天：{now.strftime('%Y-%m-%d')} {weekday_zh}，**現在時間 {now.strftime('%H:%M')}**\n"
        f"⚠️ 對方訊息中提到的時間是對方的『行程時間』，不是「現在時間」。要分清楚。"
    )

    # 自己所在地天氣（依 persona 不同）
    # 🔴 Phase 3: 台灣 persona 優先用 CWA 中央氣象署（更細：紫外線/降雨機率/舒適度），香港 persona 走 wttr.in
    cwa_self = _fetch_cwa_weather(self_loc_zh)
    if cwa_self:
        parts.append(cwa_self)
    else:
        self_weather = _safe_call_tool("fetch_weather", self_loc_en)
        if self_weather:
            parts.append(f"🌤 {nick} 所在地（{self_loc_zh}）天氣：\n{self_weather[:300]}")

    if opponent_location:
        opp_weather = _safe_call_tool("fetch_weather", opponent_location)
        if opp_weather:
            parts.append(f"🌤 對方所在地（{opponent_location}）天氣：\n{opp_weather[:300]}")

    global_news = _safe_call_tool("fetch_global_news", count=3)
    if global_news:
        parts.append(f"📰 今日要聞：\n{global_news[:600]}")

    # 🔴 Phase 1: 台灣節日 / 節氣（只在今天或 3 天內有節才注入）
    holiday_info = _fetch_taiwan_holiday(check_days_ahead=3)
    if holiday_info:
        parts.append(holiday_info)

    # 🔴 Phase 2: 台灣本土 RSS 新聞 top5
    tw_news = _fetch_taiwan_news(top=5)
    if tw_news:
        parts.append(tw_news)

    context = "\n\n".join(parts)
    _WORLD_CONTEXT_CACHE[cache_key] = (context, time.time())
    return context


def _resolve_image_tag(persona, tag):
    """從 {send_xxx} 標記找對應圖片（從該 persona 的 images/ 目錄）"""
    # tag 形如 "send_自拍"、"send_食物"
    category = tag.replace("send_", "", 1).strip()
    cat_dir = persona_images_dir(persona) / category
    if not cat_dir.exists():
        return None
    images = list(cat_dir.glob("*.jpg")) + list(cat_dir.glob("*.jpeg")) + list(cat_dir.glob("*.png"))
    if not images:
        return None
    return str(random.choice(images))


def _resolve_audio_tag(persona, tag):
    """從 {audio_xxx} 標記找對應音訊（從該 persona 的 audio/ 目錄）"""
    # tag 形如 "audio_早安"、"audio_笑聲"
    category = tag.replace("audio_", "", 1).strip()
    audio_dir = PERSONA_CONFIG[persona]["root"] / "audio" / category
    if not audio_dir.exists():
        return None
    candidates = (
        list(audio_dir.glob("*.m4a"))
        + list(audio_dir.glob("*.mp3"))
        + list(audio_dir.glob("*.wav"))
    )
    if not candidates:
        return None
    return str(random.choice(candidates))

# 🔴 PERSONA_NICKNAME 已移除，請用 persona_nickname(persona) 動態取
# 每個 persona 的 LINE 暱稱在 PERSONA_CONFIG 設定

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
    # LINE 日期分隔線（居中無氣泡的「今天 / 昨天 / 前天」字眼會被 OCR 抓到）
    # 解 5/16 bug：「今天」被當對方訊息 → Ruby 回「你剛才說『今天』然後呢」
    if text in ("今天", "昨天", "前天"):
        return True
    # LINE 完整日期分隔線（如「2026.05.16 星期六」/「2026/5/16」）
    if re.match(r'^\d{4}[./年-]\d{1,2}', text):
        return True
    # 「以下為尚未...的訊息」前綴（OCR 對「閱讀」辨識多種變體：閱請/開讀/開請/閱讀的息）
    if text.startswith("以下為尚未") or "以下為尚未" in text:
        return True
    # 啟發式：純標點/符號（沒有中文、英文、數字）→ 雜訊
    if not re.search(r'[一-鿿A-Za-z0-9]', text):
        return True
    return False


def _sanitize_filename(name):
    """OCR 抓的名字可能含特殊字符，移除後當檔名"""
    return re.sub(r'[/\\:*?"<>|\r\n\t]', '_', name).strip() or "unknown"


def _today_subdir(persona):
    """該 persona 今日的子目錄（personas/{persona}/histories/YYYY-MM-DD/）"""
    today = datetime.now().strftime("%Y-%m-%d")
    p = persona_histories_dir(persona) / today
    p.mkdir(parents=True, exist_ok=True)
    return p


def _history_file_path(persona, name):
    """該 persona × 該觀眾的 .txt 路徑（新檔放今日子目錄）"""
    return _today_subdir(persona) / f"{_sanitize_filename(name)}.txt"


def _resolve_history_file(persona, name):
    """找 name 在 persona 內對應的 .txt 檔。

    搜尋順序（只在該 persona 的 histories/ 內找，不會跨 persona）：
      1. 今日子目錄精確
      2. 平級舊檔精確（向後相容）
      3. 該 persona 全目錄遞迴 fuzzy match
      4. 都沒有 → 回今日子目錄路徑（會建新檔）
    """
    histories_root = persona_histories_dir(persona)
    today_path = _history_file_path(persona, name)
    if today_path.exists():
        return today_path

    # 平級舊檔（向後相容）
    legacy = histories_root / f"{_sanitize_filename(name)}.txt"
    if legacy.exists():
        return legacy

    # 正規化 fuzzy match（NFKC + zhconv + 去空白 — 業界 entity resolution 標準）
    def _normalize(s):
        try:
            from 反詐_locate import canonicalize_name
            return canonicalize_name(s).lower()
        except Exception:
            # fallback：保留舊邏輯（去空白 + 小寫）
            s = re.sub(r'[\s　]+', '', s)
            return s.lower()

    target = _normalize(_sanitize_filename(name))
    if not target:
        return today_path

    # 遞迴掃該 persona 的所有子目錄
    for f in histories_root.rglob("*.txt"):
        if _normalize(f.stem) == target:
            # 同對話只 log 一次（避免主迴圈每輪都印導致 log 爆炸）
            log_key = (persona, name, str(f))
            if log_key not in _FUZZY_MATCH_LOGGED:
                _FUZZY_MATCH_LOGGED.add(log_key)
                print(f"[history/{persona}] 模糊匹配：'{name}' → 沿用既有檔 '{f.relative_to(histories_root)}'", flush=True)
            return f

    # 找不到 → 用今日子目錄（建新檔）
    return today_path


def _load_customer_history(persona, name):
    """從 .txt 讀對應觀眾的歷史，回傳 history list（{text, sender, y} 格式）。

    使用 LINE 匯出格式：
        2026.05.05 星期二                    ← 日期分隔行（跳過）
        HH:MM <發送者名稱> <訊息>             ← 標準訊息行
        <延續訊息>                           ← 沒時間戳的行 = 接續上一條訊息

    判斷誰是 me/them：發送者 == 客戶名（檔名）→ them，否則 → me
    """
    path = _resolve_history_file(persona, name)
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
                # 日期分隔行：2026.05.05 星期X → 當「系統 marker」塞 history
                # 讓 Claude 看到時序：「[系統] === 2026.05.16 星期六 ===」
                # 解 5/16 bug：之前直接 continue 跳過，Claude 看不到隔天
                date_match = re.match(r'^(\d{4}\.\d{2}\.\d{2}\s*星期.)', line)
                if date_match:
                    history.append({
                        "text": f"=== {date_match.group(1).strip()} ===",
                        "sender": "system",
                        "y": 0,
                    })
                    last_idx = len(history) - 1
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

def _profile_file_path(persona, name):
    """該 persona × 該觀眾的 profile.json 路徑（跟對應 .txt 同目錄）"""
    txt_path = _resolve_history_file(persona, name)
    return txt_path.parent / f"{_sanitize_filename(name)}.profile.json"


def _resolve_profile_file(persona, name):
    """找該 persona 的 name 對應的 .profile.json（fuzzy match 在該 persona 內）"""
    histories_root = persona_histories_dir(persona)
    expected = _profile_file_path(persona, name)
    if expected.exists():
        return expected

    legacy = histories_root / f"{_sanitize_filename(name)}.profile.json"
    if legacy.exists():
        return legacy

    def _normalize(s):
        try:
            from 反詐_locate import canonicalize_name
            return canonicalize_name(s).lower()
        except Exception:
            s = re.sub(r'[\s　]+', '', s)
            return s.lower()

    target = _normalize(_sanitize_filename(name))
    if not target:
        return expected

    suffix = ".profile.json"
    for f in histories_root.rglob(f"*{suffix}"):
        bare = f.name[:-len(suffix)]
        if _normalize(bare) == target:
            return f
    return expected


def _load_profile(persona, name):
    """載入 profile.json，沒有回 None"""
    path = _resolve_profile_file(persona, name)
    if not path.exists():
        return None
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[profile/{persona}] 讀取 {path.name} 失敗：{e}", flush=True)
        return None


def _save_profile(persona, name, profile):
    """存 profile.json — D-1: 改 atomic write 防斷電/crash 半寫入毀壞"""
    path = _resolve_profile_file(persona, name)
    if not atomic_write_json(profile, path):
        print(f"[profile/{persona}] 寫入 {path.name} 失敗", flush=True)
        return False
    return True


# ============================================================
# Atomic write helper（避免 .json 半寫壞）
# ============================================================
def atomic_write_json(data, path):
    """先寫 .tmp → fsync → os.replace 原子 rename，crash/斷電也不會毀壞原檔（D-1）"""
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    tmp = path.with_suffix(path.suffix + ".tmp")
    try:
        with io.open(tmp, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)  # noqa: atomic-write-internal
            f.flush()
            os.fsync(f.fileno())  # 強制落地磁碟，避免 OS buffer 未刷
        os.replace(str(tmp), str(path))  # OS 保證原子操作
        return True
    except Exception as e:
        print(f"[atomic_write] {path.name} 失敗: {e}", flush=True)
        try:
            tmp.unlink(missing_ok=True)
        except Exception:
            pass
        return False


def atomic_read_json(path, default=None):
    """讀 JSON，檔案不存在/解析失敗回 default"""
    path = Path(path)
    if not path.exists():
        return default
    try:
        with io.open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception as e:
        print(f"[atomic_read] {path.name} 失敗: {e}", flush=True)
        return default


# ============================================================
# LINE 時間戳 parser（將「下午 4:43」「昨天」「星期 X」轉成 datetime）
# ============================================================
_WEEKDAY_NAMES = {
    "星期一": 0, "星期二": 1, "星期三": 2, "星期四": 3,
    "星期五": 4, "星期六": 5, "星期日": 6, "星期天": 6,
}


def parse_line_timestamp(time_str: str, now=None):
    """LINE 列表時間戳 → datetime

    支援：
      「下午 4:43」「下午4:43」  → today 16:43
      「上午 11:30」「11:30」     → today 11:30
      「昨天」                     → yesterday 12:00 (粗估)
      「星期一」「星期二」...      → 上週對應日 12:00
      「2 月 15 日」「02/15」     → today.year-02-15 12:00
      無法解析                    → None
    """
    from datetime import datetime, timedelta
    if now is None:
        now = datetime.now()
    if not time_str:
        return None
    s = time_str.strip().replace(" ", "")

    # ---- 上午/下午 HH:MM ----
    import re
    m = re.match(r"^(上午|下午)?(\d{1,2}):(\d{2})$", s)
    if m:
        ampm, hh, mm = m.group(1), int(m.group(2)), int(m.group(3))
        if ampm == "下午" and hh < 12:
            hh += 12
        elif ampm == "上午" and hh == 12:
            hh = 0
        try:
            return now.replace(hour=hh, minute=mm, second=0, microsecond=0)
        except ValueError:
            return None

    # ---- 昨天 ----
    if s == "昨天":
        y = now - timedelta(days=1)
        return y.replace(hour=12, minute=0, second=0, microsecond=0)

    # ---- 星期 X ----
    if s.startswith("星期") and len(s) >= 3:
        target = _WEEKDAY_NAMES.get(s[:3])
        if target is not None:
            current = now.weekday()
            # 算最近的「上一個」星期 X（若今天就是 → 算上週）
            days_back = (current - target) % 7
            if days_back == 0:
                days_back = 7
            d = now - timedelta(days=days_back)
            return d.replace(hour=12, minute=0, second=0, microsecond=0)

    # ---- 日期 02/15、2/15、2月15日 ----
    m = re.match(r"^(\d{1,2})[月/](\d{1,2})日?$", s)
    if m:
        mo, dy = int(m.group(1)), int(m.group(2))
        try:
            d = datetime(now.year, mo, dy, 12, 0)
            if d > now:  # 跨年（如現在 2026-01，列表顯示「12/15」是去年）
                d = d.replace(year=now.year - 1)
            return d
        except ValueError:
            return None

    return None  # 無法解析


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


def _extract_profile_from_history(persona, name, history):
    """用 Haiku 4.5 從整個 history 抽結構化 profile（首次見觀眾時用）"""
    if not history:
        return None

    me_label = f"[{persona_nickname(persona)}]"
    history_text_lines = []
    for msg in history:
        label = "[對方]" if msg["sender"] == "them" else me_label
        history_text_lines.append(f"{label} {msg['text']}")
    history_text = "\n".join(history_text_lines)

    nick = persona_nickname(persona)
    prompt = (
        f"你是對話分析師。以下是 {nick} 跟對方（{name}）的完整聊天記錄。"
        f"請從中**精確抽出**對方揭露的所有事實，輸出嚴格的 JSON。\n\n"
        f"<對話記錄>\n{history_text}\n</對話記錄>\n\n"
        f"<抽取規則>\n"
        f"1. **只抽對方真的說過的事**，{nick} 說的不算對方的事實\n"
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
        # 🔴 加 timeout 60 秒避免 Haiku 偶爾卡死占住 worker
        client = anthropic.Anthropic(timeout=60.0)
        resp = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=4096,
            messages=[{"role": "user", "content": prompt}]
        )
        text = _strip_json_codeblock(resp.content[0].text)
        profile = json.loads(text)
        # 🔴 schema 校驗：補齊缺漏欄位、補 metadata
        profile = _validate_and_normalize_profile(profile, name=name)
        print(f"[profile/{persona}] {name} 首次抽取完成（{len(profile.get('shared_disclosures') or [])} 筆 disclosures）", flush=True)
        return profile
    except Exception as e:
        print(f"[profile/{persona}] 首次抽取 {name} 失敗：{e}", flush=True)
        return None


def _update_profile_incrementally(persona, name, old_profile, new_messages):
    """每輪 reply 後增量更新 profile（只送舊 profile + 最新對話 → Haiku）

    含 critic 驗證：舊 profile 的 disclosures / core_facts 不能消失
    """
    if not old_profile:
        return None
    if not new_messages:
        return old_profile

    me_label = f"[{persona_nickname(persona)}]"
    new_text_lines = []
    for msg in new_messages:
        label = "[對方]" if msg["sender"] == "them" else me_label
        new_text_lines.append(f"{label} {msg['text']}")
    new_text = "\n".join(new_text_lines)

    old_profile_json = json.dumps(old_profile, ensure_ascii=False, indent=2)

    prompt = (
        f"你是對話分析師。{persona_nickname(persona)} 跟對方（{name}）有最新一輪對話。請**更新** profile（只加新事實，舊的全保留）。\n\n"
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
        # 🔴 加 timeout 60 秒避免 Haiku 偶爾卡死占住 worker
        client = anthropic.Anthropic(timeout=60.0)
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
            print(f"[profile/{persona}] ⚠️ critic 偵測到漏掉舊 disclosures：{missing}，merge 回去", flush=True)
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
                print(f"[profile/{persona}] ⚠️ critic 偵測到 core_facts.{key} 從 '{old_val}' 變空，還原", flush=True)
                new_core[key] = old_val
        new_profile["core_facts"] = new_core

        # 🔴 schema 校驗 + 累計輪數
        new_profile = _validate_and_normalize_profile(new_profile, name=name)
        # total_turns 累加（此次更新算 1 輪）
        new_profile["total_turns"] = (old_profile.get("total_turns") or 0) + 1
        # 保留 first_seen
        if old_profile.get("first_seen"):
            new_profile["first_seen"] = old_profile["first_seen"]

        # 🔴 LRU 淘汰：disclosures 超過 80 筆 → 留最新 80（FIFO，list 已按時序加入）
        # Why: 14 天連跑可能累積 500-1000 筆事實，prompt 注入會肥到 Claude 抓不到重點
        # 核心 core_facts 不動（critic 已保護），只淘汰 shared_disclosures 的舊筆
        DISCLOSURES_MAX = 80
        disclosures = new_profile.get("shared_disclosures") or []
        if len(disclosures) > DISCLOSURES_MAX:
            removed_n = len(disclosures) - DISCLOSURES_MAX
            new_profile["shared_disclosures"] = disclosures[-DISCLOSURES_MAX:]  # 留最新 80
            print(f"[profile/{persona}] 🔴 LRU 淘汰：disclosures {len(disclosures)} → {DISCLOSURES_MAX}（刪最舊 {removed_n} 筆）", flush=True)

        return new_profile
    except Exception as e:
        print(f"[profile/{persona}] 增量更新 {name} 失敗：{e}，沿用舊 profile", flush=True)
        return old_profile


# 🔴 對方檔案 schema 不一致防呆（Haiku 抽取時偶爾寫 dict 而非 str）
# 解 5/15 Natsuki bug：family_relationships 抽成 [{"relation":"父親","fact":"買房"}]
# 但 join 預期 str → TypeError 整個 handle_one_customer crash
def _coerce_item_to_str(item):
    """把 dict / str / 其他型別轉成顯示文字（防 join 爆）"""
    if isinstance(item, str):
        return item
    if isinstance(item, dict):
        relation = item.get("relation") or item.get("role") or item.get("member")
        fact = item.get("fact") or item.get("description")
        if relation and fact:
            return f"{relation}（{fact}）"
        if relation:
            return relation
        if fact:
            return fact
        # 沒命中常見 key，退回 str(dict)
        return str(item)
    return str(item)


def _safe_join(items, sep="、"):
    """把混合 str/dict 清單黏成顯示文字"""
    return sep.join(_coerce_item_to_str(x) for x in items if x)


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
        parts.append(f"- 個性：{_safe_join(traits)}")

    interests = profile.get("interests") or []
    if interests:
        parts.append(f"- 興趣愛好：{_safe_join(interests)}")

    family = profile.get("family_relationships") or []
    if family:
        parts.append(f"- 家庭/關係：{_safe_join(family)}")

    disclosures = profile.get("shared_disclosures") or []
    # 加 isinstance(d, dict) 保險：Haiku 偶爾寫成 str 而非 dict
    them_d = [d for d in disclosures if isinstance(d, dict) and d.get("speaker") == "them"]
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
            parts.append(f"  - {_coerce_item_to_str(m)}")

    stage = profile.get("current_stage")
    if stage:
        parts.append(f"\n**當前階段：** {stage}")

    return "\n".join(parts)


def _ensure_profile(persona, name, history):
    """保證該觀眾有 profile：沒有就抽，並存檔。回傳 profile dict 或 None。"""
    profile = _load_profile(persona, name)
    if profile is None and history:
        print(f"[profile/{persona}] {name} 首次見，從 {len(history)} 條 history 抽 profile...", flush=True)
        profile = _extract_profile_from_history(persona, name, history)
        if profile:
            _save_profile(persona, name, profile)
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
# Why: 2 worker + main loop 並發讀寫 all_profiles[persona] dict，status snapshot iterate 時 worker 寫入會 RuntimeError；GIL 只保證 single op atomic，RMW 必須 lock
_PROFILE_DICT_LOCK = _threading.RLock()


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

        persona = task.get("persona", DEFAULT_PERSONA)
        name = task["name"]
        history_snapshot = task["history"]
        round_msgs = task["round_msgs"]
        all_profiles_ref = task["all_profiles"]  # 已經是該 persona 的 dict

        try:
            with _PROFILE_DICT_LOCK:
                current_profile = all_profiles_ref.get(name)

            if current_profile:
                # 已有 profile → 增量更新
                updated = _update_profile_incrementally(persona, name, current_profile, round_msgs)
                action = "增量更新"
            elif len(history_snapshot) >= MIN_TURNS_FOR_FIRST_PROFILE:
                # 全新觀眾累積夠了 → 首次抽
                print(f"[profile-worker/{persona}] {name} 累積 {len(history_snapshot)} 條，首次抽 profile...", flush=True)
                updated = _extract_profile_from_history(persona, name, history_snapshot)
                action = "首次抽取"
            else:
                # 資料太少，跳過（加 log 方便事後追）
                print(f"[profile-worker/{persona}] {name} history 不足 ({len(history_snapshot)}/{MIN_TURNS_FOR_FIRST_PROFILE})，跳過首次抽取", flush=True)
                _PROFILE_TASK_QUEUE.task_done()
                continue

            if updated:
                with _PROFILE_DICT_LOCK:
                    all_profiles_ref[name] = updated
                _save_profile(persona, name, updated)
                disc_n = len((updated.get("shared_disclosures") or []))
                print(f"[profile-worker/{persona}] {name} {action}完成（{disc_n} 筆 disclosures）", flush=True)
                _emit_event(
                    persona,
                    "profile_first_extract" if action == "首次抽取" else "profile_updated",
                    customer=name,
                    data={"disclosures": disc_n, "action": action},
                )
        except Exception as e:
            err = str(e)
            if "rate" in err.lower() or "429" in err:
                print(f"[profile-worker/{persona}] {name} 撞 rate limit，退避 30 秒後重排：{err[:100]}", flush=True)
                _emit_event(persona, "rate_limit", customer=name, data={"err": err[:200]})
                time.sleep(30)
                try:
                    _PROFILE_TASK_QUEUE.put_nowait(task)
                except _queue_mod.Full:
                    print(f"[profile-worker/{persona}] queue 滿，丟棄 {name} 的 retry", flush=True)
                    _write_dead_letter(persona, name, "queue_full_after_rate_limit", task, err)
            else:
                print(f"[profile-worker/{persona}] {name} 處理失敗：{err[:200]}", flush=True)
                _write_dead_letter(persona, name, "worker_exception", task, err)
                _emit_event(persona, "profile_failed", customer=name, data={"err": err[:200]})
        finally:
            _PROFILE_TASK_QUEUE.task_done()


def _write_dead_letter(persona, name, reason, task, err):
    """profile worker 失敗的任務寫到 persona logs/profile_failures.jsonl"""
    try:
        f = persona_logs_dir(persona) / "profile_failures.jsonl"
        record = {
            "ts": datetime.now().isoformat(timespec="seconds"),
            "persona": persona,
            "customer": name,
            "reason": reason,
            "err": str(err)[:500],
            "round_msgs_count": len(task.get("round_msgs") or []),
            "history_len": len(task.get("history") or []),
        }
        with io.open(f, "a", encoding="utf-8") as fh:
            fh.write(json.dumps(record, ensure_ascii=False) + "\n")
    except Exception:
        pass


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
# 識破偵測：對方說「你是 AI / 機器人 / 不是真人」→ console 警告 + 標記 profile
# ============================================================
# 強質疑短語（含完整指控結構，不會誤報）
_AI_SUSPICION_STRONG = (
    "你是ai", "你是 ai", "你是AI", "你是 AI",
    "你是不是ai", "你是不是 ai", "你是不是AI",
    "你是不是機器人", "你是不是机器人", "你是不是機器", "你是不是机器",
    "你是不是真人", "你是不是真的",
    "你是机器人", "你是機器人", "你是机器", "你是機器",
    "你是程式", "你是程序", "你是bot",
    "你是chatgpt", "你是gpt", "你是claude",
    "你是假的", "假的吧",
    "ai 寫的", "ai寫的", "AI 寫的", "AI寫的", "ai写的", "ai 写的",
    "ai回的", "AI回的", "ai 回的",
    "不是真人", "不是真的人", "不是人吧",
    "回應太快", "回应太快", "回的太快", "回得太快",
    "機器人吧", "机器人吧", "ai吧", "AI吧",
    # 單獨提到這些 AI 名稱在曖昧聊天中 99% 是質疑
    "chatgpt", "ChatGPT", "ChatGpt",
)


def _detect_ai_suspicion(messages):
    """偵測對方訊息中是否有「懷疑 AI」的關鍵字。

    只用「強質疑短語」（含「你是 X」這種完整指控結構），避免誤報：
    - 「AI 智慧科技很厲害」← 沒有「你是」prefix，不會 match
    - 「我看 AI 新聞」← 同上

    回傳：(suspected_bool, matched_quote 或 None)
    """
    for msg_text in messages:
        # 全形空格也要去掉
        low = msg_text.lower().replace(" ", "").replace("　", "")
        for kw in _AI_SUSPICION_STRONG:
            kw_norm = kw.lower().replace(" ", "")
            if kw_norm in low:
                return True, msg_text
    return False, None


# ============================================================
# Event log (buffered append，避免高頻 IO)
# ============================================================
_EVENT_BUFFER = []
_EVENT_BUFFER_LOCK = _threading.Lock()
_EVENT_BUFFER_MAX = 10        # 累積 10 個事件就 flush
_EVENT_BUFFER_FLUSH_SEC = 5   # 或 5 秒沒寫盤就 flush
_EVENT_LAST_FLUSH = [time.time()]  # list 包起來避免 closure 問題


def _flush_event_buffer():
    """把 buffer 內的 event 一次寫盤（thread-safe，按 persona 分檔）"""
    with _EVENT_BUFFER_LOCK:
        if not _EVENT_BUFFER:
            return
        records = list(_EVENT_BUFFER)
        _EVENT_BUFFER.clear()
        _EVENT_LAST_FLUSH[0] = time.time()

    # 按 persona 分組寫到各自 logs 目錄
    by_persona = {}
    for r in records:
        p = r.get("persona", DEFAULT_PERSONA)
        by_persona.setdefault(p, []).append(r)

    today = datetime.now().strftime("%Y-%m-%d")
    for p, recs in by_persona.items():
        try:
            events_file = persona_logs_dir(p) / f"events_{today}.jsonl"
            with io.open(events_file, "a", encoding="utf-8") as f:
                for record in recs:
                    f.write(json.dumps(record, ensure_ascii=False) + "\n")
        except Exception as e:
            print(f"[event/{p}] flush failed: {e}", flush=True)


def _emit_event(persona, event_type, customer="", data=None):
    """關鍵事件 buffered append，10 條或 5 秒 flush 一次（按 persona 分檔）"""
    record = {
        "ts": datetime.now().isoformat(timespec="seconds"),
        "persona": persona,
        "type": event_type,
        "customer": customer,
        "data": data or {},
    }
    with _EVENT_BUFFER_LOCK:
        _EVENT_BUFFER.append(record)
        should_flush = (
            len(_EVENT_BUFFER) >= _EVENT_BUFFER_MAX
            or (time.time() - _EVENT_LAST_FLUSH[0]) >= _EVENT_BUFFER_FLUSH_SEC
        )
    if should_flush:
        _flush_event_buffer()


atexit.register(_flush_event_buffer)


def _enqueue_profile_task(persona, name, history, round_msgs, all_profiles_per_persona):
    """把 profile 任務丟進 queue（非阻塞，queue 滿就寫 dead letter）

    all_profiles_per_persona: 該 persona 自己的 profile dict（不是頂層 dict）
    """
    _ensure_profile_workers_started()
    history_snapshot = list(history)
    task = {
        "persona": persona,
        "name": name,
        "history": history_snapshot,
        "round_msgs": round_msgs,
        "all_profiles": all_profiles_per_persona,
    }
    try:
        _PROFILE_TASK_QUEUE.put_nowait(task)
    except _queue_mod.Full:
        print(f"[profile-worker/{persona}] ⚠️ queue 滿（>200），丟棄 {name} 的更新", flush=True)
        _write_dead_letter(persona, name, "queue_full_at_enqueue", task, "queue full")


# 🔴 cache 首次見面日期（避免每次 status snapshot 都讀整個 .txt）
_FIRST_CHAT_DATE_CACHE = {}


def _get_first_chat_date(persona, name):
    """從 .txt 找最早的日期分隔行（含 cache，key=(persona,name)）"""
    cache_key = (persona, name)
    if cache_key in _FIRST_CHAT_DATE_CACHE:
        return _FIRST_CHAT_DATE_CACHE[cache_key]

    profile = _load_profile(persona, name)
    if profile:
        first_seen = profile.get("first_seen")
        if first_seen:
            try:
                d = datetime.fromisoformat(first_seen).date()
                _FIRST_CHAT_DATE_CACHE[cache_key] = d
                return d
            except (ValueError, TypeError):
                pass

    path = _resolve_history_file(persona, name)
    result = datetime.now().date()
    if path.exists():
        try:
            with io.open(path, "r", encoding="utf-8") as f:
                for line in f:
                    m = re.match(r'^(\d{4})\.(\d{2})\.(\d{2})', line)
                    if m:
                        result = datetime(int(m.group(1)), int(m.group(2)), int(m.group(3))).date()
                        break
        except Exception:
            pass
    _FIRST_CHAT_DATE_CACHE[cache_key] = result
    return result


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

# 🔴 主動發訊息話術庫（按 persona 分開，避免語腔污染）
# 每個 persona 有自己的早安/午餐/晚安/關心/Day5+ 話術
PROACTIVE_MESSAGES = {
    "Angela": {
        "morning": ["morning☺", "早安~ 起來啦?", "早安喇~ 今日忙嗎"],
        "lunch": ["lunch time~", "在吃飯嗎😆", "午飯食緊咩呀"],
        "night": ["準備休息咯~", "好啦我去沖涼了，晚安🌙", "晚安喇~ 明天聊"],
        "concern": ["在嗎？🫣", "忙到沒看訊息嗎？😬", "怎麼那麼久沒回我", "Hello~"],
        "day5_plus": [
            "剛處理完店鋪訂單😬",
            "今晚利潤還不錯~",
            "店鋪剛接到訂單，你那邊在忙啥",
        ],
    },
    "Ruby": {
        "morning": ["早安呀", "早安~", "早安唷", "早安❤️"],
        "lunch": ["午餐吃ㄌ嗎", "在吃飯ㄇ哈哈", "我等等吃午餐"],
        "night": ["晚安😴", "我先休息ㄌ", "晚安❤️ 早點休息ㄛ"],
        "concern": ["在嗎～", "怎麼這麼久沒回我哈哈", "Hello~", "你還活著嗎🥹"],
        # Ruby 沒副業 → Day5+ 改成日常分享
        "day5_plus": [
            "剛跟同事吃完飯",
            "我先洗澡先",
            "今天上班好燒腦😔",
        ],
    },
}


def _get_proactive_messages(persona, key):
    """取該 persona 的指定類型話術。

    若該 persona 沒設定 PROACTIVE_MESSAGES → log warning + 返回空 list（不 fallback 到別 persona）
    避免 Ruby 沒設定時拿 Angela 的粵語話術污染。
    """
    p_msgs = PROACTIVE_MESSAGES.get(persona)
    if p_msgs is None:
        print(f"[⚠️ proactive] persona '{persona}' 沒設定 PROACTIVE_MESSAGES，跳過此次主動發", flush=True)
        return []
    msgs = p_msgs.get(key)
    if not msgs:
        print(f"[⚠️ proactive] persona '{persona}' 沒設定 key='{key}'，跳過", flush=True)
        return []
    return msgs


def _scan_all_customers(persona):
    """掃描該 persona 的 histories/ 遞迴回傳所有觀眾名稱"""
    histories_root = persona_histories_dir(persona)
    if not histories_root.exists():
        return []
    return [p.stem for p in histories_root.rglob("*.txt") if p.is_file()]


def _get_last_message_time(persona, name):
    """從 .txt 算最後一條訊息的時間（datetime），找不到回 None"""
    path = _resolve_history_file(persona, name)
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


def _get_last_sender(persona, name):
    """從 .txt 看最後一條訊息是誰發的（'me' or 'them'），找不到回 None"""
    path = _resolve_history_file(persona, name)
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


def _was_greeting_sent_today(persona, name, greeting_type):
    """今天的這個時段問候是否已發過（看 .txt 內最後幾條 me 訊息）"""
    path = _resolve_history_file(persona, name)
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
    nickname = persona_nickname(persona)
    for msg in today_messages:
        if not msg.split(" ", 1)[1].startswith(nickname):
            continue
        text = msg
        if any(kw in text for kw in keywords):
            return True
    return False


def _check_proactive_trigger(persona, name):
    """檢查觀眾是否需要主動發訊息（針對特定 persona）

    🔴 2026-05-16 全部停用（用戶要求）：
    - 拿掉「沉默 2/4/6 小時關心」（SILENCE_THRESHOLD）
    - 拿掉「時段問候」（GREETING_SLOTS 早安 / 午餐 / 晚安）
    - 腳本只回覆對方傳來的訊息，不主動搜尋好友 / 發訊
    要重啟：把下面這一行 return 拿掉即可，後面邏輯保留
    """
    return False, "", "Proactive 已停用（用戶要求 2026-05-16）"
    # 以下邏輯保留但不執行
    if name.startswith(_BLACKLIST_PREFIXES):
        return False, "", ""
    if _is_quiet_hour():
        return False, "", "深夜時段"

    # Cooldown：嘗試發過（無論成功失敗）30 分鐘內不再對同客戶觸發
    key = (persona, name)
    last_attempt = _LAST_PROACTIVE_ATTEMPT.get(key, 0)
    elapsed = time.time() - last_attempt
    if elapsed < PROACTIVE_COOLDOWN_SEC:
        return False, "", f"cooldown（上次嘗試 {elapsed/60:.1f} 分鐘前）"

    last_msg_time = _get_last_message_time(persona, name)
    if last_msg_time is None:
        return False, "", "沒有歷史"

    last_sender = _get_last_sender(persona, name)
    now = datetime.now()
    hours_silent = (now - last_msg_time).total_seconds() / 3600

    if last_sender == "me" and hours_silent < 1:
        return False, "", "剛發過訊息"

    h = now.hour
    for s, e, gtype in GREETING_SLOTS:
        in_slot = (s <= h < e) if s <= e else (h >= s or h < e)
        if in_slot and not _was_greeting_sent_today(persona, name, gtype):
            day_n = _calculate_day_n(persona, name)
            candidates = _get_proactive_messages(persona, gtype)
            if not candidates:
                continue
            msg = random.choice(candidates)
            return True, msg, f"時段問候 {gtype}（Day {day_n}）"

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
            day_n = _calculate_day_n(persona, name)
            if day_n >= 5 and random.random() < 0.4:
                cands = _get_proactive_messages(persona, "day5_plus")
                if cands:
                    return True, random.choice(cands), f"Day {day_n} 深化話題（沉默 {hours_silent:.1f} 小時）"
            cands = _get_proactive_messages(persona, "concern")
            if cands:
                return True, random.choice(cands), f"沉默關心（{hours_silent:.1f} 小時 ≥ {threshold} 小時）"
            return False, "", "沒可用話術"

    return False, "", "不需主動發"


def _send_proactive_message(persona, name, message, monitor=None):
    """主動發訊息給特定觀眾（針對該 persona）"""
    from 反詐_locate import (
        locate_line_regions, switch_page, search_friend_and_scan,
        enter_chat_from_search,
    )
    from 反詐_chat import send_reply

    print(f"[Proactive/{persona}] 主動發給 {name}: {message[:50]}", flush=True)

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
        regions = enter_chat_from_search(friend_pos, regions, monitor)
        time.sleep(1)

        # 3. 模擬「思考延遲」（短一點，主動發應該比較快）
        # 用一半的時段延遲（30-90 秒 → 15-45 秒）
        think_delay = get_current_delay() * 0.5
        print(f"[Proactive] 主動發前等 {think_delay:.1f} 秒（模擬打字）", flush=True)
        if not interruptible_sleep(think_delay):
            return False  # 被中斷

        # 4. 送訊息（主動發通常很短，不分段）
        send_reply(message, regions)
        print(f"[Proactive] → {message}", flush=True)

        # 5. 寫進 .txt
        _append_to_history_file(persona, name, "me", message)

        # 6. 退出
        time.sleep(1)
        pyautogui.press("escape")
        time.sleep(0.5)

        return True
    except Exception as e:
        print(f"[Proactive/{persona}] 主動發給 {name} 失敗: {e}", flush=True)
        try:
            pyautogui.press("escape")
        except Exception:
            pass
        return False


def _calculate_day_n(persona, name):
    """計算當前是跟對方認識第幾天（從第一次對話算起）"""
    first_date = _get_first_chat_date(persona, name)
    today = datetime.now().date()
    return (today - first_date).days + 1


# 🔴 不同 persona 走不同的階段策略（Angela 走副業伏筆 / Ruby 走日常深化）
STAGE_HINTS_BY_PERSONA = {
    "Angela": {
        "1-2":  "剛認識/寒暄階段。重點：親切但不黏人，留懸念（如『去沖涼喇 明天再聊』）。不要急著深入。",
        "3-4":  "關係加深階段。重點：互換家庭背景、製造『巧合』（星座/食物/位置）、表現專一（『其他軟體都沒在用了』）。",
        "5-6":  "埋下副業伏筆。可以隨口提『去處理店鋪訂單』。對方好奇 → 簡單描述跨境電商。不要主動解釋細節。",
        "7-8":  "副業細節+情感升溫。可以秀利潤（『半個月 14 萬』）、規劃見面、暗示好感。",
        "9+":   "情感確認階段。表達對對方好感、規劃未來、用感情包裝（『等你賺到請我吃大餐』）。",
    },
    "Ruby": {
        "1-2":  "剛認識/寒暄階段。重點：可愛系開場（噗/哈哈哈/早安呀），不黏人，晚安固定『我先休息ㄌ』。不要急著深入。",
        "3-4":  "關係加深。露 1-2 個生活細節（玉米蛋餅梗 / 周星馳粉 / 1.25 倍速看陸劇）。共鳴升級（沒錯沒錯/真的/我也是）。",
        "5-6":  "信任建立。提媽媽煮晚餐 / 跟同事吃飯 / 上班壓力。允許小撒嬌（加分加分❤️）。",
        "7-8":  "更深生活分享。聊買房規劃 / 工作未來 / 假設一起去玩（『以後想去高雄玩 你陪我嗎』）。",
        "9+":   "情感確認 + 邊界階段。可以露副業話題鉤子（『剛在整理店舖的東西』），但只在對方主動問才答。⚠️ 不切到 day 14+ 切錢手法。",
    },
}


def _build_stage_hint(persona, day_n):
    """根據 persona + Day N 產出對話階段提示"""
    stages = STAGE_HINTS_BY_PERSONA.get(persona) or STAGE_HINTS_BY_PERSONA["Angela"]
    if day_n <= 2:
        stage = stages["1-2"]
    elif day_n <= 4:
        stage = stages["3-4"]
    elif day_n <= 6:
        stage = stages["5-6"]
    elif day_n <= 8:
        stage = stages["7-8"]
    else:
        stage = stages["9+"]
    return (
        f"\n\n=== 當前對話階段 ===\n"
        f"你跟對方認識 **Day {day_n}**（從你們最初聯絡那天起算）。\n"
        f"階段策略：{stage}\n"
        f"嚴禁跳階段：Day 1 不要突然提敏感話題，Day 9 才能講感情。"
    )


def _today_date_header():
    """產生今天日期分隔行：'2026.05.05 星期二'"""
    now = datetime.now()
    return f"{now.strftime('%Y.%m.%d')} {_WEEKDAY_TW[now.weekday()]}"


# C-6: cache 每個 (persona, name) 的最後 date header，避免每次 append 都 read 整檔
_LAST_DATE_HEADER_CACHE = {}
_LAST_DATE_HEADER_LOCK = _threading.Lock()


def _append_to_history_file(persona, name, sender, text):
    """即時把一條訊息 append 到觀眾的 .txt（LINE 匯出格式）"""
    path = _resolve_history_file(persona, name)
    nickname = persona_nickname(persona) if sender == "me" else name
    timestamp = datetime.now().strftime("%H:%M")
    today_header = _today_date_header()
    cache_key = (persona, str(path))

    try:
        new_file = not path.exists()
        need_date_header = new_file

        with _LAST_DATE_HEADER_LOCK:
            cached_header = _LAST_DATE_HEADER_CACHE.get(cache_key)

        if not new_file:
            if cached_header is None:
                # cache miss → read 整檔取最後 date header（per-customer 一次性成本）
                with io.open(path, "r", encoding="utf-8") as f:
                    content = f.read()
                date_lines = re.findall(r'^\d{4}\.\d{2}\.\d{2}\s*星期.', content, re.MULTILINE)
                cached_header = date_lines[-1] if date_lines else None
                if cached_header:
                    with _LAST_DATE_HEADER_LOCK:
                        _LAST_DATE_HEADER_CACHE[cache_key] = cached_header
            if cached_header != today_header:
                need_date_header = True

        with io.open(path, "a", encoding="utf-8") as f:
            if need_date_header:
                if not new_file:
                    f.write("\n")  # 跨日空行分隔
                f.write(f"{today_header}\n")
            f.write(f"{timestamp} {nickname} {text}\n")

        # 寫成功後更新 cache（跨日後第一次 append 也會更新成今日 header）
        with _LAST_DATE_HEADER_LOCK:
            _LAST_DATE_HEADER_CACHE[cache_key] = today_header
    except Exception as e:
        print(f"[history/{persona}] 寫入 {path.name} 失敗：{e}", flush=True)


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


def _send_with_realistic_delay(reply, regions, persona=DEFAULT_PERSONA, name="unknown", history=None):
    """分段送 reply（段間延遲 + 字數比例打字時間），送一段立刻寫一段。

    persona: 哪個人設（決定圖片庫位置 + 寫到哪個 .txt）
    """
    from 反詐_chat import send_reply, send_image, send_audio_via_drop

    parts = [p.strip() for p in reply.split("|||") if p.strip()]
    for i, part in enumerate(parts):
        # 圖片標籤 {send_自拍} / {send_食物}
        m = re.match(r'^\{(send_[\w一-鿿]+)\}$', part)
        # 音訊標籤 {audio_早安} / {audio_笑聲}
        m_audio = re.match(r'^\{(audio_[\w一-鿿]+)\}$', part)
        if m_audio:
            tag = m_audio.group(1)
            audio_path = _resolve_audio_tag(persona, tag)
            if audio_path:
                send_audio_via_drop(audio_path, regions)
                print(f"[Reply/{persona}] → 🎤 {Path(audio_path).name}（{tag}）", flush=True)
                audio_marker = f"[音訊:{Path(audio_path).name}]"
                if history is not None:
                    history.append({"text": audio_marker, "sender": "me", "y": 0})
                _append_to_history_file(persona, name, "me", audio_marker)
            else:
                print(f"[Reply/{persona}] ⚠ {tag} 找不到對應音檔，跳過", flush=True)
        elif m:
            tag = m.group(1)
            img_path = _resolve_image_tag(persona, tag)
            if img_path:
                send_image(img_path, regions)
                print(f"[Reply/{persona}] → 📷 {Path(img_path).name}（{tag}）", flush=True)
                img_marker = f"[圖片:{Path(img_path).name}]"
                if history is not None:
                    history.append({"text": img_marker, "sender": "me", "y": 0})
                _append_to_history_file(persona, name, "me", img_marker)
            else:
                print(f"[Reply/{persona}] ⚠ {tag} 找不到對應圖片，跳過", flush=True)
        elif part.startswith("{") and part.endswith("}"):
            print(f"[Reply/{persona}] 跳過佔位符: {part}", flush=True)
        else:
            send_reply(part, regions)
            print(f"[Reply/{persona}] → {part[:60]}", flush=True)
            if history is not None:
                history.append({"text": part, "sender": "me", "y": 0})
            _append_to_history_file(persona, name, "me", part)

        if i < len(parts) - 1:
            # 🔴 段間延遲 = 基礎延遲 + 「下一段字數比例」打字時間
            # 模擬真人：邊想（基礎）+ 邊打字（字數越多打越久）
            next_part = parts[i + 1] if i + 1 < len(parts) else ""
            is_img_tag = next_part.startswith("{") and next_part.endswith("}")
            # 短段（< 5 字）走快速延遲 [2, 5]，長段才用設定的 [5, 12]
            if len(next_part) < 5 and not is_img_tag:
                base = random.uniform(2, 5)
            else:
                base = get_inter_message_delay()
            # 排除圖片標記（{send_xxx}）的字數計算
            if not is_img_tag:
                # 每字 0.3-0.5 秒（手機打字速度）
                typing_time = len(next_part) * random.uniform(0.3, 0.5)
                # 上限 30 秒（避免長段卡死）
                typing_time = min(typing_time, 30)
            else:
                typing_time = 0
            total_delay = base + typing_time
            print(f"[Reply/{persona}] 段間延遲 {total_delay:.1f}s（基礎 {base:.1f}s + 打字 {typing_time:.1f}s）", flush=True)
            if not interruptible_sleep(total_delay):
                return  # 收到 stop signal，中止剩下段落

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
    """印 startup state（按 persona 分組顯示累積資料）"""
    try:
        print(f"\n[Startup] 累積觀眾資料（按 persona）：", flush=True)
        for persona_name, cfg in PERSONA_CONFIG.items():
            histories_root = persona_histories_dir(persona_name)
            txt_files = list(histories_root.rglob("*.txt"))
            json_files = list(histories_root.rglob("*.profile.json"))
            active_str = "active" if cfg.get("active") else "inactive"
            print(f"  [{persona_name}] {active_str}: {len(txt_files)} txt / {len(json_files)} profile", flush=True)
            # stage 分佈
            stage_count = {}
            for jf in json_files:
                try:
                    with io.open(jf, "r", encoding="utf-8") as f:
                        p = json.load(f)
                    stage = (p.get("current_stage") or "未知")[:30]
                    stage_count[stage] = stage_count.get(stage, 0) + 1
                except Exception:
                    continue
            for stage, n in sorted(stage_count.items(), key=lambda x: -x[1])[:5]:
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
    print(f"\n[STOP] 收到信號 {signum}，主迴圈將跳出（atexit 接手 drain）...", flush=True)
    # Why: signal handler 內不做 blocking I/O（業界 best practice），drain 留給 atexit.register(_graceful_shutdown) 兜底

signal.signal(signal.SIGINT, _signal_handler)
# Windows 沒原生 SIGTERM 但 Python 模擬，schtasks /End 走 WM_CLOSE 不會觸發；Linux 才實用
signal.signal(signal.SIGTERM, _signal_handler)
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


def interruptible_sleep(seconds):
    """可中斷的 sleep — 每秒 check should_stop，避免長 sleep 內 stop flag 不被偵測

    回傳：True = 正常 sleep 完，False = 被中斷
    """
    if seconds <= 1:
        time.sleep(seconds)
        return not should_stop()
    elapsed = 0.0
    while elapsed < seconds:
        if should_stop():
            return False
        chunk = min(1.0, seconds - elapsed)
        time.sleep(chunk)
        elapsed += chunk
    return True

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

def find_unread_with_metadata(monitor=None):
    """直接拿 DBSCAN 卡 + 卡綁徽章的結果，不再啟發式配對

    新流程（2026-05-07，砍掉 LINE16 寫死掃描）：
      1. locate_line_regions → DBSCAN 切卡（每張 1 筆，含 name/y/center/time_str）
      2. find_unread_per_card → 對每張卡的右側檢查綠色 → has_unread
      3. 過濾 has_unread=True → 拿 name/center/time_str 直接用

    優點：
      - name 跟卡 1:1 對應（不再巧合配對到錯名字）
      - center 來自卡（不再寫死 y=157, 228 → LINE 動態重排序也對人）
      - time_str 來自卡（不再啟發式猜）

    回傳：(regions, [{y, center, name, last_msg_time, last_msg_time_str}])
    """
    from datetime import datetime
    from 反詐_locate import locate_line_regions, find_unread_per_card

    regions = locate_line_regions(monitor)
    cards = find_unread_per_card(regions, monitor)

    enriched = []
    now = datetime.now()
    for c in cards:
        if not c.get("has_unread"):
            continue
        time_str = c.get("time_str")
        time_dt = parse_line_timestamp(time_str, now) if time_str else None
        enriched.append({
            "y": c["y"],
            "center": c["center"],
            "name": c["name"],
            "last_msg_time_str": time_str,
            "last_msg_time": time_dt,
        })

    return regions, enriched



# ============================================================
# Per-Customer Scheduler — 解 Head-of-Line blocking
# 每個 customer 獨立計時，不會因第一個 sleep 19 分鐘卡死後面所有
# ============================================================
class PendingCustomerScheduler:
    """事件驅動 Per-Customer Scheduler（業界 Outbox Pattern）

    Key 設計：(persona, name)
      - name = LINE 列表偵測到的暱稱（OCR）
      - 暱稱穩定 → 跨 LINE 列表浮動都能對應同 customer

    Timer 邏輯：
      first_seen: datetime（從 LINE 時間戳抓 = 對方真實傳訊息時間）
                  fallback 用 datetime.now()（OCR 沒抓到時間戳時）
      target_delay: 抽到的思考延遲（秒）
      ready_at = first_seen + target_delay

    Persistence：
      pending_queue.json 寫進 personas/{persona}/
      atomic write 防 crash 半寫壞
      啟動時 load 回 in-memory state
    """
    def __init__(self):
        self._pending = {}  # {(persona, name): state}
        self._lock = _threading.Lock()

    @staticmethod
    def _key(persona, name):
        return (persona, name)

    @staticmethod
    def _persistence_path(persona):
        return persona_root(persona) / "pending_queue.json"

    def register_if_new(self, persona, name, last_msg_time, target_delay):
        """事件驅動：新 customer 註冊計時器，已存在不動。
        last_msg_time: datetime（對方傳訊息時間，從 LINE 列表抓）
                       None 時 fallback 用 datetime.now()
        target_delay: 秒
        回傳 True=新註冊
        """
        from datetime import datetime
        key = self._key(persona, name)
        first_seen = last_msg_time if last_msg_time else datetime.now()
        with self._lock:
            if key in self._pending:
                return False
            self._pending[key] = {
                "persona": persona,
                "name": name,
                "first_seen": first_seen,  # datetime
                "target_delay": target_delay,
                "ready_at": first_seen + _td_seconds(target_delay),
                "first_detected": datetime.now(),
            }
        self._persist(persona)
        return True

    def get_ready(self):
        """回傳所有「等夠了」的 customer，按 first_seen 排序（FIFO 公平）"""
        from datetime import datetime
        now = datetime.now()
        with self._lock:
            ready = [
                (k, v) for k, v in self._pending.items()
                if v["ready_at"] <= now
            ]
        ready.sort(key=lambda x: x[1]["first_seen"])
        return ready

    def remove(self, key):
        """處理完該 customer，清掉 pending"""
        persona = key[0] if isinstance(key, tuple) and len(key) >= 1 else None
        with self._lock:
            self._pending.pop(key, None)
        if persona:
            self._persist(persona)

    def remove_by_name(self, persona, name):
        """用 (persona, name) 移除"""
        self.remove(self._key(persona, name))

    def cleanup_expired(self, ttl_factor=5):
        """TTL 清理：超過 target_delay × ttl_factor 還沒處理 → 丟棄"""
        from datetime import datetime
        now = datetime.now()
        affected_personas = set()
        with self._lock:
            expired = []
            for k, v in self._pending.items():
                age_seconds = (now - v["first_seen"]).total_seconds()
                if age_seconds > v["target_delay"] * ttl_factor:
                    expired.append(k)
            for k in expired:
                state = self._pending.pop(k)
                age = (now - state["first_seen"]).total_seconds()
                print(f"[Sched] TTL 過期清掉 {k}（等了 {age:.0f}s）", flush=True)
                affected_personas.add(state["persona"])
        for p in affected_personas:
            self._persist(p)

    def get_active_names(self, persona):
        """回傳該 persona 內所有 pending customer 的 name（用於 diff）"""
        with self._lock:
            return {k[1] for k in self._pending if k[0] == persona}

    def stats(self):
        with self._lock:
            return {
                "pending_count": len(self._pending),
                "by_persona": {p: sum(1 for k in self._pending if k[0] == p)
                              for p in set(k[0] for k in self._pending)},
            }

    def _persist(self, persona):
        """寫盤該 persona 的 pending_queue.json（atomic）"""
        try:
            from datetime import datetime
            with self._lock:
                customers = []
                for k, v in self._pending.items():
                    if k[0] != persona:
                        continue
                    customers.append({
                        "name": v["name"],
                        "first_seen": v["first_seen"].isoformat(),
                        "target_delay": v["target_delay"],
                        "ready_at": v["ready_at"].isoformat(),
                        "first_detected": v["first_detected"].isoformat(),
                    })
            data = {
                "persona": persona,
                "last_updated": datetime.now().isoformat(timespec="seconds"),
                "customers": customers,
            }
            atomic_write_json(data, self._persistence_path(persona))
        except Exception as e:
            print(f"[Sched] persist {persona} 失敗: {e}", flush=True)

    def load_from_disk(self, persona):
        """啟動時從 pending_queue.json 載入 in-memory（crash recovery）"""
        from datetime import datetime
        path = self._persistence_path(persona)
        data = atomic_read_json(path, default=None)
        if not data or not data.get("customers"):
            return 0
        loaded = 0
        with self._lock:
            for c in data["customers"]:
                try:
                    name = c["name"]
                    first_seen = datetime.fromisoformat(c["first_seen"])
                    target_delay = c["target_delay"]
                    ready_at = datetime.fromisoformat(c["ready_at"])
                    first_detected = datetime.fromisoformat(c.get("first_detected", c["first_seen"]))
                    self._pending[(persona, name)] = {
                        "persona": persona,
                        "name": name,
                        "first_seen": first_seen,
                        "target_delay": target_delay,
                        "ready_at": ready_at,
                        "first_detected": first_detected,
                    }
                    loaded += 1
                except Exception as e:
                    print(f"[Sched] load {persona}/{c.get('name')} 失敗: {e}", flush=True)
        return loaded


def _td_seconds(seconds):
    """秒 → timedelta"""
    from datetime import timedelta
    return timedelta(seconds=seconds)


# 全域 scheduler 單例
_PENDING_SCHEDULER = PendingCustomerScheduler()

# Event-driven diff 偵測：上一輪每個 persona 看到的 unread name 集合
_LAST_SEEN_UNREAD = {}   # {persona: set(names)}

# Backpressure：兩次回覆最少間隔（避免大量 ready 連續點擊被 LINE 風控）
MIN_INTER_REPLY_SEC = 30
_LAST_REPLY_TS = 0.0


def _write_heartbeat(persona, stats=None):
    """每輪寫 personas/{persona}/heartbeat.json，外部可監控腳本是否還活著"""
    try:
        from datetime import datetime
        path = persona_root(persona) / "heartbeat.json"
        data = {
            "persona": persona,
            "last_heartbeat": datetime.now().isoformat(timespec="seconds"),
            "pending_count": (stats or {}).get("pending_count", 0),
        }
        atomic_write_json(data, path)
    except Exception:
        pass  # heartbeat 失敗不該阻塞主流程


# ============================================================
# 處理單一客戶（演示版：簡化、純自然對話、無業務邏輯）
# ============================================================
def handle_one_customer(persona, conv, regions, system_prompt, all_histories, all_profiles, monitor=None):
    """處理單一觀眾的對話（針對特定 persona）

    ⚠️ 此函式假設「思考延遲已經在外部 PendingCustomerScheduler 等過了」
    → 不再做 think_delay sleep（避免 head-of-line blocking）
    → 直接點進去處理

    all_histories / all_profiles: 該 persona 自己的 dict（兩層 dict 的內層）
    """
    from 反詐_locate import locate_line_regions, ocr_scan_panel, screenshot_line, wait_for_region_change, safe_click
    from 反詐_chat import (
        is_only_sticker, analyze_sticker,
    )
    from difflib import SequenceMatcher

    cx, cy = conv["center"]
    target_name = conv.get("name", "unknown")

    if should_stop():
        return regions

    # 點擊前抓 chat_title region — 用於 hash 比對等待
    ct_before = regions.get("chat_title", {}) if regions else {}
    ct_bbox = None
    if ct_before:
        l = ct_before.get("left", 0)
        t = ct_before.get("top", 0)
        r = ct_before.get("right", 0)
        b = ct_before.get("bottom", 0)
        if r > l and b > t:
            ct_bbox = (l, t, r - l, b - t)

    # 🔴 A 方案：點擊前先看當前對話標題列是不是已是目標 → 是 → 跳過點擊（省 6 秒 retry 浪費）
    # Why: 同一個人連續傳訊息時，畫面已經停在他的對話，safe_click 的 verify 找不到變化 → 誤判失敗 retry 3 次
    already_open = False
    if ct_bbox:
        try:
            from 反詐_locate import canonicalize_name as _canon
            full_img, line_crop, (il, it, ir, ib), mon = screenshot_line(monitor)
            sx_r = full_img.size[0] / mon["width"]
            sy_r = full_img.size[1] / mon["height"]
            _ct_il = max(0, int((ct_before.get("left", 0) - mon["left"]) * sx_r) - il)
            _ct_it = max(0, int((ct_before.get("top", 0) - mon["top"]) * sy_r) - it)
            _ct_ir = min(line_crop.size[0], int((ct_before.get("right", 0) - mon["left"]) * sx_r) - il)
            _ct_ib = min(line_crop.size[1], int((ct_before.get("bottom", 0) - mon["top"]) * sy_r) - it)
            if _ct_ir > _ct_il and _ct_ib > _ct_it:
                _title_crop = line_crop.crop((_ct_il, _ct_it, _ct_ir, _ct_ib))
                _title_items = ocr_scan_panel(_title_crop)
                if _title_items:
                    _current_title = _title_items[0].get("text", "").strip()
                    if _canon(_current_title) == _canon(target_name):
                        already_open = True
                        print(f"[Customer/{persona}] '{target_name}' 已是當前對話（chat_title 已就位）→ 跳過點擊", flush=True)
        except Exception as e:
            print(f"[Customer/{persona}] pre-check chat_title 失敗 ({e}) → 改走 safe_click", flush=True)

    if not already_open:
        # safe_click：retry + idempotency + verify + logging（業界 RPA 2026 標準）
        # action_id 含 timestamp → 每次處理都是新 action（不會誤合併）
        from datetime import datetime as _dt
        action_id = f"open_chat_{target_name}_{_dt.now().strftime('%Y%m%d_%H%M%S')}"

        print(f"\n[Customer] 點擊未讀對話 '{target_name}' at ({cx}, {cy})", flush=True)
        success = safe_click(
            cx, cy,
            verify_region=ct_bbox,        # 點完 chat_title 區域要變化
            action_id=action_id,
            max_retries=3,
            verify_timeout=2.5,
            persona=persona,
        )
        if not success:
            # 全部 retry 都失敗 → log 已寫到 click_failures.jsonl
            print(f"[Customer/{persona}] ⚠️ safe_click 失敗（已 retry 3 次），跳過 '{target_name}'", flush=True)
            # 不 return regions，下面 locate_line_regions 還是試試看

        # 沒 ct_bbox 時 safe_click 不 verify，補個寫死 sleep 確保 LINE 載完
        if not ct_bbox:
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
                _emit_event(persona, "ocr_failed", customer=name, data={"err": err_str})
                return regions
    if current_messages is None:
        return regions

    # 🆕 階段 2：偵測語音訊息 → 點 ▶ → 錄音 → 轉文字 → 替換訊息 text
    # 「儲存|另存新檔|分享|傳送至Keep筆記」是 LINE 語音氣泡下的選單列指紋
    # 替換在過濾雜訊之前做（否則指紋會被當系統訊息殺掉）
    try:
        from 反詐_voice import find_voice_markers, detect_play_button, record_loopback, transcribe_wav
        import threading
        import numpy as np
        voice_markers = find_voice_markers(current_messages)
        if voice_markers:
            print(f"[Customer/{persona}] 發現 {len(voice_markers)} 條語音訊息，開始轉文字", flush=True)
            arr = np.array(chat_img)
            ca = regions["chat_area"]
            for idx, marker in enumerate(voice_markers):
                play_x, play_y, method = detect_play_button(arr, marker["marker_y"])
                if play_x is None:
                    print(f"[Voice/{persona}] ⚠ 找不到 ▶ marker_y={marker['marker_y']} → 跳過", flush=True)
                    for m in current_messages:
                        if m.get("y") == marker["marker_y"] and m.get("sender") == "them":
                            m["text"] = "[語音 — 無法定位播放鈕]"
                            break
                    continue
                screen_x = ca["left"] + play_x
                screen_y = ca["top"] + play_y
                print(f"[Voice/{persona}] #{idx+1} 點 ▶ ({method}) chat=({play_x},{play_y}) screen=({screen_x},{screen_y})", flush=True)

                text_holder = [""]
                def _record_task(label):
                    try:
                        wav = record_loopback(15, label=label)
                        text_holder[0] = transcribe_wav(wav)
                    except Exception as e:
                        print(f"[Voice/{persona}] 錄音/STT 失敗：{type(e).__name__}: {e}", flush=True)

                t = threading.Thread(target=_record_task, args=(f"voice_{idx}",), daemon=True)
                t.start()
                time.sleep(0.3)  # 等錄音 stream ready
                try:
                    safe_click(screen_x, screen_y, verify_timeout=2.5)
                except Exception as e:
                    print(f"[Voice/{persona}] safe_click 失敗：{e}", flush=True)
                t.join(timeout=20)

                # 替換訊息 text
                for m in current_messages:
                    if m.get("y") == marker["marker_y"] and m.get("sender") == "them":
                        if text_holder[0]:
                            m["text"] = f"[語音轉文字] {text_holder[0]}"
                            print(f"[Voice/{persona}] #{idx+1} → 「{text_holder[0]}」", flush=True)
                        else:
                            m["text"] = "[語音 — 轉文字失敗]"
                        break
    except ImportError as e:
        print(f"[Voice/{persona}] 模組 import 失敗（先略過語音處理）：{e}", flush=True)
    except Exception as e:
        print(f"[Voice/{persona}] 語音處理異常：{type(e).__name__}: {e}", flush=True)

    raw_count = len(current_messages)
    # 🆕 過濾 OCR 雜訊（LINE 系統訊息、按鈕、海報等）
    current_messages = [m for m in current_messages if not _is_ocr_noise(m["text"])]
    if len(current_messages) < raw_count:
        print(f"[Customer] OCR 讀到 {raw_count} 條 → 過濾雜訊後 {len(current_messages)} 條", flush=True)
    else:
        print(f"[Customer] OCR 讀到 {len(current_messages)} 條訊息", flush=True)

    if name not in all_histories:
        loaded = _load_customer_history(persona, name)
        if loaded:
            print(f"[Customer/{persona}] {name} 載入 {len(loaded)} 條歷史", flush=True)
        else:
            print(f"[Customer/{persona}] {name} 全新觀眾", flush=True)
        _emit_event(persona, "customer_seen", customer=name, data={"prior_history_len": len(loaded)})
        all_histories[name] = loaded
        cached_profile = _load_profile(persona, name) if loaded else None
        if cached_profile:
            all_profiles[name] = cached_profile
            print(f"[Customer/{persona}] {name} profile cache 命中（{len(cached_profile.get('shared_disclosures') or [])} 筆）", flush=True)
        else:
            all_profiles[name] = None
            if loaded and len(loaded) >= MIN_TURNS_FOR_FIRST_PROFILE:
                print(f"[Customer/{persona}] {name} 無 profile，背景排程首次抽", flush=True)
                _enqueue_profile_task(persona, name, loaded, [], all_profiles)
    history = all_histories[name]
    profile = all_profiles.get(name)

    if not history:
        new_them = [m["text"] for m in current_messages if m["sender"] == "them"]
        for m in current_messages:
            history.append(m)
        for m in current_messages:
            if m["sender"] == "them":
                _append_to_history_file(persona, name, "them", m["text"])
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
                print(f"[Customer/{persona}] cutoff 用 history[-{matched_via}] 對到", flush=True)
            new_msgs = current_messages[match_idx + 1:]
            new_them = [m["text"] for m in new_msgs if m["sender"] == "them"]
            for m in new_msgs:
                history.append(m)
            for m in new_msgs:
                _append_to_history_file(persona, name, m["sender"], m["text"])
        else:
            print(f"[Customer/{persona}] ⚠️ history[-5:] 都對不上 OCR，保守取最後 1 條對方訊息", flush=True)
            them_msgs = [m for m in current_messages if m["sender"] == "them"]
            new_them = [them_msgs[-1]["text"]] if them_msgs else []

    if not new_them:
        print(f"[Customer/{persona}] 沒有新的對方訊息，跳過", flush=True)
        return regions

    print(f"[Customer/{persona}] 新訊息: {new_them}", flush=True)

    # 識破偵測（fingerprint 去重）
    suspected, suspect_quote = _detect_ai_suspicion(new_them)
    if suspected:
        cur_profile = all_profiles.get(name)
        existing_quotes = set()
        if cur_profile:
            existing_quotes = {f.get("quote") for f in (cur_profile.get("ai_suspicion_flags") or []) if f.get("quote")}
        if suspect_quote not in existing_quotes:
            print(f"[Customer/{persona}] ⚠️⚠️ {name} 疑似識破 AI！對方說：「{suspect_quote}」", flush=True)
            _emit_event(persona, "ai_suspicion", customer=name, data={"quote": suspect_quote})
            if cur_profile:
                flags = cur_profile.get("ai_suspicion_flags") or []
                flags.append({
                    "timestamp": datetime.now().isoformat(timespec="seconds"),
                    "quote": suspect_quote,
                })
                cur_profile["ai_suspicion_flags"] = flags
                _save_profile(persona, name, cur_profile)

    # 純貼圖 → Vision
    if is_only_sticker(new_them):
        meaning = analyze_sticker(regions, monitor)
        print(f"[Customer/{persona}] 純貼圖 → Vision 解讀為「{meaning}」", flush=True)
        new_them = [f"[貼圖含意：{meaning}]"]
    else:
        from 反詐_chat import analyze_recent_photo
        photo_desc = analyze_recent_photo(regions, monitor)
        if photo_desc:
            print(f"[Customer/{persona}] 偵測到對方傳照片 → Vision 描述: {photo_desc}", flush=True)
            new_them = list(new_them) + [f"[對方傳了一張照片：{photo_desc}]"]

    day_n = _calculate_day_n(persona, name)
    stage_hint = _build_stage_hint(persona, day_n)
    opponent_location = _detect_opponent_location(name, history)
    world_context = _get_world_context(persona, opponent_location)
    print(f"[Customer/{persona}] {name} 對話 Day {day_n}, 對方所在地: {opponent_location or '未知'}", flush=True)
    augmented_prompt = (
        system_prompt + stage_hint +
        "\n\n=== 今天的世界資訊（給你即時參考，避免破綻）===\n" + world_context
    )
    profile_text = _format_profile_for_prompt(profile) if profile else ""
    if profile_text:
        print(f"[Customer/{persona}] {name} profile 注入（{len((profile or {}).get('shared_disclosures') or [])} 筆事實）", flush=True)

    reply = generate_reply(augmented_prompt, history, new_them, profile_text=profile_text, persona_name=persona_nickname(persona))
    if not reply or len(reply) <= 1:
        print(f"[Customer/{persona}] ⚠️ generate_reply 回空（thinking 被切斷或無內容），跳過 {name}", flush=True)
        time.sleep(0.5)
        return regions

    reply = filter_reply(reply)
    if not reply:
        print(f"[Customer/{persona}] ⚠️ filter_reply 過濾後為空（只有系統訊息/分析格式），跳過 {name}", flush=True)
        return regions

    print(f"[Customer/{persona}] 回覆: {reply[:80]}", flush=True)
    _emit_event(persona, "new_messages", customer=name, data={"new_them_count": len(new_them), "day_n": day_n})
    _send_with_realistic_delay(reply, regions, persona=persona, name=name, history=history)
    _emit_event(persona, "reply_sent", customer=name, data={"reply_len": len(reply), "segments": reply.count("|||") + 1, "day_n": day_n})

    round_msgs = [{"sender": "them", "text": t} for t in new_them]
    round_msgs.append({"sender": "me", "text": reply})
    _enqueue_profile_task(persona, name, history, round_msgs, all_profiles)

    # ============================================================
    # 🆕 Stay-and-watch：送完 reply 後不要立刻離開，再 OCR 看有沒有新訊息
    # 對方在「思考延遲」或「分段送出」期間插話 → 立刻接續回覆，避免斷層
    # 最多 3 輪、每輪等 5 秒讓對方有時間反應
    # ============================================================
    from difflib import SequenceMatcher as _SM
    MAX_FOLLOWUP = 3
    WATCH_WAIT = 5  # 秒

    for followup_round in range(MAX_FOLLOWUP):
        if not interruptible_sleep(WATCH_WAIT):
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
        for m in new_msgs2:
            history.append(m)
            _append_to_history_file(persona, name, m["sender"], m["text"])

        # 再生成 reply 接續回應（也帶 profile）
        # 用 cache 裡最新的 profile（背景 thread 可能已經更新完）
        latest_profile = all_profiles.get(name) or profile
        latest_profile_text = _format_profile_for_prompt(latest_profile) if latest_profile else ""
        reply2 = generate_reply(augmented_prompt, history, new_them2, profile_text=latest_profile_text, persona_name=persona_nickname(persona))
        if not reply2 or len(reply2) <= 1:
            print(f"[Customer/{persona}] ⚠️ 接續輪 {followup_round+1} generate_reply 回空，離開", flush=True)
            break
        reply2 = filter_reply(reply2)
        if not reply2:
            print(f"[Customer/{persona}] ⚠️ 接續輪 {followup_round+1} filter_reply 為空，離開", flush=True)
            break

        print(f"[Customer/{persona}] 接續輪 {followup_round+1} 回覆: {reply2[:80]}", flush=True)
        _send_with_realistic_delay(reply2, regions, persona=persona, name=name, history=history)

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

    # 🔴 啟動前自動跑 audit（找 persona 隔離 violation，有 P0 拒絕啟動）
    # C-5: 加 retry — transient timeout/IO 失敗不該讓 14 天連跑掉一天，但 returncode != 0 (P0 found) 仍拒絕啟動
    import subprocess as _sp
    audit_script = Path("C:/Users/blue_/.claude/scripts/audit/audit.py")
    if audit_script.exists():
        print("\n[Audit] 啟動前 persona 隔離稽核...", flush=True)
        r = None
        for attempt in range(3):
            try:
                r = _sp.run(
                    [sys.executable, str(audit_script), "--config", "反詐demo_persona"],
                    capture_output=True, text=True, encoding="utf-8", errors="replace",
                    timeout=30,
                )
                break  # 取得 result（無論 returncode 為何）
            except (_sp.TimeoutExpired, OSError) as e:
                wait = 2 ** attempt
                print(f"[Audit] ⚠️ 第 {attempt+1}/3 次嘗試失敗 ({type(e).__name__})，{wait}s 後重試", flush=True)
                time.sleep(wait)
            except Exception as e:
                print(f"[Audit] ⚠️ audit 跑失敗（非 transient）: {e}（跳過繼續啟動）", flush=True)
                r = None
                break

        if r is None:
            print("[Audit] ⚠️ 3 次 retry 後仍失敗（跳過繼續啟動，避免 14 天連跑斷天）", flush=True)
        elif r.returncode == 0:
            print("[Audit] ✅ 0 P0 violation，啟動", flush=True)
        else:
            print(r.stdout, flush=True)
            print("\n[Audit] ❌ 偵測到 P0 violation，拒絕啟動。請先修 violation。", flush=True)
            print("[Audit] 跳過 audit 啟動：python 反詐_multi.py <time> --skip-audit", flush=True)
            if "--skip-audit" not in sys.argv:
                return False
    else:
        print(f"[Audit] ⚠️ audit script 不存在（{audit_script}），跳過稽核", flush=True)

    # 載入時段延遲設定
    _load_time_settings(TIME_SETTINGS_PATH)

    # 🔴 載入所有 active persona 的 prompt
    persona_prompt_cache = {}
    for persona_name, cfg in get_active_personas().items():
        try:
            persona_prompt_cache[persona_name] = load_persona_prompt(persona_persona_file(persona_name))
            print(f"[Init] persona={persona_name} (sandbox={cfg['sandbox']}, nick={cfg['nickname']}) 人設載入成功 ({len(persona_prompt_cache[persona_name])} chars)", flush=True)
        except Exception as e:
            print(f"[Init] persona={persona_name} 載入人設失敗: {e}", flush=True)

    if not persona_prompt_cache:
        print("[ERROR] 沒有任何 active persona，請檢查 PERSONA_CONFIG", flush=True)
        return False

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

    # 🔴 兩層 dict：{persona: {customer_name: ...}}，避免跨 persona 撞名
    all_histories = {p: {} for p in persona_prompt_cache}
    all_profiles = {p: {} for p in persona_prompt_cache}
    POLL_INTERVAL = 10

    _print_startup_recovery()

    # 🔴 crash recovery：載入 pending_queue.json
    for p in persona_prompt_cache:
        loaded = _PENDING_SCHEDULER.load_from_disk(p)
        if loaded > 0:
            print(f"[Sched/{p}] 從 pending_queue.json 載回 {loaded} 個 pending customer", flush=True)

    print(f"\n[Monitor] 開始監控未讀訊息（active personas: {list(persona_prompt_cache.keys())}）...", flush=True)
    print(f"[Monitor] 停止方式：touch {STOP_FILE}", flush=True)

    from 反詐_locate import set_active_box, find_line_window
    import win32gui as _w32g

    # 🔴 GPU 記憶體週期釋放（每 GPU_CLEAN_INTERVAL 個主迴圈跑一次）
    # Why: 從 50 縮到 30，配合 paddle 3.x leak 行為（每 50 OCR call 後 idle pool 顯著膨脹）；30 loop ≈ ~200 OCR call ≈ 5 分鐘
    GPU_CLEAN_INTERVAL = 30
    STATUS_PRINT_INTERVAL = 6  # ~ 6 × 10 秒 = 約每 1 分鐘印 status
    _loop_counter = 0

    while is_before_stop_time(stop_time):
        if should_stop():
            break

        # ============================================================
        # 🔴 Event-Driven Scheduler 模型（業界 Outbox + Stable Identifier + Wall-Clock）
        # 每輪：① 偵測 unread → diff 找新增 → 用 LINE 時間戳註冊 timer
        #       ② 找 ready 的 customer 用 name 找對應對話框處理（含 backpressure）
        #       ③ TTL 清理
        # ============================================================
        for persona_name, persona_cfg in get_active_personas().items():
            if should_stop():
                break
            try:
                sandbox = persona_cfg["sandbox"]
                set_active_box(sandbox)
                persona_prompt = persona_prompt_cache.get(persona_name)
                if not persona_prompt:
                    continue

                line = find_line_window()
                if not line:
                    continue
                hwnd = line[0]
                try:
                    _w32g.SetForegroundWindow(hwnd)
                except Exception:
                    pass
                time.sleep(1.0)

                # ① 偵測所有 unread + 抓暱稱 + 抓時間戳
                regions, enriched_unread = find_unread_with_metadata(monitor)
                current_names = {u["name"] for u in enriched_unread if u["name"]}

                # diff: 新增的（事件驅動）
                last_names = _LAST_SEEN_UNREAD.get(persona_name, set())
                new_names = current_names - last_names
                gone_names = last_names - current_names

                # 處理新增：註冊 timer
                for u in enriched_unread:
                    name = u["name"]
                    if name in new_names:
                        delay = get_current_delay()
                        last_msg_time = u.get("last_msg_time")  # datetime or None
                        if _PENDING_SCHEDULER.register_if_new(persona_name, name, last_msg_time, delay):
                            time_label = u.get("last_msg_time_str") or "未知時間"
                            print(f"[Sched/{persona_name}] 新未讀 '{name}'（對方 {time_label} 傳的），等 {delay:.0f}s 後處理", flush=True)

                # 處理消失：清掉 token（對方手動已讀 / 我們已處理）
                for name in gone_names:
                    _PENDING_SCHEDULER.remove_by_name(persona_name, name)
                    print(f"[Sched/{persona_name}] '{name}' 已不在 LINE 未讀清單，清 token", flush=True)

                _LAST_SEEN_UNREAD[persona_name] = current_names

                if new_names or gone_names:
                    stats = _PENDING_SCHEDULER.stats()
                    print(f"[Sched] 全域 pending: {stats['pending_count']} 個（{stats['by_persona']}）", flush=True)

                # 寫 heartbeat（每輪都寫，外部監控用）
                _write_heartbeat(persona_name, stats=_PENDING_SCHEDULER.stats())

            except Exception as e:
                print(f"[ERR/{persona_name}] 偵測階段失敗: {e}", flush=True)
                import traceback
                traceback.print_exc()
                time.sleep(2)

        # ② 處理 ready 的 customer（用 name 找對話框 + backpressure）
        # 🔴 Batch wall-clock budget：5 分鐘上限，避免單一 batch 卡死整個 scheduler（A2 已給 LLM 加 120s timeout，最壞 4min/customer，5min batch 至少跑 1 個）
        ready = _PENDING_SCHEDULER.get_ready()
        _batch_start = time.monotonic()
        _BATCH_BUDGET_SEC = 300
        _batch_processed = 0
        for key, state in ready:
            if should_stop():
                break
            if time.monotonic() - _batch_start > _BATCH_BUDGET_SEC:
                print(f"[Sched] batch 超過 {_BATCH_BUDGET_SEC}s（已處理 {_batch_processed}/{len(ready)}），跳出讓主 loop 重排剩 {len(ready)-_batch_processed} 個", flush=True)
                break
            persona_name = state["persona"]
            target_name = state["name"]

            # 🔴 Backpressure：兩次點擊最少間隔
            global _LAST_REPLY_TS
            now_mono = time.monotonic()
            wait_left = MIN_INTER_REPLY_SEC - (now_mono - _LAST_REPLY_TS)
            if wait_left > 0:
                print(f"[Sched/{persona_name}] backpressure 等 {wait_left:.0f}s（防點擊過快）", flush=True)
                if not interruptible_sleep(wait_left):
                    break

            try:
                sandbox = PERSONA_CONFIG[persona_name]["sandbox"]
                set_active_box(sandbox)
                line = find_line_window()
                # C-9: hwnd validity check，避免 stale handle 點空（LINE 視窗 minimize/close）
                if not line or not win32gui.IsWindow(line[0]):
                    print(f"[Sched/{persona_name}] LINE 視窗 handle 失效，跳過 '{target_name}' 等下輪重抓", flush=True)
                    _PENDING_SCHEDULER.remove(key)
                    continue
                try:
                    _w32g.SetForegroundWindow(line[0])
                except Exception:
                    pass
                time.sleep(0.5)

                # 重新掃 LINE，用 name 找對應對話框
                regions, enriched_unread = find_unread_with_metadata(monitor)
                target_conv = next((u for u in enriched_unread if u["name"] == target_name), None)

                if not target_conv:
                    # 對方不在 LINE 未讀清單（已讀/封鎖）→ 清 token
                    print(f"[Sched/{persona_name}] ⚠️ ready 但 '{target_name}' 不在 unread 清單，清 token", flush=True)
                    _PENDING_SCHEDULER.remove(key)
                    continue

                persona_prompt = persona_prompt_cache.get(persona_name)
                from datetime import datetime
                elapsed = (datetime.now() - state["first_seen"]).total_seconds()
                print(f"[Sched/{persona_name}] ✅ 處理 '{target_name}'（等了 {elapsed:.0f}s，queue 還 {len(ready)-1} 個）", flush=True)

                handle_one_customer(
                    persona_name, target_conv, regions, persona_prompt,
                    all_histories[persona_name], all_profiles[persona_name],
                    monitor,
                )
                pyautogui.press("escape")
                time.sleep(0.5)
                _PENDING_SCHEDULER.remove(key)
                # 🔴 修 diff 漏網 bug：處理完後從「上輪名單」拿掉這個 name
                # Why: 仁輝若在「處理完→徽章消失前」又傳新訊息，徽章不會真的消失
                #       下次偵測 current 跟 last 都有他 → diff 判定「老面孔跳過」漏處理
                #       主動 discard 讓下次徽章重現算「新出現」
                _LAST_SEEN_UNREAD.get(persona_name, set()).discard(target_name)
                _LAST_REPLY_TS = time.monotonic()
                _batch_processed += 1
            except Exception as e:
                print(f"[ERR/{persona_name}] 處理 '{target_name}' 失敗: {e}", flush=True)
                import traceback
                traceback.print_exc()
                _PENDING_SCHEDULER.remove(key)  # 失敗就清掉避免卡住（remove 是幂等，不再二次呼叫）
                _LAST_SEEN_UNREAD.get(persona_name, set()).discard(target_name)
                time.sleep(2)

        # ③ TTL 清理：超過 target_delay × 5 還沒處理 → 對方應該封鎖/刪聊天
        _PENDING_SCHEDULER.cleanup_expired(ttl_factor=5)

        # ============================================================
        # 🆕 排程器：每輪 box 處理完後，檢查是否該主動發訊息給某些觀眾
        # 觸發條件：時段問候 / 沉默太久關心 / Day 5+ 副業伏筆
        # ============================================================
        if not should_stop():
            try:
                for persona_name, persona_cfg in get_active_personas().items():
                    sandbox = persona_cfg["sandbox"]
                    set_active_box(sandbox)
                    line2 = find_line_window()
                    if not line2:
                        continue
                    try:
                        _w32g.SetForegroundWindow(line2[0])
                    except Exception:
                        pass
                    time.sleep(0.5)

                    all_customers = _scan_all_customers(persona_name)
                    if not all_customers:
                        continue

                    proactive_sent = 0
                    MAX_PROACTIVE_PER_ROUND = 3
                    for cust_name in all_customers:
                        if proactive_sent >= MAX_PROACTIVE_PER_ROUND:
                            break
                        if should_stop():
                            break
                        should_send, msg, reason = _check_proactive_trigger(persona_name, cust_name)
                        if not should_send:
                            continue
                        print(f"[Proactive/{persona_name}] {cust_name} 觸發：{reason}", flush=True)
                        # 嘗試前先記 timestamp（無論成功失敗都 cooldown 30 分鐘，避免重複）
                        _LAST_PROACTIVE_ATTEMPT[(persona_name, cust_name)] = time.time()
                        ok = _send_proactive_message(persona_name, cust_name, msg, monitor)
                        if ok:
                            proactive_sent += 1
                            time.sleep(2)
                        else:
                            print(f"[Proactive/{persona_name}] {cust_name} 失敗，{PROACTIVE_COOLDOWN_SEC//60} 分鐘內不再嘗試", flush=True)

            except Exception as e:
                print(f"[ERR] 排程器失敗: {e}", flush=True)
                import traceback
                traceback.print_exc()

        # 🔴 status snapshot 每分鐘印一次（按 persona 分組）
        _loop_counter += 1
        if _loop_counter % STATUS_PRINT_INTERVAL == 0:
            try:
                queue_n = _PROFILE_TASK_QUEUE.qsize() if _PROFILE_WORKER_STARTED else 0
                lines = [f"[Status] global queue={queue_n}"]
                for persona_name in get_active_personas():
                    # 🔴 snapshot 避免 iterate during write 撞 worker RuntimeError
                    with _PROFILE_DICT_LOCK:
                        p_hist_keys = list((all_histories.get(persona_name) or {}).keys())
                        p_prof_values = list((all_profiles.get(persona_name) or {}).values())
                    active_n = len(p_hist_keys)
                    profile_n = sum(1 for v in p_prof_values if v)
                    day_dist = {}
                    for cn in p_hist_keys:
                        try:
                            dn = _calculate_day_n(persona_name, cn)
                            day_dist[dn] = day_dist.get(dn, 0) + 1
                        except Exception:
                            continue
                    susp_n = sum(
                        len((v.get("ai_suspicion_flags") or []))
                        for v in p_prof_values if v
                    )
                    day_str = " ".join(f"day{d}:{n}" for d, n in sorted(day_dist.items()))
                    lines.append(f"  {persona_name}: active={active_n} profile={profile_n} suspicion={susp_n} | {day_str}")
                print("\n".join(lines), flush=True)
            except Exception as e:
                print(f"[Status] 失敗：{e}", flush=True)

        # 🔴 GPU 記憶體週期釋放（避免 PaddleOCR 長時間 leak）
        # 業界 best practice: del → gc.collect → empty_cache 三連，cycle ref 要兩輪 collect 才能完全釋放
        if _loop_counter % GPU_CLEAN_INTERVAL == 0:
            try:
                gc.collect()
                gc.collect()  # 第二輪解 cycle reference
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
    total_n = sum(len(d) for d in all_histories.values())
    print(f"共處理 {total_n} 個觀眾", flush=True)
    return True


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        print("用法：python 反詐_multi.py <監控時間HH:MM>")
        sys.exit(0)

    stop = sys.argv[1]
    main(stop)
