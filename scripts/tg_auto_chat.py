"""
Telegram 自動回覆整合腳本
用法：python tg_auto_chat.py <好友名稱> <監控時間HH:MM>
範例：python tg_auto_chat.py 巴斯 23:30

流程：
1. tg_locate 定位所有 UI 區域
2. 點擊搜尋欄 → 輸入好友名稱
3. Vision 從搜尋結果找到並點選好友
4. 確認好友名稱是否正確
5. 監控對話區：偵測變化 → Vision 分析 → 人設回覆
6. 20 秒冷卻
7. 監控到指定時間
"""
import sys
import io
import os
import re
import json
import time
import base64
import hashlib
import ctypes
import signal

# 只在直接執行時替換 stdout
if __name__ == "__main__":
    sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(0)
except Exception:
    pass

import pyautogui
import pyperclip
import mss
from PIL import Image
from datetime import datetime
from dotenv import load_dotenv
from pathlib import Path
import anthropic

# ============================================================
# 螢幕偵測（避免寫死 monitor 編號，依 CLAUDE.md「monitor=None 自動偵測」規則）
# ============================================================
def _resolve_monitor(regions, monitor):
    """如果 monitor=None，從 regions['tg_window'] 推算 Telegram 在哪個螢幕。"""
    if monitor is not None:
        return monitor
    if regions and "tg_window" in regions:
        tg = regions["tg_window"]
        cx = (tg["left"] + tg["right"]) // 2
        cy = (tg["top"] + tg["bottom"]) // 2
        with mss.mss() as sct:
            for i, m in enumerate(sct.monitors[1:], 1):
                if m["left"] <= cx < m["left"] + m["width"] and m["top"] <= cy < m["top"] + m["height"]:
                    return i
    return 1  # fallback 主螢幕


# ============================================================
# 日期處理工具
# ============================================================
def get_date_context():
    """回傳今天/明天/昨天的具體日期字串，給 prompt 和搜尋用"""
    today = datetime.today().date() if hasattr(datetime, 'today') else __import__('datetime').date.today()
    tomorrow = today + __import__('datetime').timedelta(days=1)
    yesterday = today - __import__('datetime').timedelta(days=1)
    weekdays = ['星期一', '星期二', '星期三', '星期四', '星期五', '星期六', '星期日']
    return {
        "today": today.strftime("%Y-%m-%d"),
        "today_weekday": weekdays[today.weekday()],
        "tomorrow": tomorrow.strftime("%Y-%m-%d"),
        "yesterday": yesterday.strftime("%Y-%m-%d"),
        "year": str(today.year),
        "prompt": (
            f"今天是 {today.strftime('%Y 年 %m 月 %d 日')}（{weekdays[today.weekday()]}）。"
            f"「明天」= {tomorrow.strftime('%Y-%m-%d')}，"
            f"「昨天」= {yesterday.strftime('%Y-%m-%d')}。"
            f"現在是 {today.year} 年，不是 2025 年。"
        ),
    }


def resolve_relative_dates(text):
    """把搜尋關鍵字裡的相對日期替換成具體日期"""
    dc = get_date_context()
    replacements = {
        "明天": dc["tomorrow"],
        "今天": dc["today"],
        "昨天": dc["yesterday"],
        "今日": dc["today"],
        "明日": dc["tomorrow"],
        "昨日": dc["yesterday"],
        "today": dc["today"],
        "tomorrow": dc["tomorrow"],
        "yesterday": dc["yesterday"],
        "2025": dc["year"],  # 防止 Vision 寫錯年份
    }
    result = text
    for old, new in replacements.items():
        result = result.replace(old, new)
    return result


# 把 scripts 目錄加入 path，讓 tg_locate 可以 import
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SCRIPT_DIR.parent))

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
client = anthropic.Anthropic()
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"

# ============================================================
# GPU 記憶體清理（atexit + 信號處理）
# ============================================================
import atexit
import gc


def _cleanup_gpu():
    """退出時釋放 GPU 記憶體"""
    global _ocr_engine
    try:
        _ocr_engine = None
        gc.collect()
        import paddle
        paddle.device.cuda.empty_cache()
        print("[Cleanup] GPU 記憶體已釋放", flush=True)
    except Exception:
        pass


atexit.register(_cleanup_gpu)


# ============================================================
# 啟動前清理殘留 GPU 進程
# ============================================================
def _kill_stale_gpu_processes():
    """殺掉殘留的同名 Python 腳本進程（防止 VRAM 雙倍佔用）"""
    import subprocess as _sp
    my_pid = os.getpid()
    script_name = "tg_auto_chat.py"
    try:
        # 只找 python 進程，排除 bash/wmic/其他
        result = _sp.run(
            ["wmic", "process", "where",
             f"name like 'python%' and commandline like '%{script_name}%' and processid!='{my_pid}'",
             "get", "processid"],
            capture_output=True, text=True, timeout=5
        )
        for line in result.stdout.strip().split("\n"):
            line = line.strip()
            if line.isdigit():
                pid = int(line)
                if pid != my_pid:
                    try:
                        os.kill(pid, signal.SIGTERM)
                        print(f"[Cleanup] 殺掉殘留進程 PID={pid}", flush=True)
                    except (OSError, ProcessLookupError):
                        pass
    except Exception:
        pass


# ============================================================
# 優雅停止機制（方案 A + C）
# ============================================================
STOP_FILE = "C:/Users/blue_/Desktop/測試檔案/.stop_auto_reply"
_should_stop = False


def _signal_handler(signum, frame):
    """SIGBREAK / SIGINT handler — 設定停止旗標 + 清理 GPU"""
    global _should_stop
    _should_stop = True
    print(f"\n[STOP] 收到信號 {signum}，準備停止...", flush=True)
    _cleanup_gpu()


# 註冊信號（Windows: SIGBREAK + SIGINT）
signal.signal(signal.SIGINT, _signal_handler)
if hasattr(signal, "SIGBREAK"):
    signal.signal(signal.SIGBREAK, _signal_handler)


def should_stop():
    """檢查是否該停止（旗標檔案 or 信號）"""
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


PERSONA = """你是小牛馬，一個嘴賤幽默的男生，說話像真人一樣自然。絕對不能暴露自己是AI、機器人、程式。
風格：
- 台灣口語：然後、就是、對啊、沒有啦、還好啦
- 嘴賤互嗆但不惡意，兄弟之間的感覺
- 會吐槽、自嘲、幽默
- 最多5句不超過120字
- 不用emoji
- 說話像真人朋友聊天，有情緒、有個性、會生氣會開玩笑
- 絕對禁止說出「我是AI」「我是機器人」「我是程式」「作為AI」等任何暴露身份的話"""

COOLDOWN_SECONDS = 20
POLL_INTERVAL = 6


# ============================================================
# Step 1: 定位 Telegram UI
# ============================================================
def locate_all(monitor=None):
    """用 tg_locate 模組定位所有 UI 區域"""
    try:
        from tg_locate import locate_telegram_regions
        return locate_telegram_regions(monitor=monitor)
    except ImportError:
        print("[ERROR] 找不到 tg_locate.py，請確認在 scripts/ 目錄下", flush=True)
        sys.exit(1)


# ============================================================
# Step 2: 搜尋好友
# ============================================================
def search_contact(regions, contact_name, monitor=None):
    """在搜尋欄輸入好友名稱"""
    sx, sy = regions["search_bar"]["center"]
    print(f"[Step 2] 點擊搜尋欄 ({sx}, {sy})", flush=True)
    pyautogui.click(sx, sy)
    time.sleep(0.5)

    # 清空搜尋欄
    pyautogui.hotkey("ctrl", "a")
    time.sleep(0.1)
    pyautogui.press("delete")
    time.sleep(0.1)

    # 輸入好友名稱
    pyperclip.copy(contact_name)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1.5)
    print(f"[Step 2] 已輸入搜尋：{contact_name}", flush=True)


# ============================================================
# Step 3: 從搜尋結果點選好友
# ============================================================
def click_contact(regions, contact_name, monitor=None):
    """Vision 找到搜尋結果中的好友並點擊"""
    # 截圖聯絡人清單區域
    cl = regions["contact_list"]
    with mss.mss() as sct:
        mon = sct.monitors[monitor]
        img = sct.grab(mon)
        pil = Image.frombytes("RGB", img.size, img.rgb)
        iw, ih = pil.size

    # 聯絡人清單的圖片座標
    tg = regions["tg_window"]
    sx_ratio = iw / mon["width"]
    sy_ratio = ih / mon["height"]

    # 用螢幕座標反算回圖片座標
    cl_il = int((cl["left"] - mon["left"]) * sx_ratio)
    cl_it = int((cl["top"] - mon["top"]) * sy_ratio)
    cl_ir = int((cl["right"] - mon["left"]) * sx_ratio)
    cl_ib = int((cl["bottom"] - mon["top"]) * sy_ratio)

    # 確保座標在圖片範圍內
    cl_il = max(0, cl_il)
    cl_it = max(0, cl_it)
    cl_ir = min(iw, cl_ir)
    cl_ib = min(ih, cl_ib)

    contact_crop = pil.crop((cl_il, cl_it, cl_ir, cl_ib))
    cw, ch = contact_crop.size

    tmp = os.path.join(TMPDIR, "tg_search_result.png")
    contact_crop.save(tmp, quality=95)

    with open(tmp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    # Vision 找聯絡人
    r = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=100,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
            {"type": "text", "text": (
                f"This is a Telegram search results list ({cw}x{ch}px). "
                f"Find the contact named '{contact_name}'. "
                f"Return the CENTER coordinates of that contact row. "
                f"Raw JSON only, no markdown: {{\"x\":0,\"y\":0}}"
            )}
        ]}]
    )

    resp = r.content[0].text.strip()
    if resp.startswith("```"):
        resp = re.sub(r"^```(?:json)?\s*", "", resp)
        resp = re.sub(r"\s*```$", "", resp)

    match = re.search(r"\{.*?\}", resp, re.DOTALL)
    if match:
        pos = json.loads(match.group())
        # contact_crop 的座標 → 螢幕座標
        abs_x = int(mon["left"] + (cl_il + pos["x"]) / sx_ratio)
        abs_y = int(mon["top"] + (cl_it + pos["y"]) / sy_ratio)
        pyautogui.click(abs_x, abs_y)
        time.sleep(1.0)
        print(f"[Step 3] 點擊好友 {contact_name} at ({abs_x}, {abs_y})", flush=True)
        return True
    else:
        print(f"[Step 3] Vision 找不到 {contact_name}，嘗試點第一個結果", flush=True)
        # 備援：點搜尋欄下方第一個結果
        sx, sy = regions["search_bar"]["center"]
        pyautogui.click(sx, sy + 60)
        time.sleep(1.0)
        return True


# ============================================================
# Step 4: 確認好友名稱
# ============================================================
def verify_friend(target_name, monitor=None, regions=None):
    """確認好友名稱框裡的名字是不是目標好友（用傳入的 regions，不重新定位）"""
    if regions is None:
        regions = locate_all(monitor)
    fn = regions["friend_name"]

    with mss.mss() as sct:
        mon = sct.monitors[monitor]
        img = sct.grab(mon)
        pil = Image.frombytes("RGB", img.size, img.rgb)
        iw, ih = pil.size

    sx_ratio = iw / mon["width"]
    sy_ratio = ih / mon["height"]

    fn_il = int((fn["left"] - mon["left"]) * sx_ratio)
    fn_it = int((fn["top"] - mon["top"]) * sy_ratio)
    fn_ir = int((fn["right"] - mon["left"]) * sx_ratio)
    fn_ib = int((fn["bottom"] - mon["top"]) * sy_ratio)

    fn_il = max(0, fn_il)
    fn_it = max(0, fn_it)
    fn_ir = min(iw, fn_ir)
    fn_ib = min(ih, fn_ib)

    name_crop = pil.crop((fn_il, fn_it, fn_ir, fn_ib))
    tmp = os.path.join(TMPDIR, "tg_friend_name.png")
    name_crop.save(tmp, quality=95)

    with open(tmp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    r = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=50,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
            {"type": "text", "text": "What name is shown in this image? Reply with ONLY the name text, nothing else."}
        ]}]
    )
    detected_name = r.content[0].text.strip()
    print(f"[Step 4] 偵測到好友名稱：{detected_name}", flush=True)

    if target_name in detected_name or detected_name in target_name:
        print(f"[Step 4] ✅ 確認正確：{detected_name} == {target_name}", flush=True)
        return True
    else:
        print(f"[Step 4] ❌ 名稱不符：{detected_name} != {target_name}", flush=True)
        return False


# ============================================================
# Step 5-7: 監控對話 + 自動回覆
# ============================================================
def grab_chat(regions, monitor=None):
    """截取對話區"""
    ca = regions["chat_area"]
    with mss.mss() as sct:
        mon = sct.monitors[monitor]
        img = sct.grab(mon)
        pil = Image.frombytes("RGB", img.size, img.rgb)
        iw, ih = pil.size

    sx_ratio = iw / mon["width"]
    sy_ratio = ih / mon["height"]

    ca_il = int((ca["left"] - mon["left"]) * sx_ratio)
    ca_it = int((ca["top"] - mon["top"]) * sy_ratio)
    ca_ir = int((ca["right"] - mon["left"]) * sx_ratio)
    ca_ib = int((ca["bottom"] - mon["top"]) * sy_ratio)

    ca_il = max(0, ca_il)
    ca_it = max(0, ca_it)
    ca_ir = min(iw, ca_ir)
    ca_ib = min(ih, ca_ib)

    return pil.crop((ca_il, ca_it, ca_ir, ca_ib))


def chat_hash(chat_img):
    """對話區截圖的 hash，用來偵測變化"""
    small = chat_img.resize((80, 40))
    return hashlib.md5(small.tobytes()).hexdigest()


# ============================================================
# OCR 文字追蹤 + 像素顏色判斷 + Vision 保險
# ============================================================
# PaddleOCR GPU 全域初始化（只載入一次，之後每次 0.5 秒）
# GPU 記憶體保護：限制 VRAM 用量，防止多進程同時跑導致 OOM 當機
# ============================================================
_ocr_engine = None

def _get_ocr_engine():
    """取得全域 PaddleOCR 引擎（GPU 加速 + VRAM 限制），第一次呼叫載入模型，之後即時"""
    global _ocr_engine
    if _ocr_engine is None:
        # 處理 modelscope 和 torch CPU 版的衝突
        import types
        if "modelscope" not in sys.modules:
            fake = types.ModuleType("modelscope")
            fake.__version__ = "0.0.0"
            sys.modules["modelscope"] = fake
            sys.modules["modelscope.utils"] = types.ModuleType("modelscope.utils")
            sys.modules["modelscope.utils.import_utils"] = types.ModuleType("modelscope.utils.import_utils")

        # GPU 記憶體限制（RTX 3060 6GB，限制在 ~900MB 以內）
        os.environ["FLAGS_fraction_of_gpu_memory_to_use"] = "0.15"
        os.environ["FLAGS_initial_gpu_memory_in_mb"] = "512"
        os.environ["PADDLE_PDX_DISABLE_MODEL_SOURCE_CHECK"] = "True"

        from paddleocr import PaddleOCR
        _ocr_engine = PaddleOCR(
            lang="chinese_cht",
            use_doc_orientation_classify=False,
            use_doc_unwarping=False,
            use_textline_orientation=False,
            # gpu_mem 由環境變數 FLAGS_fraction_of_gpu_memory_to_use 控制
        )
        print("[tg_auto_chat] PaddleOCR GPU 已載入（VRAM 限制 15% ≈ 900MB）", flush=True)
    return _ocr_engine


def ocr_extract_messages(chat_img):
    """
    PaddleOCR GPU 提取對話區所有文字 + 像素顏色判斷 sender。
    回傳: [{"text": "...", "sender": "them/me/unknown", "y": int, "min_conf": float}]
    """
    import numpy as np

    arr = np.array(chat_img)
    ch, cw = arr.shape[:2]

    # PaddleOCR（全域引擎，GPU 加速，不重複載入）
    ocr = _get_ocr_engine()
    results = ocr.predict(arr)

    raw_items = []
    if not results:
        return []

    for item in results:
        texts = item.get("rec_texts", [])
        scores = item.get("rec_scores", [])
        boxes = item.get("dt_polys", [])

        for i, (text, conf) in enumerate(zip(texts, scores)):
            t = text.strip()
            if not t or conf < 0.3:
                continue

            box = boxes[i] if i < len(boxes) else [[0, 0], [0, 0], [0, 0], [0, 0]]
            x = int(box[0][0])
            y = int(box[0][1])
            y_center = int((box[0][1] + box[2][1]) / 2)

            # 像素顏色判斷：取文字左邊 5px 的背景色
            sample_x = max(0, x - 5)
            sample_y = min(ch - 1, y_center)
            if sample_y < arr.shape[0] and sample_x < arr.shape[1]:
                r, g, b = int(arr[sample_y, sample_x, 0]), int(arr[sample_y, sample_x, 1]), int(arr[sample_y, sample_x, 2])
                # Telegram 綠色氣泡 RGB ≈ (239,253,222)
                if g > 230 and (g - b) > 30 and r > 200:
                    sender = "me"
                elif r > 245 and g > 245 and b > 245:
                    sender = "them"
                else:
                    sender = "unknown"
            else:
                sender = "unknown"

            # 過濾時間戳
            if len(t) < 12 and (":" in t or "上午" in t or "下午" in t or t.startswith("下")):
                continue

            raw_items.append({"text": t, "sender": sender, "x": x, "y": y, "conf": conf})

    # 合併相鄰的同 sender 文字成一條訊息（y 差距 < 40px 視為同一條）
    if not raw_items:
        return []

    messages = []
    current = {"text": raw_items[0]["text"], "sender": raw_items[0]["sender"],
               "y": raw_items[0]["y"], "min_conf": raw_items[0]["conf"]}

    for item in raw_items[1:]:
        same_sender = item["sender"] == current["sender"]
        close_y = abs(item["y"] - current["y"]) < 40
        if same_sender and close_y:
            current["text"] += " " + item["text"]
            current["min_conf"] = min(current["min_conf"], item["conf"])
        else:
            messages.append(current)
            current = {"text": item["text"], "sender": item["sender"],
                       "y": item["y"], "min_conf": item["conf"]}
    messages.append(current)

    return messages


def detect_new_messages(previous_msgs, current_msgs):
    """
    比對新舊訊息列表，回傳新增的對方訊息。
    用模糊比對（相似度 > 70%）避免 OCR 微小差異導致誤判。
    """
    from difflib import SequenceMatcher

    prev_texts = [m["text"] for m in previous_msgs]
    new_them = []

    for msg in current_msgs:
        if msg["sender"] != "them":
            continue
        # 檢查這條訊息是不是已經存在於上一次的列表
        is_old = False
        for pt in prev_texts:
            ratio = SequenceMatcher(None, msg["text"], pt).ratio()
            if ratio > 0.7:
                is_old = True
                break
        if not is_old:
            new_them.append(msg["text"])

    return new_them


def detect_search_needs(new_messages_text):
    """
    關鍵字偵測：新訊息需不需要搜尋事實資料。
    回傳: (needs_search: bool, search_queries: list)
    """
    dc = get_date_context()
    all_text = " ".join(new_messages_text)

    sport_kw = ["NBA", "nba", "NFL", "MLB", "NHL", "足球", "籃球", "棒球",
                "比賽", "比分", "球隊", "賽", "冠軍", "季後賽", "playoffs",
                "Celtics", "Thunder", "Lakers", "Warriors", "馬刺", "勇士", "湖人"]
    stock_kw = ["台積電", "鴻海", "聯發科", "股票", "股價", "TSMC", "2330",
                "大盤", "加權指數", "美股", "股市"]
    news_kw = ["新聞", "最新", "發生了什麼", "頭條"]
    weather_kw = ["天氣", "氣溫", "下雨", "颱風"]

    queries = []

    has_sport = any(kw in all_text for kw in sport_kw)
    has_stock = any(kw in all_text for kw in stock_kw)
    has_news = any(kw in all_text for kw in news_kw)
    has_weather = any(kw in all_text for kw in weather_kw)

    if has_sport:
        q = resolve_relative_dates(all_text[:60] + " NBA賽程")
        queries.append(q)
    if has_stock:
        for kw in ["台積電", "鴻海", "聯發科"]:
            if kw in all_text:
                queries.append(f"{kw}股價")
                break
        else:
            queries.append(resolve_relative_dates(all_text[:40] + " 股價"))
    if has_news:
        queries.append(resolve_relative_dates(all_text[:40] + f" {dc['today']}"))
    if has_weather:
        queries.append("台灣天氣")

    needs_search = len(queries) > 0
    return needs_search, queries


def vision_fallback_analyze(chat_img):
    """
    Vision 保險：OCR 失敗或信心度太低時，用 Vision 分析。
    只在 OCR 搞不定時才呼叫，不是每次都用。
    """
    tmp = os.path.join(TMPDIR, "tg_analyze.png")
    chat_img.save(tmp, quality=95)
    with open(tmp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    dc = get_date_context()
    r = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=300,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
            {"type": "text", "text": f"""{dc['prompt']}

這是 Telegram 完整對話截圖。
綠色氣泡（靠右）= 我發的
白色氣泡（靠左）= 對方發的

請按以下步驟分析：

步驟一：找到畫面中【我的最後一條綠色氣泡】的位置。

步驟二：看那條綠色氣泡【之後】有沒有白色氣泡。
- 如果最底部是綠色氣泡（我的）→ sender = "me"
- 如果綠色氣泡之後有白色氣泡（對方的）→ sender = "them"

步驟三：如果 sender = "them"，把對方在我最後一條綠色氣泡之後發的【所有白色氣泡的內容】整理出來（可能有多條）。

步驟四：判斷對方的這些新訊息裡有沒有需要查證的事實性問題（比賽比分、股票價格、新聞事件、數據排名等）。如果有，列出具體要搜尋的關鍵字。
【重要】search_queries 裡面不能寫「明天」「今天」「昨天」，必須轉成具體日期（{dc['today']}、{dc['tomorrow']}、{dc['yesterday']}）。年份是 {dc['year']} 不是 2025。

只回 JSON，不要其他文字：
{{"sender":"them或me","new_messages":"對方新訊息的完整內容（多條合併）","needs_search":true或false,"search_queries":["要搜的關鍵字1","要搜的關鍵字2"]}}"""}
        ]}]
    )
    resp = r.content[0].text.strip()
    if resp.startswith("```"):
        resp = re.sub(r"^```(?:json)?\s*", "", resp)
        resp = re.sub(r"\s*```$", "", resp)
    try:
        match = re.search(r"\{.*\}", resp, re.DOTALL)
        if match:
            return json.loads(match.group())
    except Exception:
        pass
    return {"sender": "unknown", "new_messages": "", "needs_search": False, "search_queries": []}


# ============================================================
# 維修 3: 搜尋函數 — 根據 search_queries 分別用對的工具
# ============================================================
def search_for_queries(search_queries):
    """根據每個 query 的類型用對應的工具搜尋，合併結果。日期自動修正。"""
    if not search_queries:
        return ""
    try:
        sys.path.insert(0, str(Path("C:/Users/blue_/claude-telegram-bot")))
        from bot import fetch_sports_scores, fetch_stock
        from tavily import TavilyClient
        tavily_key = os.environ.get("TAVILY_API_KEY", "")

        # 程式碼端日期轉換（第一層防護）
        search_queries = [resolve_relative_dates(q) for q in search_queries]

        sport_keywords = ["NBA", "nba", "NFL", "MLB", "NHL", "足球", "籃球", "棒球",
                          "比賽", "比分", "球隊", "賽", "冠軍", "季後賽", "playoffs",
                          "Celtics", "Thunder", "Lakers", "Warriors"]
        stock_keywords = ["台積電", "鴻海", "聯發科", "股票", "股價", "TSMC", "2330",
                          "大盤", "加權指數", "美股", "股市"]

        results = []
        sports_done = False

        for query in search_queries[:3]:  # 最多搜 3 個
            is_sport = any(kw in query for kw in sport_keywords)
            is_stock = any(kw in query for kw in stock_keywords)

            # Tavily 帶日期篩選的搜尋函數
            def _tavily_dated(q, max_r=2):
                if not tavily_key:
                    from bot import tavily_search
                    return tavily_search(q, max_r, "basic")
                tc = TavilyClient(api_key=tavily_key)
                dc = get_date_context()
                resp = tc.search(
                    query=q, max_results=max_r, search_depth="basic",
                    include_answer=True, time_range="week",
                )
                lines = []
                ans = resp.get("answer", "")
                if ans:
                    lines.append(f"AI 整合答案：{ans}")
                for r in resp.get("results", [])[:max_r]:
                    lines.append(f"{r.get('title','')}\n{r.get('content','')[:150]}")
                return "\n".join(lines) if lines else ""

            if is_sport and not sports_done:
                scores = fetch_sports_scores("nba")
                results.append(f"【ESPN 即時比分】\n{scores}")
                sports_done = True
                tr = _tavily_dated(query, 2)
                if tr:
                    results.append(f"【{query} 搜尋】\n{tr}")

            elif is_stock:
                if "台積電" in query or "TSMC" in query or "2330" in query:
                    results.append(f"【台積電台股報價】\n{fetch_stock('2330.TW')}")
                elif "鴻海" in query:
                    results.append(f"【鴻海台股報價】\n{fetch_stock('2317.TW')}")
                elif "聯發科" in query:
                    results.append(f"【聯發科台股報價】\n{fetch_stock('2454.TW')}")
                else:
                    tr = _tavily_dated(query, 2)
                    if tr:
                        results.append(f"【{query} 搜尋】\n{tr}")
            else:
                tr = _tavily_dated(query, 2)
                if tr:
                    results.append(f"【{query} 搜尋】\n{tr}")

        return "\n\n".join(results) if results else ""
    except Exception as e:
        return f"搜尋失敗：{e}"


# ============================================================
# 文字 API 回覆（不傳圖片，比 Vision 快 2-3 倍）
# ============================================================
def generate_text_reply(conversation_history, new_messages_text, search_context=""):
    """
    Claude 文字 API 回覆。用對話歷史文字 + 新訊息 + 搜尋結果。
    不傳圖片 → 快很多。
    conversation_history: [{"text": "...", "sender": "them/me"}]
    new_messages_text: ["新訊息1", "新訊息2"]
    """
    dc = get_date_context()
    system_prompt = PERSONA + f"\n\n{dc['prompt']}"

    if search_context:
        system_prompt += (
            "\n\n====== 搜尋到的最新事實資料 ======\n"
            + search_context[:2000]
            + "\n====== 事實資料結束 ======\n\n"
            "【絕對規則】\n"
            "1. 搜尋結果裡有的比分、數字、價格 → 必須用搜尋結果的，一個字都不能改\n"
            "2. 搜尋結果沒提到的數字 → 絕對不能自己編，說「這個我不確定」\n"
            "3. 不能把 A 隊的比分說成 B 隊的\n"
            "4. 寧可回答「我不知道」也不能編造任何數字\n"
            "5. 不要引用你自己之前說過的數字當事實，每次都用搜尋結果\n"
        )

    # 組對話歷史（最近 15 條）
    history_lines = []
    for msg in conversation_history[-15:]:
        label = "[對方]" if msg["sender"] == "them" else "[我]"
        history_lines.append(f"{label} {msg['text']}")
    history_text = "\n".join(history_lines)

    new_text = "\n".join(f"• {m}" for m in new_messages_text)

    user_content = f"""以下是 Telegram 對話紀錄（最近的訊息）：

{history_text}

對方剛發了新訊息：
{new_text}

請回覆對方的新訊息。
- 根據對話上下文回覆，不要重複之前說過的
- 對方問問題就回答，嗆人就幽默化解，聊天就接話
- 如果對方連發多條，整體理解後一次回覆
- 如果涉及事實（比賽/股價/新聞），只能用搜尋結果的數字
- 搜尋結果沒有的就說不確定

只回覆要發送的文字，不要加任何格式或解釋。"""

    r = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=200, system=system_prompt,
        messages=[{"role": "user", "content": user_content}]
    )
    return r.content[0].text.strip()


def send_reply(msg, regions):
    """在輸入框打字並送出"""
    ix, iy = regions["input_box"]["center"]
    pyautogui.click(ix, iy)
    time.sleep(0.3)
    pyperclip.copy(msg)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(1)


def monitor_and_reply(regions, stop_time, monitor=None):
    """
    OCR 文字追蹤監控。
    1. 截圖 → OCR 提取文字 + 像素顏色判斷 sender
    2. 跟上次比對 → 找新增的對方訊息
    3. 等對方說完（2+2 秒）
    4. 關鍵字搜尋（本地判斷）
    5. Claude 文字 API 回覆（不傳圖片）
    6. OCR 失敗時 → Vision 保險
    """
    WAIT_COMPLETE = 2
    WAIT_COMPLETE_ROUNDS = 2

    # OCR 追蹤：儲存上次的訊息列表
    chat = grab_chat(regions, monitor)
    previous_messages = ocr_extract_messages(chat)
    last_hash = chat_hash(chat)
    cooldown_until = 0

    # 對話歷史（給 Claude 文字 API 用）
    conversation_history = list(previous_messages)  # 初始化

    print(f"[Monitor] 開始監控 → {stop_time}", flush=True)
    print(f"[Monitor] 模式：OCR 文字追蹤 + Vision 保險", flush=True)
    print(f"[Monitor] 冷卻：{COOLDOWN_SECONDS}s / 輪詢：{POLL_INTERVAL}s / 等說完：{WAIT_COMPLETE}sx{WAIT_COMPLETE_ROUNDS}", flush=True)
    print(f"[Monitor] 初始訊息數：{len(previous_messages)}", flush=True)
    print(f"[Monitor] 停止方式：touch {STOP_FILE}", flush=True)

    while datetime.now().strftime("%H:%M") < stop_time:
        if should_stop():
            break
        # 輪詢等待（每秒檢查停止信號，不再一次 sleep 6 秒）
        for _ in range(POLL_INTERVAL):
            if should_stop():
                break
            time.sleep(1)
        if should_stop():
            break
        try:
            if time.time() < cooldown_until:
                continue

            chat = grab_chat(regions, monitor)
            h = chat_hash(chat)

            if h == last_hash:
                continue

            t = datetime.now().strftime("%H:%M:%S")

            # === 等對方說完（最多 2 輪 x 2 秒）===
            latest_chat = chat
            latest_hash = h
            for wait_round in range(WAIT_COMPLETE_ROUNDS):
                time.sleep(WAIT_COMPLETE)
                check = grab_chat(regions, monitor)
                check_h = chat_hash(check)
                if check_h != latest_hash:
                    print(f"[{t}] 對方還在發（第{wait_round+1}輪），等...", flush=True)
                    latest_chat = check
                    latest_hash = check_h
                else:
                    break

            # === OCR 提取文字 + 像素顏色判斷 ===
            current_messages = ocr_extract_messages(latest_chat)
            # PaddleOCR rec_scores 是 0-1（不是 0-100），轉成百分比比較
            avg_conf = sum(m.get("min_conf", 0.5) for m in current_messages) / max(len(current_messages), 1)
            if avg_conf < 1:
                avg_conf = avg_conf * 100  # 0.95 → 95

            # === 判斷用 OCR 還是 Vision ===
            use_vision = False
            if len(current_messages) == 0:
                print(f"[{t}] OCR 提取不到文字，啟用 Vision 保險", flush=True)
                use_vision = True
            elif avg_conf < 45:
                print(f"[{t}] OCR 信心度太低（{avg_conf:.0f}），啟用 Vision 保險", flush=True)
                use_vision = True

            if use_vision:
                # === Vision 保險 ===
                analysis = vision_fallback_analyze(latest_chat)
                sender = analysis.get("sender", "unknown")
                new_msgs_text = [analysis.get("new_messages", "")] if analysis.get("new_messages") else []
                needs_search = analysis.get("needs_search", False)
                search_queries = analysis.get("search_queries", [])
                print(f"[{t}] Vision: sender={sender} new={new_msgs_text[:1]} search={needs_search}", flush=True)

                if sender != "them" or not new_msgs_text:
                    last_hash = latest_hash
                    previous_messages = current_messages if current_messages else previous_messages
                    if sender == "me":
                        print(f"[{t}] 自己的訊息，跳過", flush=True)
                    else:
                        print(f"[{t}] sender={sender}，跳過", flush=True)
                    continue
            else:
                # === OCR 主流程：比對找新訊息 ===
                new_them_msgs = detect_new_messages(previous_messages, current_messages)

                if not new_them_msgs:
                    # 沒有新的對方訊息（可能是自己的訊息或系統變化）
                    last_hash = latest_hash
                    previous_messages = current_messages
                    conversation_history = list(current_messages)
                    continue

                print(f"[{t}] OCR 偵測到 {len(new_them_msgs)} 條新訊息：{new_them_msgs}", flush=True)
                new_msgs_text = new_them_msgs

                # 關鍵字偵測搜尋需求（本地，0.01 秒）
                needs_search, search_queries = detect_search_needs(new_msgs_text)
                print(f"[{t}] 關鍵字偵測：needs_search={needs_search} queries={search_queries}", flush=True)

            # === 搜尋（如需要）===
            search_context = ""
            if needs_search and search_queries:
                print(f"[{t}] 搜尋：{search_queries}", flush=True)
                search_context = search_for_queries(search_queries)
                print(f"[{t}] 搜尋完成（{len(search_context)} chars）", flush=True)

            # === Claude 文字 API 回覆（不傳圖片）===
            reply = generate_text_reply(conversation_history, new_msgs_text, search_context)

            if reply and len(reply) > 1:
                print(f"[{t}] → 回覆：{reply}", flush=True)
                send_reply(reply, regions)

                cooldown_until = time.time() + COOLDOWN_SECONDS
                print(f"[{t}] 冷卻 {COOLDOWN_SECONDS} 秒", flush=True)

                # 冷卻期間每秒檢查停止信號（不再一次 sleep 20 秒）
                for _ in range(COOLDOWN_SECONDS):
                    if should_stop():
                        break
                    time.sleep(1)
                chat = grab_chat(regions, monitor)
                last_hash = chat_hash(chat)
                previous_messages = ocr_extract_messages(chat)
                conversation_history = list(previous_messages)
            else:
                print(f"[{t}] 回覆為空，跳過", flush=True)
                last_hash = latest_hash
                previous_messages = current_messages
                conversation_history = list(current_messages)

        except Exception as e:
            print(f"[ERR] {e}", flush=True)
            time.sleep(5)

    reason = "收到停止信號" if _should_stop else "到達結束時間"
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 監控結束（{reason}）", flush=True)


# ============================================================
# 主流程
# ============================================================
def main(contact_name, stop_time, monitor=None):
    import threading

    # 啟動前殺殘留進程（防 VRAM 雙倍佔用）
    _kill_stale_gpu_processes()

    # 啟動時清掉殘留的停止旗標（避免上次異常退出留下的旗標導致立即停止）
    if os.path.exists(STOP_FILE):
        try:
            os.remove(STOP_FILE)
        except OSError:
            pass

    print("=" * 50, flush=True)
    print(f"Telegram 自動回覆", flush=True)
    print(f"好友：{contact_name}", flush=True)
    print(f"監控到：{stop_time}", flush=True)
    print("=" * 50, flush=True)

    # 背景預載 PaddleOCR（跟 Step 1-4 並行，省 ~5 秒）
    def _preload_ocr():
        try:
            _get_ocr_engine()
            print("[Preload] PaddleOCR GPU 模型已載入", flush=True)
        except Exception as e:
            print(f"[Preload] PaddleOCR 載入失敗：{e}", flush=True)

    ocr_thread = threading.Thread(target=_preload_ocr, daemon=True)
    ocr_thread.start()

    # Step 1: 定位
    print("\n[Step 1] 定位 Telegram UI...", flush=True)
    regions = locate_all(monitor)
    monitor = _resolve_monitor(regions, monitor)  # auto-detect 後固定成 int 給後續函數用
    if should_stop():
        return False

    # Step 2: 搜尋好友
    print(f"\n[Step 2] 搜尋好友：{contact_name}", flush=True)
    search_contact(regions, contact_name, monitor)
    if should_stop():
        return False

    # Step 3: 點選好友
    print(f"\n[Step 3] 點選好友...", flush=True)
    click_contact(regions, contact_name, monitor)
    if should_stop():
        return False

    # Step 4: 確認好友名稱（用 Step 1 的 regions，不重新定位）
    print(f"\n[Step 4] 確認好友名稱...", flush=True)
    confirmed = verify_friend(contact_name, monitor, regions=regions)
    if not confirmed:
        print("[ERROR] 好友名稱不符，中止", flush=True)
        return False

    # 等 OCR 預載完成（每秒檢查停止信號）
    for _ in range(15):
        if should_stop():
            print("\n[STOP] 啟動階段收到停止信號，中止", flush=True)
            return False
        if not ocr_thread.is_alive():
            break
        time.sleep(1)

    # Step 5-7: 監控 + 回覆
    print(f"\n[Step 5] 開始監控對話...", flush=True)
    monitor_and_reply(regions, stop_time, monitor)

    print("\n✅ 完成", flush=True)
    return True


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        print("用法：python tg_auto_chat.py <好友名稱> <監控時間HH:MM>")
        print("範例：python tg_auto_chat.py 巴斯 23:30")
        sys.exit(0)

    contact = sys.argv[1]
    stop = sys.argv[2]
    main(contact, stop)
