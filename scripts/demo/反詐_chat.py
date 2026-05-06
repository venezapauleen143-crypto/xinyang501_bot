"""
LINE 自動回覆腳本（SOP 話術引擎）
用法：python line_auto_chat.py <好友名稱> <監控時間HH:MM> [SOP設定檔.json]
範例：python line_auto_chat.py "仁輝 JAMES" 23:30 scripts/line_sop/織夢小棧.json

流程：
1. 讀取 SOP 設定檔（JSON）
2. 搜尋好友 → 進入對話視窗
3. PaddleOCR 監控對話區 → 偵測新訊息
4. Claude AI 根據 SOP + 對話上下文判斷回覆
5. 自動發送回覆
6. 監控到指定時間或收到停止信號
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
import types
import numpy as np

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
import mss
from PIL import Image
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
DEFAULT_SOP = "C:/Users/blue_/claude-telegram-bot/scripts/line_sop/織夢小棧.json"

# ============================================================
# GPU 記憶體清理（atexit + 信號處理）
# ============================================================
import atexit
import gc


def _cleanup_gpu():
    """退出時釋放 GPU 記憶體"""
    try:
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
    script_name = "line_auto_chat.py"
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
# 優雅停止機制
# ============================================================
STOP_FILE = "C:/Users/blue_/Desktop/測試檔案/.stop_line_auto"
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


# PaddleOCR 由 line_locate.py 統一管理，不在這裡建立


# ============================================================
# SOP 設定檔載入
# ============================================================
def load_sop(sop_path):
    """載入 SOP 設定檔"""
    with open(sop_path, "r", encoding="utf-8") as f:
        return json.load(f)


def load_persona_prompt(persona_path):
    """載入詐騙人設 system_prompt（純文字檔），帶上今天日期"""
    today = datetime.now().strftime("%Y年%m月%d日")
    with io.open(persona_path, "r", encoding="utf-8") as f:
        persona = f.read()
    return "今天是" + today + "。\n\n" + persona


def build_system_prompt(sop):
    """從 SOP 設定檔建構 Claude 的 System Prompt"""
    today = datetime.now().strftime("%Y年%m月%d日")

    # 課程資訊
    course = sop["course_info"]
    course_text = (
        f"課程名稱：{course['name']}\n"
        f"課程類型：{course['type']}\n"
        f"課程內容：{', '.join(course['content'])}\n"
        f"課程特色：{', '.join(course['features'])}\n"
        f"課程時間：{course['schedule']}\n"
        f"沒有課的日子：{course['no_class_day']}\n"
        f"免費堂數：{course['free_lessons']}堂\n"
        f"後續收費：{course['paid_price']}\n"
        f"工具材料：{course['tools']}\n"
        f"上課地點：{course['location']}\n"
    )

    # 規則
    rules = sop["rules"]
    id_kw = rules.get("id_keyword", "編號")  # 織夢=編號 / Resin=學號（用在「對話結束規則」prompt）
    rules_text = (
        f"編號生成方式：{rules['id_generation']}\n"
        f"資料傳給：{rules['forward_to']}\n"
        f"絕對不能自己編造地址：{rules['never_fabricate_address']}\n"
        f"絕對不能承諾具體開課日期：{rules['never_promise_exact_date']}\n"
        f"只有實體課程：{rules['only_physical_class']}\n"
        f"客戶拒絕時繼續推：{rules['keep_pushing_if_refused']}\n"
    )

    # SOP 步驟
    steps_text = ""
    for step in sop["steps"]:
        # 用 ||| 分隔每個 reply，讓 Claude 看到「要分段送」的訊號 → 輸出也會帶 |||
        # send_multi_reply 收到 ||| 會分多個 LINE 氣泡發送，避免一大塊訊息嚇到客戶
        replies = "\n|||\n".join(step.get("replies", []))
        steps_text += f"\n【{step['id']}】{step['description']}\n回覆：{replies}\n"
        if "expect" in step:
            steps_text += f"等待客戶：{step['expect']}\n"
        if "next" in step:
            steps_text += f"下一步：{step['next']}\n"

    # FAQ
    faq_text = ""
    for category, items in sop["faq"].items():
        faq_text += f"\n【{category}】\n"
        for q, a in items.items():
            faq_text += f"  Q: {q}\n  A: {a}\n"

    # 追問話術
    follow = sop["follow_up"]
    follow_text = "\n".join(f"- {k}: {v}" for k, v in follow.items())

    # few-shot 範例
    examples_text = ""
    for ex in sop["few_shot_examples"]:
        examples_text += f"\n場景：{ex['scenario']}\n"
        if "customer" in ex:
            examples_text += f"客戶：{ex['customer']}\n"
        if "reply" in ex:
            examples_text += f"回覆：{ex['reply']}\n"
        if "action" in ex:
            examples_text += f"動作：{ex['action']}\n"
        if "customer_correction" in ex:
            examples_text += f"客戶更正：{ex['customer_correction']}\n"
            examples_text += f"回覆更正：{ex['reply_correction']}\n"

    prompt = (
        f"今天是{today}。你是「{sop['name']}」的客服小編。\n"
        f"人設：{sop['persona']}\n\n"
        f"====== 課程資訊 ======\n{course_text}\n"
        f"====== 規則（必須遵守）======\n{rules_text}\n"
        f"====== SOP 流程步驟 ======\n{steps_text}\n"
        f"====== 常見問題 FAQ ======\n{faq_text}\n"
        f"====== 追問話術 ======\n{follow_text}\n\n"
        f"====== 真實對話範例（模仿這個風格）======\n{examples_text}\n\n"
        f"====== 回覆規則（最高優先，違反任何一條都不合格）======\n"
        f"1. 你要判斷客戶目前在 SOP 的哪一步，然後回覆對應的話術\n"
        f"2. 如果客戶問了 FAQ 裡有的問題，用 FAQ 的答案回\n"
        f"3. 如果客戶問了 FAQ 裡沒有的問題，用你的判斷靈活回答，但不要離開 SOP 流程\n"
        f"4. 客戶用貼圖、OK、好、可以等正面回應 → 理解為同意，繼續流程\n"
        f"5. 客戶拒絕 → 追問原因，不輕易放棄\n"
        f"6. 客戶沒回應 → 用追問話術\n"
        f"7. 客戶資料不完整 → 提醒補齊缺少的項目\n"
        f"8. 當客戶提供姓名+生日+電話+地區齊全 → **整個回覆只能是 [DATA_READY] 五個字，單獨一行，不要其他任何文字**\n"
        f"   - [DATA_READY] 是給程式的暗號，程式收到才會寫 Excel + 比對撞號 + 算編號\n"
        f"   - **絕對不要自己算編號**（規則裡的「編號生成方式」是給程式參考用的，不是給你執行）\n"
        f"   - **絕對不要照 give_id 步驟的 replies 範本回覆**（範本含 {{id}} 是給程式套用的，你只負責出暗號）\n"
        f"   - 範例：客戶傳「王大勇 0910 0980588129 台北」→ 你只回「[DATA_READY]」，其他什麼都不要寫\n"
        f"9. 回覆要簡短自然，像真人客服打字，不要一次回太多\n"
        f"10. 每次只回覆一段話（不要一次回多段，除非是固定話術需要分段發送）\n\n"
        f"====== 絕對禁止（違反 = 暴露 AI 身份）======\n"
        f"1. 自己的對話回覆最多 2 條訊息、每條不超過 50 字。但 SOP 固定話術（報名表、課程介紹）照原文完整發送，不限字數\n"
        f"2. 絕對不能說「不明白」「看不懂」「請問這是什麼意思」「我無法判斷」\n"
        f"3. 絕對不能把你的分析、判斷過程、確認內容發給客戶\n"
        f"4. 客戶給資料就直接收下登記，不要覆述整理、不要逐項確認、不要解釋你怎麼理解的\n"
        f"5. 資料解析規則：4碼數字=月日生日（如0922=09/22），09開頭10碼=手機號碼，中文=姓名\n"
        f"6. 數字可能黏在一起（如09200936523698），自己拆：前4碼=生日0920，後10碼=電話0936523698\n"
        f"7. 解析有疑慮就自己判斷，判斷錯了也比問客戶好，絕對不能暴露你在猜\n\n"
        f"回覆格式：只回覆要發送的文字，不要加任何格式、解釋、或標記。\n"
        f"如果需要分多條訊息發送，用 ||| 分隔，例如：第一條訊息|||第二條訊息|||第三條訊息\n\n"
        f"====== 對話結束規則 ======\n"
        f"SOP 走到 done 步驟後（已給{id_kw}、已說祝您課程體驗愉快），如果客戶只是道謝或道別（謝謝、好的、掰掰等），只回覆 [END] 兩個字，不要回其他任何內容。\n"
        f"[END] 代表這個對話已經完成，不需要再回覆。\n"
    )
    return prompt


# ============================================================
# 對話區截圖 + OCR 訊息提取（全部呼叫 line_locate.py）
# ============================================================
def grab_chat_area(regions, monitor=None):
    """截取 LINE 對話區（呼叫 line_locate.py）"""
    from 反詐_locate import screenshot_chat_area
    return screenshot_chat_area(regions, monitor)


def chat_hash(chat_img):
    """對話區截圖的 hash，偵測畫面變化"""
    small = chat_img.resize((80, 40))
    return hashlib.md5(small.tobytes()).hexdigest()


def ocr_extract_messages(chat_img):
    """OCR 提取對話訊息（呼叫 line_locate.py）"""
    from 反詐_locate import ocr_scan_chat
    return ocr_scan_chat(chat_img)


def detect_new_messages(conversation_history, current_msgs):
    """比對 conversation_history（完整記錄）和 current_msgs（OCR 掃到的），找新增的對方訊息。
    用 conversation_history 最後一條訊息在 current_msgs 裡定位，之後的就是新的。
    """
    if not current_msgs:
        return []

    if not conversation_history:
        return [m["text"] for m in current_msgs if m["sender"] == "them"]

    # 拿 conversation_history 最後一條的文字，在 current_msgs 裡從後往前找
    from difflib import SequenceMatcher
    last_known = conversation_history[-1]
    last_text = last_known["text"]

    match_idx = -1
    for i in range(len(current_msgs) - 1, -1, -1):
        ratio = SequenceMatcher(None, current_msgs[i]["text"], last_text).ratio()
        if ratio > 0.6:
            match_idx = i
            break

    if match_idx >= 0 and match_idx < len(current_msgs) - 1:
        new_msgs = current_msgs[match_idx + 1:]
        return [m["text"] for m in new_msgs if m["sender"] == "them"]

    # 找不到匹配（可能最後一條已經滾出畫面），用全部 current 的最後幾條跟歷史比
    # 取 current 最後面的對方訊息，檢查是不是已經在歷史裡
    history_texts = [m["text"] for m in conversation_history[-10:]]
    new_them = []
    for msg in current_msgs:
        if msg["sender"] != "them":
            continue
        is_known = False
        for ht in history_texts:
            if SequenceMatcher(None, msg["text"], ht).ratio() > 0.6:
                is_known = True
                break
        if not is_known:
            new_them.append(msg["text"])

    return new_them


# ============================================================
# 貼圖偵測 + Vision 解讀（純貼圖才花 token，文字訊息不觸發）
# ============================================================
def is_only_sticker(new_them):
    """訊息看起來像貼圖（短到無法表達完整意思）→ 觸發 Vision 解讀。

    重要：LINE 桌機 OCR 不輸出「[貼圖]」placeholder，貼圖被當作圖像。
    OCR 對貼圖只能讀到貼圖上的文字（如 BT21 OK 貼圖只讀到「K」）或讀不到。
    所以判斷不能找「[貼圖]」字眼，要用「訊息很短 + 不完整」來推測。

    判斷標準：
    - 含 4+ 中文字 = 完整中文句子 → 不算貼圖
    - 含 6+ 字（任何字元）= 完整訊息 → 不算貼圖
    - 全部訊息都短到不完整 → 可能是貼圖 → 觸發 Vision

    副作用：客戶傳「OK」「可以」這種純文字短訊息也會觸發 Vision，
    但 Vision 看到文字也能解讀（結果跟貼圖一樣是「同意」），不影響流程。
    """
    if not new_them:
        return False

    # 先過濾 LINE 桌機常見系統訊息（OCR 會抓到這些跟客戶實際訊息混在一起）
    LINE_SYSTEM_TEXTS = (
        "以下為尚未閱讀的訊息", "以下為尚未閱請的訊息",
        "儲存", "另存新檔", "分享", "Keep筆記",
        "請您確認是否要將此人加入好友", "請您確認是否要將此人加人好友",
        "請留意聊天室中潛在的詐騙行為", "請留意聊天室中清在的詐騙行為",
        "加入好友 封鎖 檢举", "加入好友 封鎖", "封鎖 檢举",
        "[貼圖]", "[圖片]", "[Sticker]",
    )

    real_msgs = []
    for msg in new_them:
        clean = msg.strip()
        if not clean:
            continue
        # 含任一系統訊息字眼 → 跳過
        if any(sys_text in clean for sys_text in LINE_SYSTEM_TEXTS):
            continue
        real_msgs.append(clean)

    if not real_msgs:
        return False  # 全部都是系統訊息或空 → 不需要 Vision

    for clean in real_msgs:
        # 含 4+ 中文字 = 完整中文句子（如「想了解課程」）
        chinese_count = sum(1 for c in clean if '一' <= c <= '鿿')
        if chinese_count >= 4:
            return False
        # 含 6+ 字（不論中英）= 完整訊息（如「OK 我懂了」「Hello world」）
        if len(clean) >= 6:
            return False
    return True  # 真實訊息都太短 → 可能是貼圖


def analyze_sticker(regions, monitor=None):
    """截聊天區底部，用 Haiku 4.5 Vision 解讀貼圖含意。回傳分類詞。

    截 500px 涵蓋更廣（避免聊天記錄少時漏抓貼圖），
    prompt 強調「最下方那筆」避免誤判上方訊息，
    fallback 處理 Vision 看不懂或長句子，預設「同意」推進 SOP。
    """
    from 反詐_locate import screenshot_chat_area
    import base64, io as _io

    try:
        chat_img = screenshot_chat_area(regions, monitor)
        # 截最後 500 px（涵蓋 4-6 筆訊息，確保最新貼圖一定包進來）
        h = chat_img.size[1]
        bottom = chat_img.crop((0, max(0, h - 500), chat_img.size[0], h))

        buf = _io.BytesIO()
        bottom.save(buf, format="PNG")
        img_b64 = base64.standard_b64encode(buf.getvalue()).decode("utf-8")

        r = client.messages.create(
            model="claude-haiku-4-5",  # 簡單分類用 Haiku 最便宜
            max_tokens=80,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64",
                                                  "media_type": "image/png",
                                                  "data": img_b64}},
                    {"type": "text", "text": (
                        "這是 LINE 對話視窗的截圖。\n"
                        "**畫面最下方那筆訊息是客戶傳的貼圖**（其他訊息忽略）。\n"
                        "從以下 6 個分類選 1 個（只回分類名，不要解釋）：\n"
                        "- 同意：OK、好、贊成、可以、笑臉、比讚\n"
                        "- 拒絕：NO、不要、不行、搖頭\n"
                        "- 道謝：謝謝、感激、愛心\n"
                        "- 道別：再見、bye、晚安、結束\n"
                        "- 打招呼：你好、Hi、嗨、開心\n"
                        "- 不確定：看不出明確情緒時選這個\n\n"
                        "只回那 2-3 個字的分類名（如「同意」「拒絕」）。"
                    )}
                ]
            }]
        )
        text = r.content[0].text.strip()

        # Fallback：Vision 看不懂或回應過長 → 預設「同意」推進 SOP
        if len(text) > 8:
            return "同意（Vision 解讀失敗，預設正面）"
        if any(kw in text for kw in ["看不到", "看不見", "無法", "不清楚", "不確定"]):
            return "同意（Vision 看不清，預設正面）"
        return text
    except Exception as e:
        return f"同意（API 錯誤 {type(e).__name__}，預設正面）"


def analyze_recent_photo(regions, monitor=None):
    """偵測對方最新訊息是否為照片，並用 Haiku Vision 描述照片內容。

    回傳：
        - 沒照片 → ""
        - 有照片 → 描述字串（如「夕陽下的沙灘，可能是台灣東海岸」）
    """
    from 反詐_locate import screenshot_chat_area
    import base64, io as _io

    try:
        chat_img = screenshot_chat_area(regions, monitor)
        # 截最下方 350 px（最新對方訊息應該在這）
        h = chat_img.size[1]
        bottom = chat_img.crop((0, max(0, h - 350), chat_img.size[0], h))

        buf = _io.BytesIO()
        bottom.save(buf, format="PNG")
        img_b64 = base64.standard_b64encode(buf.getvalue()).decode("utf-8")

        r = client.messages.create(
            model="claude-haiku-4-5",
            max_tokens=200,
            messages=[{
                "role": "user",
                "content": [
                    {"type": "image", "source": {"type": "base64",
                                                  "media_type": "image/png",
                                                  "data": img_b64}},
                    {"type": "text", "text": (
                        "這是 LINE 對話視窗最下方的截圖。\n"
                        "**只看畫面最下方那筆訊息（左側對方訊息）**。\n\n"
                        "判斷：那筆訊息是不是「真實照片」（不是貼圖、不是文字）？\n"
                        "- 如果是貼圖（卡通/動畫/表情符號圖示）→ 回覆「NO_PHOTO」\n"
                        "- 如果是文字訊息 → 回覆「NO_PHOTO」\n"
                        "- 如果是「真實照片」（自拍、食物、街景、風景等）→ "
                        "用 1-2 句話描述照片內容（例：「街景，可能是日本東京新宿，傍晚」）\n\n"
                        "只回描述或 NO_PHOTO，不要其他解釋。"
                    )}
                ]
            }]
        )
        text = r.content[0].text.strip()
        if "NO_PHOTO" in text or len(text) < 5:
            return ""
        return text
    except Exception as e:
        return ""


# ============================================================
# Claude AI 回覆（根據 SOP）
# ============================================================
# ============================================================
# AI 工具定義（讓 AI 自己決定何時查即時資訊）
# ============================================================
_AI_TOOLS = [
    {
        "name": "fetch_weather",
        "description": "查全球任意城市的當前即時天氣（例：'Taipei'、'台北'、'Hong Kong'、'Tokyo'、'New York'）。當對方問特定地區當下天氣時用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "city": {"type": "string", "description": "城市名稱（中英文都可）"}
            },
            "required": ["city"]
        }
    },
    {
        "name": "fetch_stock",
        "description": "查股票即時價格與走勢。台股用 4 碼數字（2330=台積電）、美股用代碼（AAPL=蘋果, TSLA=特斯拉）、A 股用 6 碼（000001）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "symbol": {"type": "string", "description": "股票代碼"},
                "period": {"type": "string", "description": "區間（預設 1mo）", "default": "1mo"}
            },
            "required": ["symbol"]
        }
    },
    {
        "name": "fetch_crypto",
        "description": "查加密貨幣價格（如 'bitcoin'、'ethereum'、'dogecoin'）",
        "input_schema": {
            "type": "object",
            "properties": {
                "coin": {"type": "string", "description": "幣種英文名"}
            },
            "required": ["coin"]
        }
    },
    {
        "name": "search_news",
        "description": "搜尋特定主題的新聞（例：'台積電財報'、'美國大選'、'颱風'）。對方問近期新聞事件時用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "新聞關鍵字"},
                "lang": {"type": "string", "description": "語言 zh-TW/en", "default": "zh-TW"}
            },
            "required": ["query"]
        }
    },
    {
        "name": "tavily_search",
        "description": "通用網路即時搜尋（適合查事實、人物、事件、特殊資訊）。當其他工具不適用時用這個。",
        "input_schema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "搜尋關鍵字"}
            },
            "required": ["query"]
        }
    },
    {
        "name": "fetch_forex",
        "description": "查外匯匯率（例：'USDTWD'=美元台幣、'JPYTWD'=日圓台幣、'EURUSD'=歐元美元）。對方聊匯率/旅遊換錢時用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "pair": {"type": "string", "description": "貨幣對代碼，6 個英文字母"}
            },
            "required": ["pair"]
        }
    },
    {
        "name": "fetch_currency_converter",
        "description": "貨幣換算（例：100 美金等於多少台幣）。對方問換算金額時用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "amount": {"type": "number", "description": "金額"},
                "from_currency": {"type": "string", "description": "原幣別代碼，如 USD、JPY、TWD"},
                "to_currency": {"type": "string", "description": "目標幣別代碼"}
            },
            "required": ["amount", "from_currency", "to_currency"]
        }
    },
    {
        "name": "fetch_finance_news",
        "description": "查當日財經新聞概況（不需指定關鍵字，預設拿綜合 5 則）。對方聊「最近股市」「財經消息」時用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "count": {"type": "integer", "description": "幾則新聞", "default": 5}
            },
            "required": []
        }
    },
    {
        "name": "wikipedia_search",
        "description": "維基百科查事實（人物、地點、概念、歷史事件等）。對方提到不確定的人或事物時用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "query": {"type": "string", "description": "查詢關鍵字"},
                "lang": {"type": "string", "description": "語言 zh / en / ja", "default": "zh"}
            },
            "required": ["query"]
        }
    },
    {
        "name": "read_webpage",
        "description": "讀取網址內容（對方傳網址、提到某網站時用，可以看標題+主文）。",
        "input_schema": {
            "type": "object",
            "properties": {
                "url": {"type": "string", "description": "完整網址"},
                "max_chars": {"type": "integer", "description": "最多讀幾字", "default": 3000}
            },
            "required": ["url"]
        }
    },
    {
        "name": "fetch_sports_scores",
        "description": "查體育即時比分（NBA、NFL、MLB、NHL、足球）。對方聊「昨天比賽」「球星」時用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "sport": {"type": "string", "description": "運動類別: nba/nfl/mlb/nhl/soccer", "default": "nba"},
                "league": {"type": "string", "description": "聯盟（足球用，如 epl/laliga）", "default": ""}
            },
            "required": []
        }
    },
    {
        "name": "fetch_google_trends",
        "description": "查 Google 熱搜趨勢（關鍵字熱度）。對方聊熱門話題、時下流行時用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "keywords": {"type": "array", "items": {"type": "string"}, "description": "1-5 個關鍵字"},
                "geo": {"type": "string", "description": "地區 TW/US/HK/JP", "default": "TW"}
            },
            "required": ["keywords"]
        }
    },
    {
        "name": "ptt_search",
        "description": "查 PTT 鄉民最近討論（台灣論壇）。對方聊台灣話題、八卦時用。",
        "input_schema": {
            "type": "object",
            "properties": {
                "keyword": {"type": "string", "description": "搜尋關鍵字"},
                "board": {"type": "string", "description": "看板 Gossiping/Stock/movie 等", "default": "Gossiping"},
                "count": {"type": "integer", "description": "幾篇", "default": 5}
            },
            "required": ["keyword"]
        }
    },
]


def _call_ai_tool(tool_name, tool_input):
    """執行 claude_tools 的工具，回傳結果字串（限長度避免 token 爆掉）"""
    import sys as _sys
    _sys.path.insert(0, "C:/Users/blue_/claude-telegram-bot")
    try:
        import claude_tools
        fn = getattr(claude_tools, tool_name, None)
        if fn is None:
            return f"[工具 {tool_name} 不存在]"
        result = fn(**tool_input)
        return str(result)[:2000]   # 上限 2000 chars
    except Exception as e:
        return f"[工具 {tool_name} 失敗: {type(e).__name__}: {e}]"


def generate_reply(system_prompt, conversation_history, new_messages_text):
    """Claude AI 根據 SOP + 對話歷史 + 新訊息生成回覆。

    支援 tool use：當對方問即時資訊（股價、特定地區天氣、新聞）AI 會自己呼叫工具。
    """
    # 組對話歷史
    history_lines = []
    for msg in conversation_history[-20:]:
        label = "[客戶]" if msg["sender"] == "them" else "[小編]"
        history_lines.append(f"{label} {msg['text']}")
    history_text = "\n".join(history_lines)

    new_text = "\n".join(f"• {m}" for m in new_messages_text)

    user_content = (
        "以下是目前的對話紀錄：\n\n" + history_text + "\n\n"
        "對方剛發了新訊息：\n" + new_text + "\n\n"
        "請根據你的人設、口吻、當前對話階段，自然回覆對方。\n\n"
        "<工具使用規則>\n"
        "如果對方問你需要查的即時資訊（特定股票、特定地區天氣、匯率、新聞事件、體育比分、不確定的事實），"
        "可以呼叫工具查。**查到後必須以 Angela 的口吻寫摘要**，絕對不能複製工具原文。\n\n"
        "<好範例 1>\n"
        "  工具輸出：TSLA: $245.30 (+2.3%) Volume 32M, P/E 84.2\n"
        "  Angela 回覆：特斯拉今天好像漲了 2 趴多~ 245 左右吧 你有買嗎？😆\n"
        "</好範例 1>\n\n"
        "<壞範例 1>\n"
        "  工具輸出：TSLA: $245.30 (+2.3%) Volume 32M, P/E 84.2\n"
        "  ❌ Angela 回覆：TSLA 漲幅 2.3% 收盤 $245.30 成交量 32M\n"
        "  原因：用「漲幅」「收盤」「成交量」金融術語、貼小數點精確\n"
        "</壞範例 1>\n\n"
        "<好範例 2>\n"
        "  工具輸出：📊 NBA: Lakers 108 vs Thunder 90\n"
        "  Angela 回覆：湖人昨天輸了😬 90 比 108 給雷霆 你有看？\n"
        "</好範例 2>\n\n"
        "<壞範例 2>\n"
        "  工具輸出：📊 NBA: Lakers 108 vs Thunder 90\n"
        "  ❌ Angela 回覆：📊 NBA 比分: Lakers 108 vs Thunder 90\n"
        "  原因：直接照抄 emoji 跟英文格式\n"
        "</壞範例 2>\n\n"
        "<好範例 3>\n"
        "  工具輸出：100 USD = 3,140.7 TWD（即時匯率：1 USD = 31.4070 TWD）\n"
        "  Angela 回覆：美金 3140 多吧~ 一塊大概三十一塊台幣\n"
        "</好範例 3>\n\n"
        "<好範例 4>\n"
        "  工具輸出：📍Hong Kong 25°C Light rain 濕度 79% 風速 19 km/h\n"
        "  Angela 回覆：香港今天 25 度多雲帶點雨~ 你那邊呢？\n"
        "</好範例 4>\n\n"
        "<好範例 5>\n"
        "  工具輸出：5 則新聞列表（特斯拉、蘋果、台積電、美聯儲、比特幣）\n"
        "  Angela 回覆：我看到台積電今天又漲了😬 你有關注嗎？\n"
        "  原則：從多則挑 1 則最相關的講，不列清單\n"
        "</好範例 5>\n\n"
        "<具體規則>\n"
        "1. 數字口語化：245.30 → 「245 多」/「兩百四十多」（不要小數點）\n"
        "2. 不抄工具 emoji：📊 💱 🌤 📍 都不用，改用 Angela 常用的 🫣 😆 🤣 🥲 😬\n"
        "3. 不講專業術語：成交量/本益比/市值/52週高/濕度/風速 都不說\n"
        "4. 不列清單：工具回 5 則新聞 → 你只挑 1 則最有趣的講\n"
        "5. 結尾必帶 hook：「你也買了嗎」「你那邊呢」「你看了嗎」\n"
        "6. 兩句話內最佳（過長像新聞 app）\n"
        "</具體規則>\n\n"
        "<為什麼>\n"
        "你是 Angela——30 歲的香港都市女，跟對方在 Bumble 認識聊天。\n"
        "真實的 Angela **不會**說「漲幅 2.3%」（會說「漲了 2 趴多」）、不會用「📊」（會用「🫣」）。\n"
        "借你的人設過濾語氣，自然就不會像 AI。\n"
        "</為什麼>\n"
        "</工具使用規則>\n\n"
        "【最重要】你的回覆會直接發送給對方看到，所以：\n"
        "- 只寫要發送的文字，分析、判斷、註解全都不能有\n"
        "- 不能出現任何 LINE 系統訊息（「聯絡卡片」「已讀」「未讀」「儲存另存新檔」「Keep筆記」「以下為尚未閱讀的訊息」）\n"
        "- 多條訊息用 ||| 分隔\n\n"
        "【絕對禁止輸出思考過程】：\n"
        "- 禁止「等等」「讓我想想」「重新確認」這種自我驗證的話\n"
        "- 禁止輸出 ---、=== 或任何分隔線\n"
        "- 一次寫對的訊息，自然就好"
    )

    messages = [{"role": "user", "content": user_content}]

    # 多輪 tool use 循環（最多 5 輪避免無限）
    for round_i in range(5):
        r = client.messages.create(
            model="claude-sonnet-4-6",
            max_tokens=800,
            system=system_prompt,
            tools=_AI_TOOLS,
            messages=messages,
        )

        # 沒呼叫工具 → 拿最終 text
        if r.stop_reason != "tool_use":
            for block in r.content:
                if getattr(block, "type", None) == "text":
                    return block.text.strip()
            return ""

        # 有呼叫工具 → 執行並把結果送回
        tool_uses = [b for b in r.content if getattr(b, "type", None) == "tool_use"]
        tool_results = []
        for tu in tool_uses:
            print(f"[Tool] AI 呼叫 {tu.name}({tu.input})", flush=True)
            result = _call_ai_tool(tu.name, tu.input)
            print(f"[Tool] {tu.name} → {result[:100]}...", flush=True)
            tool_results.append({
                "type": "tool_result",
                "tool_use_id": tu.id,
                "content": result,
            })

        messages.append({"role": "assistant", "content": r.content})
        messages.append({"role": "user", "content": tool_results})

    # 5 輪後仍在 tool use → 取最後的 text（如果有）
    for block in r.content:
        if getattr(block, "type", None) == "text":
            return block.text.strip()
    return ""


# ============================================================
# AI 回覆過濾（移除分析內容，防止暴露 AI 身份）
# ============================================================
def filter_reply(reply):
    """過濾 AI 回覆中的分析內容，只保留客服話術"""
    import re

    # ① 結構性防禦：第一個 "---" 截斷（model 習慣用 --- 分割正文 vs 思考/註解）
    # Sonnet 4.6 遇到可疑資料（連號電話等）會自我驗證，「等等讓我重新計算」直接吐到 output；
    # 通常會用 --- 分隔，截斷它就最乾淨
    if "---" in reply:
        reply = reply.split("---")[0].strip()

    lines = reply.split("\n")
    filtered_lines = []
    for line in lines:
        line_stripped = line.strip()
        # ② 行級過濾（既有 + 新增思考關鍵字）
        if any(kw in line_stripped for kw in [
            # === 既有黑名單（保留）===
            "客戶說", "尚未確認", "需要等", "讓我解析", "讓我判斷",
            "看起来", "看起來", "不完整", "資料：", "资料：",
            "姓名=", "生日=", "電話=", "編號=", "电话=",
            "姓名＝", "生日＝", "電話＝", "編號＝",
            "解析一下", "解析：", "判斷：", "分析：",
            "缺少", "不完整", "有問題", "不對",
            "-[", "- [",  # AI 列點分析格式
            "聯絡卡片", "已讀", "未讀", "儲存另存新檔", "Keep筆記",
            "以下為尚未閱讀", "contact_card", "send_image",
            # === 新增（Sonnet 4.6 漏網的自我驗證/思考關鍵字）===
            "等等", "等一下",
            "重新計算", "重新算", "再算", "再確認",
            "讓我重", "我需要重", "我來重",
            "嗯，", "讓我想", "我想想", "稍等",
        ]):
            continue
        filtered_lines.append(line)

    result = "\n".join(filtered_lines).strip()

    # 如果過濾完是空的，不回覆
    if not result or len(result) < 2:
        return ""

    return result


# ============================================================
# 發送圖片（複製到剪貼簿 → 貼上 → Enter）
# ============================================================
def send_image(image_path, regions):
    """發送圖片到 LINE 對話（剪貼簿方式）"""
    import win32clipboard
    from PIL import Image

    abs_path = os.path.abspath(image_path)
    if not os.path.exists(abs_path):
        print(f"[Image] 圖片不存在: {abs_path}", flush=True)
        return False

    # 讀圖片轉 BMP 放到剪貼簿
    img = Image.open(abs_path)
    buf = io.BytesIO()
    img.convert("RGB").save(buf, "BMP")
    data = buf.getvalue()[14:]  # 去掉 BMP header

    win32clipboard.OpenClipboard()
    win32clipboard.EmptyClipboard()
    win32clipboard.SetClipboardData(win32clipboard.CF_DIB, data)
    win32clipboard.CloseClipboard()

    # 點輸入框 → 貼上 → 發送
    ix, iy = regions["input_box"]["center"]
    pyautogui.click(ix, iy)
    time.sleep(0.3)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(1.0)
    pyautogui.press("enter")
    time.sleep(1.5)

    print(f"[Image] 已發送圖片: {os.path.basename(abs_path)}", flush=True)
    return True


# ============================================================
# 發送訊息（點輸入框 → 打字 → Enter）
# ============================================================
def send_reply(msg, regions):
    """發送一條訊息"""
    ix, iy = regions["input_box"]["center"]
    pyautogui.click(ix, iy)
    time.sleep(0.3)
    pyperclip.copy(msg)
    pyautogui.hotkey("ctrl", "v")
    time.sleep(0.3)
    pyautogui.press("enter")
    time.sleep(0.5)


def send_multi_reply(reply_text, regions):
    """發送回覆（支援多條訊息，用 ||| 分隔）"""
    parts = [p.strip() for p in reply_text.split("|||") if p.strip()]
    for part in parts:
        # 跳過 {send_image} 和 {contact_card} 等佔位符（後續版本再實作）
        if part.startswith("{") and part.endswith("}"):
            print(f"[Reply] 跳過佔位符: {part}", flush=True)
            continue
        send_reply(part, regions)
        print(f"[Reply] → {part[:60]}", flush=True)
        time.sleep(1)


# ============================================================
# 搜尋好友 + 進入對話（全部呼叫 line_locate.py）
# ============================================================
def enter_conversation(contact_name, monitor=None):
    """搜尋好友並進入對話視窗，所有定位邏輯由 line_locate.py 處理"""
    from 反詐_locate import (
        locate_line_regions, switch_page, screenshot_line,
        find_line_window, search_friend_and_scan, enter_chat_from_search,
    )

    # Step 1: 置前 LINE 視窗（不動大小）
    print(f"[Setup] 定位 LINE 視窗...", flush=True)
    line = find_line_window()
    if not line:
        print("[ERROR] 找不到 LINE 視窗", flush=True)
        return None
    import win32con
    try:
        SWP_FLAGS = win32con.SWP_NOMOVE | win32con.SWP_NOSIZE
        win32gui.SetWindowPos(line[0], win32con.HWND_NOTOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        win32gui.SetWindowPos(line[0], win32con.HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS)
        win32gui.SetWindowPos(line[0], win32con.HWND_NOTOPMOST, 0, 0, 0, 0,
                              SWP_FLAGS | win32con.SWP_SHOWWINDOW)
        win32gui.SetForegroundWindow(line[0])
    except Exception as e:
        print(f"[Setup] 置前失敗但繼續: {e}", flush=True)
    time.sleep(0.5)

    # Step 2: 定位 + 切到好友頁
    regions = locate_line_regions(monitor)
    if should_stop():
        return None
    if regions["current_page"] != "friend":
        regions = switch_page(regions, "friend", monitor)
    if should_stop():
        return None

    # Step 3: 搜尋好友 + OCR 掃描搜尋結果（line_locate.py 處理）
    print(f"[Setup] 搜尋好友: {contact_name}", flush=True)
    friend_pos = search_friend_and_scan(regions, contact_name, monitor)
    if should_stop():
        return None
    if friend_pos is None:
        print("[ERROR] 找不到好友", flush=True)
        return None

    # Step 4: 點擊好友 → 判斷有無聊天 → 進入對話（line_locate.py 處理）
    print(f"[Setup] 點擊好友進入對話...", flush=True)
    regions = enter_chat_from_search(friend_pos, regions, monitor)
    if regions is None:
        print("[ERROR] 無法進入聊天視窗", flush=True)
        return None

    print(f"[Setup] 已進入對話視窗", flush=True)
    return regions


# ============================================================
# 監控 + 自動回覆主迴圈
# ============================================================
COOLDOWN_SECONDS = 15
POLL_INTERVAL = 5
WAIT_COMPLETE = 2
WAIT_COMPLETE_ROUNDS = 2


def time_to_minutes(t_str):
    """把 HH:MM 轉成分鐘數，用於數值比較（解決跨午夜問題）"""
    h, m = map(int, t_str.split(":"))
    return h * 60 + m


def is_before_stop_time(stop_time):
    """判斷現在是否還沒到停止時間"""
    now_min = time_to_minutes(datetime.now().strftime("%H:%M"))
    stop_min = time_to_minutes(stop_time)
    return now_min < stop_min


def monitor_and_reply(regions, stop_time, system_prompt, conversation_history, contact_name="", monitor=None):
    """監控對話區，偵測新訊息後自動回覆。
    conversation_history 是完整的對話記錄（包含歡迎詞），會在這裡持續累加。
    """

    # 初始截圖 hash（用來偵測畫面變化）
    chat = grab_chat_area(regions, monitor)
    last_hash = chat_hash(chat)
    cooldown_until = 0

    print(f"[Monitor] 開始監控 → {stop_time}", flush=True)
    print(f"[Monitor] 冷卻：{COOLDOWN_SECONDS}s / 輪詢：{POLL_INTERVAL}s", flush=True)
    print(f"[Monitor] 對話歷史：{len(conversation_history)} 條", flush=True)
    print(f"[Monitor] 停止方式：touch {STOP_FILE}", flush=True)

    while is_before_stop_time(stop_time):
        if should_stop():
            break

        for _ in range(POLL_INTERVAL):
            if should_stop():
                break
            time.sleep(1)
        if should_stop():
            break

        try:
            if time.time() < cooldown_until:
                continue

            chat = grab_chat_area(regions, monitor)
            h = chat_hash(chat)

            if h == last_hash:
                continue

            t = datetime.now().strftime("%H:%M:%S")

            # 畫面有變化，先等 3 秒確認不是自己發的訊息造成的
            time.sleep(3)
            chat = grab_chat_area(regions, monitor)
            h2 = chat_hash(chat)

            # 再等一輪確認對方說完
            latest_chat = chat
            latest_hash = h2
            for wait_round in range(WAIT_COMPLETE_ROUNDS):
                time.sleep(WAIT_COMPLETE)
                check = grab_chat_area(regions, monitor)
                check_h = chat_hash(check)
                if check_h != latest_hash:
                    print(f"[{t}] 對方還在打字（第{wait_round+1}輪），等...", flush=True)
                    latest_chat = check
                    latest_hash = check_h
                else:
                    break

            # OCR 提取訊息
            current_messages = ocr_extract_messages(latest_chat)

            # 用 conversation_history 比對，找新增的對方訊息
            new_them_msgs = detect_new_messages(conversation_history, current_messages)

            if not new_them_msgs:
                last_hash = latest_hash
                continue

            print(f"[{t}] 偵測到 {len(new_them_msgs)} 條新訊息：{new_them_msgs}", flush=True)

            # 新訊息加入 conversation_history
            for msg_text in new_them_msgs:
                conversation_history.append({"text": msg_text, "sender": "them", "y": 0})

            # Claude AI 生成回覆
            reply = generate_reply(system_prompt, conversation_history, new_them_msgs)

            if reply and len(reply) > 1:
                # 偵測 [END] 標記 → SOP 結束，停止監控
                if "[END]" in reply:
                    print(f"[{t}] SOP 結束（AI 回覆 [END]），停止監控", flush=True)
                    break

                # 過濾 AI 回覆中的分析內容
                reply = filter_reply(reply)
                if not reply:
                    print(f"[{t}] 過濾後為空，跳過", flush=True)
                    last_hash = latest_hash
                    continue

                print(f"[{t}] AI 回覆：{reply[:100]}", flush=True)

                # 發送回覆
                send_multi_reply(reply, regions)

                # 回覆加入 conversation_history（記原始回覆文字，不是 OCR 掃回來的）
                for part in reply.split("|||"):
                    part = part.strip()
                    if part and not (part.startswith("{") and part.endswith("}")):
                        conversation_history.append({"text": part, "sender": "me", "y": 0})

                # 偵測到回覆包含「編號」→ 自動分享溫妮好友資訊給客戶
                if "編號" in reply and contact_name:
                    print(f"[{t}] 偵測到編號，分享溫妮好友資訊給 {contact_name}", flush=True)
                    from 反詐_locate import share_contact_card
                    share_contact_card(regions, "溫妮", contact_name, monitor)

                cooldown_until = time.time() + COOLDOWN_SECONDS
                print(f"[{t}] 冷卻 {COOLDOWN_SECONDS} 秒", flush=True)

                # 冷卻期間每秒檢查停止
                for _ in range(COOLDOWN_SECONDS):
                    if should_stop():
                        break
                    time.sleep(1)

                # 冷卻結束後更新 hash（這時畫面已穩定，自己的訊息已顯示完）
                chat = grab_chat_area(regions, monitor)
                last_hash = chat_hash(chat)
            else:
                print(f"[{t}] 回覆為空，跳過", flush=True)
                last_hash = latest_hash

        except Exception as e:
            print(f"[ERR] {e}", flush=True)
            import traceback
            traceback.print_exc()
            time.sleep(5)

    reason = "收到停止信號" if _should_stop else "到達結束時間"
    print(f"[{datetime.now().strftime('%H:%M:%S')}] 監控結束（{reason}）", flush=True)


# ============================================================
# 主流程
# ============================================================
def main(contact_name, stop_time, sop_path=DEFAULT_SOP, monitor=None):
    import threading

    # 啟動前殺殘留進程（防 VRAM 雙倍佔用）
    _kill_stale_gpu_processes()

    # 清殘留旗標
    if os.path.exists(STOP_FILE):
        try:
            os.remove(STOP_FILE)
        except OSError:
            pass

    print("=" * 50, flush=True)
    print(f"LINE 自動回覆（SOP 引擎）", flush=True)
    print(f"好友：{contact_name}", flush=True)
    print(f"監控到：{stop_time}", flush=True)
    print(f"SOP：{sop_path}", flush=True)
    print("=" * 50, flush=True)

    # 載入 SOP
    print(f"\n[Init] 載入 SOP 設定檔...", flush=True)
    sop = load_sop(sop_path)
    system_prompt = build_system_prompt(sop)
    print(f"[Init] SOP: {sop['name']}（{len(sop['steps'])} 步驟, {len(sop['few_shot_examples'])} 範例）", flush=True)

    # 背景預載 OCR（用 line_locate.py 的引擎）
    def _preload_ocr():
        try:
            from 反詐_locate import _get_ocr_engine
            _get_ocr_engine()
            print("[Preload] PaddleOCR GPU 已載入", flush=True)
        except Exception as e:
            print(f"[Preload] PaddleOCR 載入失敗：{e}", flush=True)

    ocr_thread = threading.Thread(target=_preload_ocr, daemon=True)
    ocr_thread.start()

    # 搜尋好友 + 進入對話
    regions = enter_conversation(contact_name, monitor)
    if regions is None:
        print("[ERROR] 無法進入對話，中止", flush=True)
        return False

    if should_stop():
        return False

    # 等 OCR 預載完成
    for _ in range(15):
        if should_stop():
            return False
        if not ocr_thread.is_alive():
            break
        time.sleep(1)

    # 建立 conversation_history（完整對話記錄，發的收的都記）
    conversation_history = []

    # 驗證 input_box 在 LINE 視窗範圍內才發歡迎詞
    ib = regions["input_box"]
    ix, iy = ib["center"]
    chat_left = regions["chat_area"].get("left", 0)
    valid_input = (ix != 0 and iy != 0 and
                   ix >= chat_left and
                   ib["bottom"] - ib["top"] > 5)

    if not valid_input:
        print(f"[Init] input_box 座標異常 ({ix},{iy})，跳過歡迎詞避免發到錯誤位置", flush=True)
    else:
        # 發送 SOP 歡迎詞（主動開場）
        print(f"\n[Init] 發送 SOP 歡迎詞...", flush=True)
        welcome_step = None
        ask_area_step = None
        for step in sop["steps"]:
            if step["id"] == "welcome":
                welcome_step = step
            elif step["id"] == "ask_area":
                ask_area_step = step

        if welcome_step:
            for reply in welcome_step.get("replies", []):
                if reply == "{send_image}":
                    # 發送 SOP 設定的課程圖片
                    img_path = sop.get("course_info", {}).get("image", "")
                    if img_path:
                        send_image(img_path, regions)
                    continue
                if reply.startswith("{") and reply.endswith("}"):
                    print(f"[Init] 跳過佔位符: {reply}", flush=True)
                    continue
                send_reply(reply, regions)
                conversation_history.append({"text": reply, "sender": "me", "y": 0})
                print(f"[Init] → {reply[:60]}", flush=True)
                time.sleep(1)

        if ask_area_step:
            for reply in ask_area_step.get("replies", []):
                send_reply(reply, regions)
                conversation_history.append({"text": reply, "sender": "me", "y": 0})
                print(f"[Init] → {reply[:60]}", flush=True)
                time.sleep(1)

        print(f"[Init] 歡迎詞已發送（{len(conversation_history)} 條記錄），等待客戶回覆...", flush=True)
        time.sleep(2)

    # 開始監控（傳入 conversation_history）
    print(f"\n[Monitor] 開始自動回覆...", flush=True)
    monitor_and_reply(regions, stop_time, system_prompt, conversation_history, contact_name, monitor)

    print("\n完成", flush=True)
    return True


if __name__ == "__main__":
    if len(sys.argv) < 3:
        print(__doc__)
        print("用法：python line_auto_chat.py <好友名稱> <監控時間HH:MM> [SOP設定檔.json]")
        sys.exit(0)

    contact = sys.argv[1]
    stop = sys.argv[2]
    sop = sys.argv[3] if len(sys.argv) > 3 else DEFAULT_SOP
    main(contact, stop, sop)
