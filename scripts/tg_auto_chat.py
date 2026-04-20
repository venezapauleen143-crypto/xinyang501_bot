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

# 把 scripts 目錄加入 path，讓 tg_locate 可以 import
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))
sys.path.insert(0, str(SCRIPT_DIR.parent))

load_dotenv(Path("C:/Users/blue_/claude-telegram-bot/.env"))
client = anthropic.Anthropic()
TMPDIR = "C:/Users/blue_/Desktop/測試檔案"

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
def locate_all(monitor=2):
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
def search_contact(regions, contact_name, monitor=2):
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
def click_contact(regions, contact_name, monitor=2):
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
def verify_friend(target_name, monitor=2):
    """重新定位，確認好友名稱框裡的名字是不是目標好友"""
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
        return regions, True
    else:
        print(f"[Step 4] ❌ 名稱不符：{detected_name} != {target_name}", flush=True)
        return regions, False


# ============================================================
# Step 5-7: 監控對話 + 自動回覆
# ============================================================
def grab_chat(regions, monitor=2):
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
# 維修 1+2: 分析函數 — 分層分析完整對話截圖
# ============================================================
def analyze_conversation(chat_img):
    """
    Vision 分析完整對話截圖（不裁切）。
    分層邏輯：
    1. 讀完整對話理解上下文
    2. 找到自己最後一條綠色氣泡的位置
    3. 之後所有白色氣泡 = 對方的新訊息
    4. 回傳：sender + 新訊息內容 + 是否需搜尋 + 搜尋關鍵字列表
    """
    tmp = os.path.join(TMPDIR, "tg_analyze.png")
    chat_img.save(tmp, quality=95)
    with open(tmp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    r = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=300,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
            {"type": "text", "text": """這是 Telegram 完整對話截圖。
綠色氣泡（靠右）= 我發的
白色氣泡（靠左）= 對方發的

請按以下步驟分析：

步驟一：找到畫面中【我的最後一條綠色氣泡】的位置。

步驟二：看那條綠色氣泡【之後】有沒有白色氣泡。
- 如果最底部是綠色氣泡（我的）→ sender = "me"
- 如果綠色氣泡之後有白色氣泡（對方的）→ sender = "them"

步驟三：如果 sender = "them"，把對方在我最後一條綠色氣泡之後發的【所有白色氣泡的內容】整理出來（可能有多條）。

步驟四：判斷對方的這些新訊息裡有沒有需要查證的事實性問題（比賽比分、股票價格、新聞事件、數據排名等）。如果有，列出具體要搜尋的關鍵字。

只回 JSON，不要其他文字：
{"sender":"them或me","new_messages":"對方新訊息的完整內容（多條合併）","needs_search":true或false,"search_queries":["要搜的關鍵字1","要搜的關鍵字2"]}"""}
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
    """根據每個 query 的類型用對應的工具搜尋，合併結果"""
    if not search_queries:
        return ""
    try:
        sys.path.insert(0, str(Path("C:/Users/blue_/claude-telegram-bot")))
        from bot import tavily_search, fetch_sports_scores, fetch_stock

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

            if is_sport and not sports_done:
                scores = fetch_sports_scores("nba")
                results.append(f"【ESPN 即時比分】\n{scores}")
                sports_done = True
                # 也用 Tavily 補充
                tr = tavily_search(query, 2, "basic")
                if tr and "失敗" not in tr:
                    results.append(f"【{query} 搜尋】\n{tr}")

            elif is_stock:
                # 判斷是哪支股票
                if "台積電" in query or "TSMC" in query or "2330" in query:
                    stock_result = fetch_stock("2330.TW")
                    results.append(f"【台積電台股報價】\n{stock_result}")
                elif "鴻海" in query:
                    stock_result = fetch_stock("2317.TW")
                    results.append(f"【鴻海台股報價】\n{stock_result}")
                elif "聯發科" in query:
                    stock_result = fetch_stock("2454.TW")
                    results.append(f"【聯發科台股報價】\n{stock_result}")
                else:
                    tr = tavily_search(query, 2, "basic")
                    if tr and "失敗" not in tr:
                        results.append(f"【{query} 搜尋】\n{tr}")
            else:
                tr = tavily_search(query, 2, "basic")
                if tr and "失敗" not in tr:
                    results.append(f"【{query} 搜尋】\n{tr}")

        return "\n\n".join(results) if results else ""
    except Exception as e:
        return f"搜尋失敗：{e}"


# ============================================================
# 維修 2+4+5: 回覆函數 — 帶搜尋結果生成回覆
# ============================================================
def generate_reply_with_context(chat_img, new_messages, search_context=""):
    """
    Vision 看完整對話截圖 + 對方新訊息摘要 + 搜尋結果 → 生成回覆
    維修 5: prompt 禁止引用自己綠色氣泡的數字
    """
    tmp = os.path.join(TMPDIR, "tg_auto_chat.png")
    chat_img.save(tmp, quality=95)
    with open(tmp, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()

    system_prompt = PERSONA
    if search_context:
        system_prompt += (
            "\n\n====== 搜尋到的最新事實資料 ======\n"
            + search_context[:2000]
            + "\n====== 事實資料結束 ======\n\n"
            "【絕對規則】\n"
            "1. 上面搜尋結果裡有的比分、數字、價格 → 必須用搜尋結果的，一個字都不能改\n"
            "2. 搜尋結果裡沒有提到的比分、數字、球員數據 → 絕對不能自己編，說「這個我不確定」\n"
            "3. 不能把 A 隊的比分說成 B 隊的，不能把昨天的說成今天的\n"
            "4. 寧可回答「我不知道細節」也不能編造任何數字\n"
            "5. 【重要】綠色氣泡是你自己之前說的話，裡面的數字不是事實來源。"
            "每次回答都用上面搜尋結果的數字，不要引用自己之前說過的任何數字。\n"
        )

    user_prompt = f"""這是 Telegram 完整對話截圖。

對方在你最後一條綠色氣泡之後發了以下新訊息：
「{new_messages}」

請針對對方的這些新訊息回覆。
- 讀完整對話理解上下文和前因後果
- 但只回覆對方的新訊息，不要重複回覆之前已經回過的內容
- 對方問問題就回答，嗆人就幽默化解，聊天就接話
- 如果對方連發多條，整體理解後一次回覆
- 如果涉及比賽結果、股價等事實 → 只能用搜尋結果的數字
- 搜尋結果沒有的就說「這個我沒查到」

只回覆要發送的文字，不要加任何格式或解釋。"""

    r = client.messages.create(
        model="claude-sonnet-4-6", max_tokens=200, system=system_prompt,
        messages=[{"role": "user", "content": [
            {"type": "image", "source": {"type": "base64", "media_type": "image/png", "data": b64}},
            {"type": "text", "text": user_prompt}
        ]}]
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


def monitor_and_reply(regions, stop_time, monitor=2):
    """
    監控對話區，偵測變化後：
    1. 等 1.5 秒穩定（避免動畫/已讀標記觸發假變化）
    2. 二次截圖確認真的有變化
    3. 等對方說完（最多 2+2=4 秒）
    4. 一次 Vision 分析完整對話（分層：上下文 + 新訊息 + sender）
    5. sender=them → 搜尋（如需要）→ 生成回覆
    """
    STABILIZE_WAIT = 1.5    # 維修 4: 從 3 秒降到 1.5 秒
    WAIT_COMPLETE = 2       # 維修 4: 等對方說完每輪 2 秒（最多 2 輪）
    WAIT_COMPLETE_ROUNDS = 2

    # 先截一張當基準
    chat = grab_chat(regions, monitor)
    last_hash = chat_hash(chat)
    cooldown_until = 0

    print(f"[Monitor] 開始監控 → {stop_time}", flush=True)
    print(f"[Monitor] 冷卻：{COOLDOWN_SECONDS}s / 輪詢：{POLL_INTERVAL}s / 穩定：{STABILIZE_WAIT}s / 等說完：{WAIT_COMPLETE}sx{WAIT_COMPLETE_ROUNDS}", flush=True)

    while datetime.now().strftime("%H:%M") < stop_time:
        time.sleep(POLL_INTERVAL)
        try:
            # 冷卻期間跳過
            if time.time() < cooldown_until:
                continue

            chat = grab_chat(regions, monitor)
            h = chat_hash(chat)

            if h != last_hash:
                t = datetime.now().strftime("%H:%M:%S")
                print(f"[{t}] 偵測到畫面變化，等 {STABILIZE_WAIT}s 穩定...", flush=True)

                # === 穩定等待 ===
                time.sleep(STABILIZE_WAIT)

                # === 二次截圖確認 ===
                chat2 = grab_chat(regions, monitor)
                h2 = chat_hash(chat2)
                if h2 == last_hash:
                    print(f"[{t}] 假變化，跳過", flush=True)
                    last_hash = h2
                    continue

                # === 等對方說完（最多 2 輪 x 2 秒 = 4 秒）===
                latest_chat = chat2
                latest_hash = h2
                for wait_round in range(WAIT_COMPLETE_ROUNDS):
                    time.sleep(WAIT_COMPLETE)
                    check = grab_chat(regions, monitor)
                    check_h = chat_hash(check)
                    if check_h != latest_hash:
                        print(f"[{t}] 對方還在發訊息（第{wait_round+1}輪），等...", flush=True)
                        latest_chat = check
                        latest_hash = check_h
                    else:
                        break  # 穩定了，對方說完了

                # === 維修 1+2: 一次 Vision 分析完整對話 ===
                analysis = analyze_conversation(latest_chat)
                sender = analysis.get("sender", "unknown")
                new_msgs = analysis.get("new_messages", "")
                needs_search = analysis.get("needs_search", False)
                search_queries = analysis.get("search_queries", [])

                print(f"[{t}] 分析：sender={sender} new_msgs={new_msgs[:50]} needs_search={needs_search} queries={search_queries}", flush=True)

                if sender == "me":
                    print(f"[{t}] 最後是自己的訊息，跳過", flush=True)
                    last_hash = latest_hash
                    continue

                if sender == "unknown":
                    print(f"[{t}] 無法判斷發送者，跳過", flush=True)
                    last_hash = latest_hash
                    continue

                # === sender == "them"：對方發了新訊息 ===

                # === 維修 3: 搜尋（用 search_queries，分別用對的工具）===
                search_context = ""
                if needs_search and search_queries:
                    print(f"[{t}] 搜尋：{search_queries}", flush=True)
                    search_context = search_for_queries(search_queries)
                    print(f"[{t}] 搜尋完成（{len(search_context)} chars）", flush=True)

                # === 維修 2+5: 生成回覆（帶新訊息摘要 + 搜尋結果）===
                reply = generate_reply_with_context(latest_chat, new_msgs, search_context)

                if reply and len(reply) > 1:
                    print(f"[{t}] → 回覆：{reply}", flush=True)
                    send_reply(reply, regions)

                    cooldown_until = time.time() + COOLDOWN_SECONDS
                    print(f"[{t}] 冷卻 {COOLDOWN_SECONDS} 秒", flush=True)

                    time.sleep(COOLDOWN_SECONDS)
                    chat = grab_chat(regions, monitor)
                    last_hash = chat_hash(chat)
                else:
                    print(f"[{t}] 生成回覆為空，跳過", flush=True)
                    last_hash = latest_hash
            else:
                last_hash = h

        except Exception as e:
            print(f"[ERR] {e}", flush=True)
            time.sleep(5)

    print(f"[{datetime.now().strftime('%H:%M:%S')}] 監控結束", flush=True)


# ============================================================
# 主流程
# ============================================================
def main(contact_name, stop_time, monitor=2):
    print("=" * 50, flush=True)
    print(f"Telegram 自動回覆", flush=True)
    print(f"好友：{contact_name}", flush=True)
    print(f"監控到：{stop_time}", flush=True)
    print("=" * 50, flush=True)

    # Step 1: 定位
    print("\n[Step 1] 定位 Telegram UI...", flush=True)
    regions = locate_all(monitor)

    # Step 2: 搜尋好友
    print(f"\n[Step 2] 搜尋好友：{contact_name}", flush=True)
    search_contact(regions, contact_name, monitor)

    # Step 3: 點選好友
    print(f"\n[Step 3] 點選好友...", flush=True)
    click_contact(regions, contact_name, monitor)

    # Step 4: 確認好友名稱
    print(f"\n[Step 4] 確認好友名稱...", flush=True)
    regions, confirmed = verify_friend(contact_name, monitor)
    if not confirmed:
        print("[ERROR] 好友名稱不符，中止", flush=True)
        return False

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
