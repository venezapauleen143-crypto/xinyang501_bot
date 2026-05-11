"""LINE UIA 驗證 — 純讀取，不點不動
跑完看 personas/uia_inspect_result.txt 結果。

5 個檢查點：
1. LINE 視窗能否被 pywinauto 找到
2. UIA tree 深度（暴露多少元素）
3. 關鍵元素（ListItem/Edit/Button）拿得到嗎
4. 沙盤是否擋 UIA
5. 性能（descendants 列舉時間）
"""
import sys
import io
import time
from pathlib import Path

# 編碼修正
sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding="utf-8", errors="replace")
sys.stderr = io.TextIOWrapper(sys.stderr.buffer, encoding="utf-8", errors="replace")

output_lines = []

def log(msg):
    print(msg, flush=True)
    output_lines.append(str(msg))


log("=" * 70)
log("LINE UIA 驗證腳本（純讀取，不影響 LINE）")
log("=" * 70)

# ============================================================
# Step 1: 找 LINE 視窗
# ============================================================
log("\n[1/5] 找 LINE 視窗...")
try:
    from pywinauto import Desktop
    desktop = Desktop(backend="uia")
    # 找標題含 LINE 的視窗
    windows = desktop.windows(title_re=".*LINE.*")
    log(f"  pywinauto 找到 {len(windows)} 個視窗標題含 LINE")
    for w in windows[:5]:
        try:
            log(f"    - title={w.window_text()!r}  class={w.class_name()}  pid={w.process_id()}")
        except Exception as e:
            log(f"    - (讀取失敗 {e})")
    if not windows:
        log("  ❌ 找不到任何 LINE 視窗")
        log("  → 確認 demo 沙盤 LINE 開著")
        Path("personas/uia_inspect_result.txt").write_text("\n".join(output_lines), encoding="utf-8")
        sys.exit(1)
    line = windows[0]
    log(f"  ✅ 用第一個視窗測試")
except Exception as e:
    log(f"  ❌ pywinauto 初始化失敗：{type(e).__name__}: {e}")
    Path("personas/uia_inspect_result.txt").write_text("\n".join(output_lines), encoding="utf-8")
    sys.exit(1)


# ============================================================
# Step 2: 印 UIA tree（depth=4，太深會爆）
# ============================================================
log("\n[2/5] 列舉 UIA tree（depth=4）...")
try:
    # capture print_control_identifiers 輸出
    import io as _io
    buf = _io.StringIO()
    old_stdout = sys.stdout
    sys.stdout = buf
    try:
        line.print_control_identifiers(depth=4)
    finally:
        sys.stdout = old_stdout
    tree_text = buf.getvalue()
    log(f"  Tree 長度 {len(tree_text)} 字元")
    # 印前 3000 字（避免太長）
    log("  --- Tree 內容（前 3000 字）---")
    log(tree_text[:3000])
    if len(tree_text) > 3000:
        log(f"  ... (略 {len(tree_text)-3000} 字)")
except Exception as e:
    log(f"  ❌ print_control_identifiers 失敗：{type(e).__name__}: {e}")


# ============================================================
# Step 3: 關鍵元素檢查
# ============================================================
log("\n[3/5] 關鍵元素檢查（list 出每類前 5 個）...")
checks = [
    ("Window", "Window"),
    ("Pane", "Pane"),
    ("ListView (對話列表)", "List"),
    ("ListItem (對話卡)", "ListItem"),
    ("Edit (輸入框)", "Edit"),
    ("Button (按鈕)", "Button"),
    ("Text (文字)", "Text"),
    ("Group", "Group"),
]
for name, ctype in checks:
    try:
        elements = line.descendants(control_type=ctype)
        log(f"  {name}: {len(elements)} 個")
        for e in elements[:5]:
            try:
                txt = e.window_text()
                log(f"    - name={txt[:50]!r}")
            except Exception:
                log(f"    - (讀取失敗)")
    except Exception as e:
        log(f"  {name}: ❌ {type(e).__name__}: {e}")


# ============================================================
# Step 4: 找特定文字 — 看能不能拿到對話 name
# ============================================================
log("\n[4/5] 找特定對話 name（你 LINE 列表現在有的人）...")
targets = ["仁輝", "JAMES", "佳莹", "佳瑩", "你好", "有聲書", "Keep筆記"]
for target in targets:
    try:
        # 用 title_re 部分匹配
        found = line.descendants(title_re=f".*{target}.*")
        log(f"  搜「{target}」：{len(found)} 個")
        for e in found[:3]:
            try:
                log(f"    - control_type={e.element_info.control_type}  name={e.window_text()[:50]!r}")
            except Exception:
                pass
    except Exception as e:
        log(f"  搜「{target}」: ❌ {e}")


# ============================================================
# Step 5: 性能
# ============================================================
log("\n[5/5] 性能測試...")
try:
    t0 = time.time()
    all_desc = line.descendants()
    elapsed = time.time() - t0
    log(f"  列舉所有 descendants：{elapsed:.2f}s（共 {len(all_desc)} 個元素）")
    if elapsed > 5:
        log("  ⚠️ 太慢，不適合每輪 polling")
    elif elapsed > 1:
        log("  🟡 可接受但偏慢")
    else:
        log("  ✅ 性能 OK")
except Exception as e:
    log(f"  ❌ {type(e).__name__}: {e}")


# ============================================================
# 結論
# ============================================================
log("\n" + "=" * 70)
log("結論")
log("=" * 70)
try:
    list_items = line.descendants(control_type="ListItem")
    edits = line.descendants(control_type="Edit")

    has_list_items = len(list_items) > 0
    has_edits = len(edits) > 0
    has_chat_names = False
    for target in ["仁輝", "佳莹", "佳瑩", "你好"]:
        if line.descendants(title_re=f".*{target}.*"):
            has_chat_names = True
            break

    log(f"  ListItem 找到：{has_list_items}")
    log(f"  Edit (輸入框) 找到：{has_edits}")
    log(f"  對話 name 拿得到：{has_chat_names}")

    if has_chat_names and has_list_items:
        log("\n  🟢 結果 A：UIA 完整 → 可走 Step 4 大改造")
    elif has_list_items or has_edits:
        log("\n  🟡 結果 B：UIA 部分 → 混合方案（按鈕 UIA + 對話 OCR）")
    else:
        log("\n  🔴 結果 C：UIA 廢（Electron 預設沒暴露）→ 取消 Step 4")
except Exception as e:
    log(f"  ❌ 結論判定失敗：{e}")


# 存結果
out_path = Path("personas/uia_inspect_result.txt")
out_path.parent.mkdir(parents=True, exist_ok=True)
out_path.write_text("\n".join(output_lines), encoding="utf-8")
log(f"\n結果存到：{out_path.absolute()}")
