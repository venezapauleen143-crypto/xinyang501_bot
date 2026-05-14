"""反詐 demo 離線分析工具（多 persona 版）

讀取 personas/{persona}/histories/.../*.profile.json + personas/{persona}/logs/events_*.jsonl，
輸出 CSV（含 persona 維度）。

用法：
    python aggregate_stats.py              # 跑全部 active personas
    python aggregate_stats.py --persona Angela   # 只跑 Angela
"""
import io
import json
import csv
import argparse
from datetime import datetime
from pathlib import Path
from collections import defaultdict

SCRIPT_DIR = Path(__file__).resolve().parent
PERSONAS_DIR = SCRIPT_DIR / "personas"
GLOBAL_LOGS_DIR = SCRIPT_DIR / "logs"


def list_all_personas():
    """掃 personas/ 目錄找所有 persona"""
    if not PERSONAS_DIR.exists():
        return []
    return [p.name for p in PERSONAS_DIR.iterdir() if p.is_dir()]


def _persona_histories(persona):
    return PERSONAS_DIR / persona / "histories"


def _persona_logs(persona):
    return PERSONAS_DIR / persona / "logs"


def _load_profiles(persona):
    """讀該 persona 所有 .profile.json"""
    profiles = []
    histories_dir = _persona_histories(persona)
    if not histories_dir.exists():
        return profiles
    for f in histories_dir.rglob("*.profile.json"):
        try:
            with io.open(f, "r", encoding="utf-8") as fh:
                p = json.load(fh)
            profiles.append((persona, f, p))
        except Exception as e:
            print(f"  load fail: {f.name}: {e}")
    return profiles


def _load_events(persona):
    """讀該 persona 所有 events_*.jsonl"""
    events = []
    logs_dir = _persona_logs(persona)
    if not logs_dir.exists():
        return events
    for f in sorted(logs_dir.glob("events_*.jsonl")):
        try:
            with io.open(f, "r", encoding="utf-8") as fh:
                for line in fh:
                    line = line.strip()
                    if not line:
                        continue
                    try:
                        events.append(json.loads(line))
                    except json.JSONDecodeError:
                        continue
        except Exception as e:
            print(f"  load fail: {f.name}: {e}")
    return events


def export_per_customer(all_profiles_with_persona, out_path):
    """每客戶一行 CSV，含 persona 欄位"""
    rows = []
    for persona, path, p in all_profiles_with_persona:
        core = p.get("core_facts") or {}
        rows.append({
            "persona": persona,
            "name": p.get("name") or path.stem.replace(".profile", ""),
            "occupation": core.get("occupation") or "",
            "location": core.get("location") or "",
            "schedule": core.get("schedule") or "",
            "age": core.get("age") or "",
            "gender": core.get("gender") or "",
            "marital_status": core.get("marital_status") or "",
            "interests_count": len(p.get("interests") or []),
            "family_count": len(p.get("family_relationships") or []),
            "disclosures_count": len(p.get("shared_disclosures") or []),
            "milestones_count": len(p.get("milestones") or []),
            "current_stage": p.get("current_stage") or "",
            "total_turns": p.get("total_turns") or 0,
            "trust_score": p.get("trust_score") or 0,
            "suspicion_count": len(p.get("ai_suspicion_flags") or []),
            "first_seen": p.get("first_seen") or "",
            "last_updated": p.get("last_updated") or "",
            "file_path": str(path.relative_to(PERSONAS_DIR.parent)),
        })
    if not rows:
        print("  no profiles found")
        return None
    fieldnames = list(rows[0].keys())
    with io.open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"  per_customer: {len(rows)} 筆 → {out_path.name}")
    return out_path


def export_daily_summary(all_events_with_persona, out_path):
    """每天每 persona 一行：date / persona / customers_seen / replies_sent / suspicions ..."""
    daily = defaultdict(lambda: defaultdict(int))
    daily_unique_customers = defaultdict(set)
    for ev in all_events_with_persona:
        ts = ev.get("ts", "")
        if not ts:
            continue
        date = ts[:10]
        persona = ev.get("persona", "unknown")
        et = ev.get("type")
        key = (date, persona)
        daily[key][et] += 1
        if ev.get("customer"):
            daily_unique_customers[key].add(ev["customer"])

    if not daily:
        print("  no events found")
        return None

    rows = []
    for (date, persona) in sorted(daily.keys()):
        d = daily[(date, persona)]
        rows.append({
            "date": date,
            "persona": persona,
            "unique_customers": len(daily_unique_customers[(date, persona)]),
            "customer_seen": d.get("customer_seen", 0),
            "new_messages": d.get("new_messages", 0),
            "reply_sent": d.get("reply_sent", 0),
            "profile_first_extract": d.get("profile_first_extract", 0),
            "profile_updated": d.get("profile_updated", 0),
            "ai_suspicion": d.get("ai_suspicion", 0),
            "ocr_failed": d.get("ocr_failed", 0),
            "rate_limit": d.get("rate_limit", 0),
            "profile_failed": d.get("profile_failed", 0),
        })
    fieldnames = list(rows[0].keys())
    with io.open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"  daily: {len(rows)} 行 → {out_path.name}")
    return out_path


def export_disclosure_timeline(all_profiles_with_persona, out_path):
    """每筆 disclosure 一行：persona / customer / fact / speaker"""
    rows = []
    for persona, path, p in all_profiles_with_persona:
        cust = p.get("name") or path.stem.replace(".profile", "")
        for d in (p.get("shared_disclosures") or []):
            rows.append({
                "persona": persona,
                "customer": cust,
                "speaker": d.get("speaker") or "",
                "fact": d.get("fact") or "",
                "timestamp": d.get("timestamp") or "",
            })
    if not rows:
        print("  no disclosures")
        return None
    fieldnames = list(rows[0].keys())
    with io.open(out_path, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"  disclosures: {len(rows)} 筆 → {out_path.name}")
    return out_path


import re


def _extract_current_day(profile):
    """從 profile.current_stage 或 milestones parse 出 Day X 數字"""
    stage = profile.get("current_stage") or ""
    m = re.search(r"Day\s*(\d+)", stage)
    if m:
        return int(m.group(1))
    milestones = profile.get("milestones") or []
    max_day = 0
    for ms in milestones:
        m = re.search(r"Day\s*(\d+)", ms)
        if m:
            max_day = max(max_day, int(m.group(1)))
    return max_day if max_day > 0 else None


def _detect_persona_alerts(persona):
    """Rule 1, 2：persona 級 health 偵測（容錯 Angela 沒檔案）"""
    alerts = []
    persona_dir = PERSONAS_DIR / persona

    # Rule 1: heartbeat > 1 小時沒動
    hb = persona_dir / "heartbeat.json"
    if hb.exists():
        try:
            with io.open(hb, "r", encoding="utf-8") as f:
                data = json.load(f)
            last_hb = data.get("last_heartbeat")
            if last_hb:
                last_dt = datetime.fromisoformat(last_hb)
                idle_sec = (datetime.now() - last_dt).total_seconds()
                if idle_sec > 3600:
                    alerts.append({
                        "level": "warning", "rule": "persona_idle",
                        "persona": persona,
                        "msg": f"{persona} 已 {idle_sec/3600:.1f} 小時沒 heartbeat",
                    })
        except Exception:
            pass

    # Rule 2: click_failures.jsonl 累計 > 10
    cf = persona_dir / "click_failures.jsonl"
    if cf.exists():
        try:
            count = sum(1 for line in io.open(cf, "r", encoding="utf-8") if line.strip())
            if count > 10:
                alerts.append({
                    "level": "warning", "rule": "click_failures_high",
                    "persona": persona,
                    "msg": f"{persona} click_failures 累計 {count} (>10)",
                })
        except Exception:
            pass

    return alerts


def _detect_customer_alerts(persona, profile, profile_path):
    """Rule 3, 4, 5：客戶級 alert"""
    alerts = []
    name = profile.get("name") or profile_path.stem.replace(".profile", "")
    current_day = _extract_current_day(profile)

    # Rule 3: Customer idle (last_updated > 12h)
    last_upd = profile.get("last_updated")
    if last_upd:
        try:
            last_dt = datetime.fromisoformat(last_upd)
            idle_h = (datetime.now() - last_dt).total_seconds() / 3600
            if idle_h > 12:
                alerts.append({
                    "level": "info", "rule": "customer_idle",
                    "persona": persona, "customer": name,
                    "msg": f"{name} 已 {idle_h:.0f}h 沒互動 (current_day={current_day})",
                })
        except Exception:
            pass

    # Rule 4: Day stagnation (milestones 數量 vs current_day 比例)
    if current_day is not None and current_day >= 3:
        milestones = profile.get("milestones") or []
        expected = current_day * 2
        if len(milestones) < expected * 0.5:
            alerts.append({
                "level": "warning", "rule": "day_stagnation",
                "persona": persona, "customer": name,
                "msg": f"{name} Day={current_day} 但只有 {len(milestones)} milestones (期望 ≥ {int(expected*0.5)})",
            })

    # Rule 5: Trust progression (Day >= 13 but trust_score < 6)
    trust = profile.get("trust_score") or 0
    if current_day is not None and current_day >= 13 and trust < 6:
        alerts.append({
            "level": "warning", "rule": "trust_low",
            "persona": persona, "customer": name,
            "msg": f"{name} 已 Day {current_day} 但 trust_score={trust} (期望 ≥ 6)",
        })

    return alerts


def _compare_personas(profiles_by_persona):
    """Rule 6: cross-persona imbalance"""
    alerts = []
    counts = {p: len(profs) for p, profs in profiles_by_persona.items()}
    if len(counts) < 2:
        return alerts
    max_p = max(counts, key=counts.get)
    min_p = min(counts, key=counts.get)
    if counts[min_p] > 0 and counts[max_p] / counts[min_p] > 3:
        alerts.append({
            "level": "warning", "rule": "persona_imbalance",
            "msg": f"persona 客戶數失衡：{max_p}={counts[max_p]} vs {min_p}={counts[min_p]} (>3:1)",
        })
    return alerts


def export_alerts(all_profiles_with_persona, out_path):
    """整合 Rule 1-6 寫 alerts JSONL"""
    all_alerts = []

    profiles_by_persona = defaultdict(list)
    for persona, path, p in all_profiles_with_persona:
        profiles_by_persona[persona].append((path, p))

    for persona in profiles_by_persona:
        all_alerts.extend(_detect_persona_alerts(persona))

    for persona, path, p in all_profiles_with_persona:
        all_alerts.extend(_detect_customer_alerts(persona, p, path))

    all_alerts.extend(_compare_personas(profiles_by_persona))

    if not all_alerts:
        print("  ✓ 無 alert（demo 健康）")
        return None

    ts = datetime.now().isoformat()
    for a in all_alerts:
        a["ts"] = ts

    with io.open(out_path, "a", encoding="utf-8") as f:
        for a in all_alerts:
            f.write(json.dumps(a, ensure_ascii=False) + "\n")

    print(f"\n  ⚠ {len(all_alerts)} alerts 觸發：")
    for a in all_alerts:
        lvl = a["level"].upper()
        print(f"    [{lvl}] [{a['rule']}] {a['msg']}")

    return out_path


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--persona", default=None,
                        help="只處理特定 persona（不指定就全部）")
    parser.add_argument("--alerts", action="store_true",
                        help="只跑 alert 偵測（B2 自動巡邏模式）")
    args = parser.parse_args()

    print("=" * 60)
    print("反詐 demo 資料 aggregate（多 persona 版）")
    print("=" * 60)

    targets = [args.persona] if args.persona else list_all_personas()
    if not targets:
        print(f"\n沒找到任何 persona 目錄（{PERSONAS_DIR}）")
        return
    print(f"\n處理 personas: {targets}")

    all_profiles = []
    all_events = []
    for persona in targets:
        print(f"\n[{persona}] 載入 profiles ...")
        profs = _load_profiles(persona)
        print(f"  {len(profs)} 個 profile")
        all_profiles.extend(profs)

        print(f"[{persona}] 載入 events ...")
        evs = _load_events(persona)
        print(f"  {len(evs)} 筆 events")
        all_events.extend(evs)

    GLOBAL_LOGS_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now().strftime("%Y%m%d_%H%M%S")
    suffix = f"_{args.persona}" if args.persona else ""

    if args.alerts:
        # B2 自動巡邏模式：只跑 alert 偵測，不出 CSV
        today = datetime.now().strftime("%Y-%m-%d")
        alert_path = GLOBAL_LOGS_DIR / f"alerts_{today}.jsonl"
        print(f"\n[B2 自動巡邏] 偵測 alerts ...")
        export_alerts(all_profiles, alert_path)
        print(f"\n輸出：{alert_path}")
        return

    print(f"\n輸出 CSV：")
    export_per_customer(all_profiles, GLOBAL_LOGS_DIR / f"aggregated_per_customer{suffix}_{ts}.csv")
    export_daily_summary(all_events, GLOBAL_LOGS_DIR / f"aggregated_daily{suffix}_{ts}.csv")
    export_disclosure_timeline(all_profiles, GLOBAL_LOGS_DIR / f"aggregated_disclosures{suffix}_{ts}.csv")

    print(f"\n輸出目錄：{GLOBAL_LOGS_DIR}")


if __name__ == "__main__":
    main()
