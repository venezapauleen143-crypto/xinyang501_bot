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


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("--persona", default=None,
                        help="只處理特定 persona（不指定就全部）")
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

    print(f"\n輸出 CSV：")
    export_per_customer(all_profiles, GLOBAL_LOGS_DIR / f"aggregated_per_customer{suffix}_{ts}.csv")
    export_daily_summary(all_events, GLOBAL_LOGS_DIR / f"aggregated_daily{suffix}_{ts}.csv")
    export_disclosure_timeline(all_profiles, GLOBAL_LOGS_DIR / f"aggregated_disclosures{suffix}_{ts}.csv")

    print(f"\n輸出目錄：{GLOBAL_LOGS_DIR}")


if __name__ == "__main__":
    main()
