"""反詐 demo 離線分析工具

讀取 histories/ 內所有 .profile.json + logs/events_*.jsonl，
輸出 CSV 報表（每客戶一行 + 每日彙總 + 事件時序）。

用法：
    python aggregate_stats.py
    → 輸出 logs/aggregated_*.csv
"""
import io
import json
import csv
from datetime import datetime
from pathlib import Path
from collections import defaultdict

HISTORIES_DIR = Path("C:/Users/blue_/claude-telegram-bot/scripts/demo/histories")
LOGS_DIR = Path("C:/Users/blue_/claude-telegram-bot/scripts/demo/logs")
OUT_DIR = LOGS_DIR


def _load_profiles():
    """讀所有 .profile.json，回傳 [(path, dict), ...]"""
    profiles = []
    for f in HISTORIES_DIR.rglob("*.profile.json"):
        try:
            with io.open(f, "r", encoding="utf-8") as fh:
                p = json.load(fh)
            profiles.append((f, p))
        except Exception as e:
            print(f"  load fail: {f.name}: {e}")
    return profiles


def _load_events():
    """讀所有 events_*.jsonl，回傳 list of dict"""
    events = []
    for f in sorted(LOGS_DIR.glob("events_*.jsonl")):
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


def export_per_customer(profiles, events):
    """每客戶一行的 CSV：name / occupation / location / disclosures / total_turns / suspicion / first_seen / last_updated"""
    out = OUT_DIR / f"aggregated_per_customer_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    rows = []
    for path, p in profiles:
        core = p.get("core_facts") or {}
        rows.append({
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
            "file_path": str(path.relative_to(HISTORIES_DIR.parent)),
        })
    if not rows:
        print("  no profiles found")
        return None
    fieldnames = list(rows[0].keys())
    with io.open(out, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"  per_customer: {len(rows)} 筆 → {out.name}")
    return out


def export_daily_summary(events):
    """每日一行：date / customers_seen / replies_sent / suspicions / rate_limits / ocr_failures"""
    out = OUT_DIR / f"aggregated_daily_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    daily = defaultdict(lambda: defaultdict(int))
    daily_unique_customers = defaultdict(set)
    for ev in events:
        ts = ev.get("ts", "")
        if not ts:
            continue
        date = ts[:10]
        et = ev.get("type")
        daily[date][et] += 1
        if ev.get("customer"):
            daily_unique_customers[date].add(ev["customer"])

    if not daily:
        print("  no events found")
        return None

    rows = []
    for date in sorted(daily.keys()):
        d = daily[date]
        rows.append({
            "date": date,
            "unique_customers": len(daily_unique_customers[date]),
            "customer_seen": d.get("customer_seen", 0),
            "new_messages": d.get("new_messages", 0),
            "reply_sent": d.get("reply_sent", 0),
            "profile_first_extract": d.get("profile_first_extract", 0),
            "profile_updated": d.get("profile_updated", 0),
            "ai_suspicion": d.get("ai_suspicion", 0),
            "ocr_failed": d.get("ocr_failed", 0),
            "rate_limit": d.get("rate_limit", 0),
        })
    fieldnames = list(rows[0].keys())
    with io.open(out, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"  daily: {len(rows)} 天 → {out.name}")
    return out


def export_disclosure_timeline(profiles):
    """每筆 disclosure 一行：customer / fact / speaker / day_n（用 milestones 推算）"""
    out = OUT_DIR / f"aggregated_disclosures_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
    rows = []
    for path, p in profiles:
        cust = p.get("name") or path.stem.replace(".profile", "")
        for d in (p.get("shared_disclosures") or []):
            rows.append({
                "customer": cust,
                "speaker": d.get("speaker") or "",
                "fact": d.get("fact") or "",
                "timestamp": d.get("timestamp") or "",
            })
    if not rows:
        print("  no disclosures")
        return None
    fieldnames = list(rows[0].keys())
    with io.open(out, "w", encoding="utf-8-sig", newline="") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        writer.writerows(rows)
    print(f"  disclosures: {len(rows)} 筆 → {out.name}")
    return out


def main():
    print("=" * 60)
    print("反詐 demo 資料 aggregate")
    print("=" * 60)
    print(f"\n載入 profiles ...")
    profiles = _load_profiles()
    print(f"  共 {len(profiles)} 個 profile")
    print(f"\n載入 events ...")
    events = _load_events()
    print(f"  共 {len(events)} 筆 events")

    print(f"\n輸出 CSV：")
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    export_per_customer(profiles, events)
    export_daily_summary(events)
    export_disclosure_timeline(profiles)

    print(f"\n輸出目錄：{OUT_DIR}")


if __name__ == "__main__":
    main()
