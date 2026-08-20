"""Triage capture: designer judgments on findings, persisted locally.

Every confirm / false-alarm click appends one JSONL line to
data/triage-log.jsonl. This file is the pilot's ground truth: per-issue-type
false-alarm rates computed from it are what gate which fixes graduate to
one-click (PRD: precision gates auto-apply). The log stays on this machine;
it contains finding metadata (issue types, severities, messages) but never
deck files.

Toggling a judgment off logs state "cleared"; aggregation keeps only the
LATEST state per record so changed minds do not pollute the stats.
"""

import json
import threading
from collections import defaultdict
from datetime import datetime, timezone
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
TRIAGE_LOG = DATA_DIR / "triage-log.jsonl"
STATES = ("confirmed", "false_positive", "cleared")

_lock = threading.Lock()


def log_triage(record: dict, state: str, deck: str, profile_id: str) -> None:
    assert state in STATES, state
    entry = {
        "at": datetime.now(timezone.utc).isoformat(timespec="seconds"),
        "deck": deck,
        "profile_id": profile_id,
        "record_id": record["record_id"],
        "issue_type": record["issue_type"],
        "module": record["module"],
        "severity": record["severity"],
        "confidence": record["confidence"],
        "arabic_flag": record["arabic_flag"],
        "state": state,
        "message": record.get("message", "")[:300],
    }
    with _lock:
        DATA_DIR.mkdir(exist_ok=True)
        with TRIAGE_LOG.open("a", encoding="utf-8") as f:
            f.write(json.dumps(entry, ensure_ascii=False) + "\n")


def stats() -> list[dict]:
    """Per-issue-type aggregate from the latest state of every record.

    Returns rows sorted by reviewed count desc:
    {issue_type, module, reviewed, confirmed, false_alarms, fp_rate}."""
    if not TRIAGE_LOG.exists():
        return []
    latest: dict[str, dict] = {}
    with _lock:
        for line in TRIAGE_LOG.read_text(encoding="utf-8").splitlines():
            if not line.strip():
                continue
            try:
                entry = json.loads(line)
            except json.JSONDecodeError:
                continue
            latest[entry["record_id"]] = entry

    agg: dict[str, dict] = defaultdict(
        lambda: {"reviewed": 0, "confirmed": 0, "false_alarms": 0, "module": ""})
    for entry in latest.values():
        if entry["state"] == "cleared":
            continue
        row = agg[entry["issue_type"]]
        row["module"] = entry["module"]
        row["reviewed"] += 1
        if entry["state"] == "confirmed":
            row["confirmed"] += 1
        else:
            row["false_alarms"] += 1

    out = []
    for issue_type, row in agg.items():
        out.append({
            "issue_type": issue_type,
            "module": row["module"],
            "reviewed": row["reviewed"],
            "confirmed": row["confirmed"],
            "false_alarms": row["false_alarms"],
            "fp_rate": row["false_alarms"] / row["reviewed"] if row["reviewed"] else 0.0,
        })
    return sorted(out, key=lambda r: r["reviewed"], reverse=True)


from .config import USING_POSTGRES as _USING_POSTGRES  # noqa: E402

if _USING_POSTGRES:  # pragma: no cover - exercised only against a live DB
    from .store_pg import log_triage, stats  # noqa: F401,E402
