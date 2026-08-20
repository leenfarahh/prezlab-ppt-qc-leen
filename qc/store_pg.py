"""Postgres/Supabase backend for the store, activated when DATABASE_URL is set.

Mirrors the public interface of qc/store.py 1:1 so it is a drop-in: store.py
imports these names over its SQLite implementations at import time when a
DATABASE_URL is present. Schema is created idempotently on first connect, so a
fresh Supabase project needs no manual SQL (supabase/schema.sql documents the
same shape).

Same confidentiality posture as everywhere: this holds users, sessions,
judgments, comments, and audit metadata/manifests. Deck files are never stored.
"""

import hashlib
import json
import secrets
import threading
from datetime import datetime, timezone

import psycopg
from psycopg.rows import dict_row

from .config import DATABASE_URL

ROLES = ("designer", "lead", "admin")
_lock = threading.Lock()
_initialized = False

_SCHEMA = """
create table if not exists users (
    id bigint generated always as identity primary key,
    name text not null unique,
    role text not null check (role in ('designer','lead','admin')),
    pin_hash text,
    created_at timestamptz not null default now());
create table if not exists sessions (
    token text primary key,
    user_name text not null,
    created_at timestamptz not null default now());
create table if not exists comments (
    id bigint generated always as identity primary key,
    deck text not null, slide_index int not null, record_id text,
    author text not null, text text not null,
    created_at timestamptz not null default now());
create index if not exists ix_comments_deck on comments (deck);
create table if not exists audits (
    id bigint generated always as identity primary key,
    deck text not null, profile_id text not null, profile_version int not null,
    user_name text not null, slides int not null, errors int not null,
    warnings int not null, info int not null, arabic int not null,
    total int not null, kind text not null default 'audit',
    manifest jsonb not null, created_at timestamptz not null default now());
create table if not exists triage (
    id bigint generated always as identity primary key,
    record_id text not null, issue_type text not null, module text not null,
    severity text not null, confidence text not null, arabic_flag boolean not null,
    state text not null, deck text not null, profile_id text not null,
    author text, message text, created_at timestamptz not null default now());
create index if not exists ix_triage_issue on triage (issue_type);
"""


def _conn():
    return psycopg.connect(DATABASE_URL, row_factory=dict_row)


def init_db() -> None:
    global _initialized
    if _initialized:
        return
    with _lock, _conn() as c:
        c.execute(_SCHEMA)
    _initialized = True


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


def _iso(v) -> str:
    return v.isoformat(timespec="seconds") if isinstance(v, datetime) else str(v)


# --- users -----------------------------------------------------------------


def add_user(name: str, role: str) -> dict:
    name = name.strip()
    if not name:
        raise ValueError("name required")
    if role not in ROLES:
        raise ValueError(f"role must be one of {ROLES}")
    init_db()
    with _lock, _conn() as c:
        c.execute("INSERT INTO users (name, role, created_at) VALUES (%s, %s, now())"
                  " ON CONFLICT (name) DO NOTHING", (name, role))
        row = c.execute("SELECT * FROM users WHERE name = %s", (name,)).fetchone()
    return _user(row)


def _user(row) -> dict | None:
    if row is None:
        return None
    row = dict(row)
    row["created_at"] = _iso(row.get("created_at"))
    return row


def list_users() -> list[dict]:
    init_db()
    with _lock, _conn() as c:
        return [_user(r) for r in c.execute(
            "SELECT * FROM users ORDER BY name").fetchall()]


def get_user(name: str) -> dict | None:
    init_db()
    with _lock, _conn() as c:
        return _user(c.execute("SELECT * FROM users WHERE name = %s",
                               (name,)).fetchone())


# --- comments --------------------------------------------------------------


def add_comment(deck, slide_index, author, text, record_id=None) -> dict:
    text = text.strip()
    if not text:
        raise ValueError("empty comment")
    init_db()
    with _lock, _conn() as c:
        row = c.execute(
            "INSERT INTO comments (deck, slide_index, record_id, author, text,"
            " created_at) VALUES (%s, %s, %s, %s, %s, now()) RETURNING *",
            (deck, slide_index, record_id, author, text)).fetchone()
    row = dict(row)
    row["created_at"] = _iso(row["created_at"])
    return row


def comments_for(deck, slide_index=None) -> list[dict]:
    init_db()
    with _lock, _conn() as c:
        if slide_index is None:
            rows = c.execute("SELECT * FROM comments WHERE deck = %s ORDER BY id",
                             (deck,)).fetchall()
        else:
            rows = c.execute("SELECT * FROM comments WHERE deck = %s AND"
                             " slide_index = %s ORDER BY id",
                             (deck, slide_index)).fetchall()
    out = []
    for r in rows:
        r = dict(r)
        r["created_at"] = _iso(r["created_at"])
        out.append(r)
    return out


def comment_counts(deck) -> dict:
    init_db()
    with _lock, _conn() as c:
        rows = c.execute("SELECT slide_index, COUNT(*) AS n FROM comments"
                         " WHERE deck = %s GROUP BY slide_index", (deck,)).fetchall()
    return {r["slide_index"]: r["n"] for r in rows}


def delete_comment(comment_id, author) -> bool:
    init_db()
    user = get_user(author)
    with _lock, _conn() as c:
        if user and user["role"] in ("lead", "admin"):
            cur = c.execute("DELETE FROM comments WHERE id = %s", (comment_id,))
        else:
            cur = c.execute("DELETE FROM comments WHERE id = %s AND author = %s",
                            (comment_id, author))
        return cur.rowcount > 0


# --- auth ------------------------------------------------------------------


def _hash_pin(pin, salt):
    return hashlib.pbkdf2_hmac("sha256", pin.encode(), salt.encode(), 200_000).hex()


def set_pin(name, pin) -> None:
    if len(pin.strip()) < 4:
        raise ValueError("PIN must be at least 4 characters")
    init_db()
    salt = secrets.token_hex(8)
    with _lock, _conn() as c:
        c.execute("UPDATE users SET pin_hash = %s WHERE name = %s",
                  (f"{salt}${_hash_pin(pin.strip(), salt)}", name))


def verify_pin(name, pin) -> bool:
    user = get_user(name)
    if not user or not user.get("pin_hash"):
        return False
    salt, expected = user["pin_hash"].split("$", 1)
    return secrets.compare_digest(_hash_pin(pin.strip(), salt), expected)


def has_pin(name) -> bool:
    user = get_user(name)
    return bool(user and user.get("pin_hash"))


def clear_pin(name) -> None:
    init_db()
    with _lock, _conn() as c:
        c.execute("UPDATE users SET pin_hash = NULL WHERE name = %s", (name,))
        c.execute("DELETE FROM sessions WHERE user_name = %s", (name,))


def create_session(name) -> str:
    init_db()
    token = secrets.token_urlsafe(32)
    with _lock, _conn() as c:
        c.execute("INSERT INTO sessions (token, user_name, created_at)"
                  " VALUES (%s, %s, now())", (token, name))
    return token


def session_user(token) -> dict | None:
    if not token:
        return None
    init_db()
    with _lock, _conn() as c:
        return _user(c.execute(
            "SELECT u.* FROM sessions s JOIN users u ON u.name = s.user_name"
            " WHERE s.token = %s", (token,)).fetchone())


def delete_session(token) -> None:
    init_db()
    with _lock, _conn() as c:
        c.execute("DELETE FROM sessions WHERE token = %s", (token,))


# --- audit history ---------------------------------------------------------


def record_audit(manifest, user_name, kind="audit") -> int:
    init_db()
    s = manifest["summary"]
    sev = s.get("by_severity", {})
    with _lock, _conn() as c:
        row = c.execute(
            "INSERT INTO audits (deck, profile_id, profile_version, user_name,"
            " slides, errors, warnings, info, arabic, total, kind, manifest,"
            " created_at) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s,%s, now())"
            " RETURNING id",
            (manifest["deck"], manifest["profile_id"], manifest["profile_version"],
             user_name, manifest["slides"], sev.get("error", 0),
             sev.get("warning", 0), sev.get("info", 0), s.get("arabic_flagged", 0),
             s.get("total", 0), kind, json.dumps(manifest))).fetchone()
    return row["id"]


def list_audits(limit=100) -> list[dict]:
    init_db()
    with _lock, _conn() as c:
        rows = c.execute(
            "SELECT id, deck, profile_id, profile_version, user_name, slides,"
            " errors, warnings, info, arabic, total, kind, created_at"
            " FROM audits ORDER BY id DESC LIMIT %s", (limit,)).fetchall()
    out = []
    for r in rows:
        r = dict(r)
        r["created_at"] = _iso(r["created_at"])
        out.append(r)
    return out


def get_audit(audit_id) -> dict | None:
    init_db()
    with _lock, _conn() as c:
        row = c.execute("SELECT * FROM audits WHERE id = %s", (audit_id,)).fetchone()
    if row is None:
        return None
    row = dict(row)
    row["created_at"] = _iso(row["created_at"])
    if isinstance(row["manifest"], str):
        row["manifest"] = json.loads(row["manifest"])
    return row


# --- triage (mirrors qc/triage.py public interface) ------------------------


def log_triage(record, state, deck, profile_id) -> None:
    init_db()
    with _lock, _conn() as c:
        c.execute(
            "INSERT INTO triage (record_id, issue_type, module, severity,"
            " confidence, arabic_flag, state, deck, profile_id, message,"
            " created_at) VALUES (%s,%s,%s,%s,%s,%s,%s,%s,%s,%s, now())",
            (record["record_id"], record["issue_type"], record["module"],
             record["severity"], record["confidence"], bool(record["arabic_flag"]),
             state, deck, profile_id, (record.get("message") or "")[:300]))


def stats() -> list[dict]:
    """Latest state per record, aggregated per issue_type (mirrors triage.stats)."""
    init_db()
    with _lock, _conn() as c:
        rows = c.execute(
            "SELECT DISTINCT ON (record_id) record_id, issue_type, module, state"
            " FROM triage ORDER BY record_id, id DESC").fetchall()
    agg: dict[str, dict] = {}
    for r in rows:
        if r["state"] == "cleared":
            continue
        row = agg.setdefault(r["issue_type"],
                             {"module": r["module"], "reviewed": 0,
                              "confirmed": 0, "false_alarms": 0})
        row["module"] = r["module"]
        row["reviewed"] += 1
        if r["state"] == "confirmed":
            row["confirmed"] += 1
        else:
            row["false_alarms"] += 1
    out = []
    for issue_type, row in agg.items():
        out.append({
            "issue_type": issue_type, "module": row["module"],
            "reviewed": row["reviewed"], "confirmed": row["confirmed"],
            "false_alarms": row["false_alarms"],
            "fp_rate": row["false_alarms"] / row["reviewed"] if row["reviewed"] else 0.0,
        })
    return sorted(out, key=lambda r: r["reviewed"], reverse=True)
