"""Local pilot store: users and inline comments (SQLite, stdlib only).

Identity here is ATTRIBUTION, not authentication: pilot users pick their
name so triage judgments, comments, and profile edits carry an author. Real
authentication is Entra ID in the production tier (PRD Section 6); this
deliberately does not pretend to be a security boundary.

Comments are keyed by deck FILENAME (not job id) so they survive the
in-memory job cache and reattach when the same deck is re-audited.
"""

import sqlite3
import threading
from datetime import datetime, timezone
from pathlib import Path

DATA_DIR = Path(__file__).resolve().parent.parent / "data"
DB_PATH = DATA_DIR / "qc.db"
ROLES = ("designer", "lead", "admin")

_lock = threading.Lock()


def _conn() -> sqlite3.Connection:
    DATA_DIR.mkdir(exist_ok=True)
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def init_db() -> None:
    with _lock, _conn() as c:
        c.execute("""CREATE TABLE IF NOT EXISTS users (
            id INTEGER PRIMARY KEY,
            name TEXT NOT NULL UNIQUE,
            role TEXT NOT NULL CHECK (role IN ('designer', 'lead', 'admin')),
            created_at TEXT NOT NULL)""")
        c.execute("""CREATE TABLE IF NOT EXISTS comments (
            id INTEGER PRIMARY KEY,
            deck TEXT NOT NULL,
            slide_index INTEGER NOT NULL,
            record_id TEXT,
            author TEXT NOT NULL,
            text TEXT NOT NULL,
            created_at TEXT NOT NULL)""")
        c.execute("CREATE INDEX IF NOT EXISTS ix_comments_deck ON comments (deck)")


def _now() -> str:
    return datetime.now(timezone.utc).isoformat(timespec="seconds")


# --- users -------------------------------------------------------------


def add_user(name: str, role: str) -> dict:
    name = name.strip()
    if not name:
        raise ValueError("name required")
    if role not in ROLES:
        raise ValueError(f"role must be one of {ROLES}")
    init_db()
    with _lock, _conn() as c:
        c.execute("INSERT OR IGNORE INTO users (name, role, created_at) VALUES (?, ?, ?)",
                  (name, role, _now()))
        row = c.execute("SELECT * FROM users WHERE name = ?", (name,)).fetchone()
    return dict(row)


def list_users() -> list[dict]:
    init_db()
    with _lock, _conn() as c:
        return [dict(r) for r in c.execute(
            "SELECT * FROM users ORDER BY name").fetchall()]


def get_user(name: str) -> dict | None:
    init_db()
    with _lock, _conn() as c:
        row = c.execute("SELECT * FROM users WHERE name = ?", (name,)).fetchone()
    return dict(row) if row else None


# --- comments ----------------------------------------------------------


def add_comment(deck: str, slide_index: int, author: str, text: str,
                record_id: str | None = None) -> dict:
    text = text.strip()
    if not text:
        raise ValueError("empty comment")
    init_db()
    with _lock, _conn() as c:
        cur = c.execute(
            "INSERT INTO comments (deck, slide_index, record_id, author, text, created_at)"
            " VALUES (?, ?, ?, ?, ?, ?)",
            (deck, slide_index, record_id, author, text, _now()))
        row = c.execute("SELECT * FROM comments WHERE id = ?",
                        (cur.lastrowid,)).fetchone()
    return dict(row)


def comments_for(deck: str, slide_index: int | None = None) -> list[dict]:
    init_db()
    with _lock, _conn() as c:
        if slide_index is None:
            rows = c.execute("SELECT * FROM comments WHERE deck = ? ORDER BY id",
                             (deck,)).fetchall()
        else:
            rows = c.execute(
                "SELECT * FROM comments WHERE deck = ? AND slide_index = ? ORDER BY id",
                (deck, slide_index)).fetchall()
    return [dict(r) for r in rows]


def comment_counts(deck: str) -> dict[int, int]:
    init_db()
    with _lock, _conn() as c:
        rows = c.execute(
            "SELECT slide_index, COUNT(*) AS n FROM comments WHERE deck = ?"
            " GROUP BY slide_index", (deck,)).fetchall()
    return {r["slide_index"]: r["n"] for r in rows}


def delete_comment(comment_id: int, author: str) -> bool:
    """Authors delete their own comments; leads/admins delete any."""
    init_db()
    user = get_user(author)
    with _lock, _conn() as c:
        if user and user["role"] in ("lead", "admin"):
            cur = c.execute("DELETE FROM comments WHERE id = ?", (comment_id,))
        else:
            cur = c.execute("DELETE FROM comments WHERE id = ? AND author = ?",
                            (comment_id, author))
    return cur.rowcount > 0


# --- authentication (pilot-grade: name + PIN) ---------------------------
# PINs are set on first sign-in and stored as salted PBKDF2 hashes. This is
# honest pilot security for a loopback tool; production auth is Entra ID.

import hashlib
import secrets


def _ensure_auth_tables() -> None:
    init_db()
    with _lock, _conn() as c:
        cols = {r["name"] for r in c.execute("PRAGMA table_info(users)").fetchall()}
        if "pin_hash" not in cols:
            c.execute("ALTER TABLE users ADD COLUMN pin_hash TEXT")
        c.execute("""CREATE TABLE IF NOT EXISTS sessions (
            token TEXT PRIMARY KEY,
            user_name TEXT NOT NULL,
            created_at TEXT NOT NULL)""")
        c.execute("""CREATE TABLE IF NOT EXISTS audits (
            id INTEGER PRIMARY KEY,
            deck TEXT NOT NULL,
            profile_id TEXT NOT NULL,
            profile_version INTEGER NOT NULL,
            user_name TEXT NOT NULL,
            slides INTEGER NOT NULL,
            errors INTEGER NOT NULL,
            warnings INTEGER NOT NULL,
            info INTEGER NOT NULL,
            arabic INTEGER NOT NULL,
            total INTEGER NOT NULL,
            kind TEXT NOT NULL DEFAULT 'audit',
            manifest TEXT NOT NULL,
            created_at TEXT NOT NULL)""")


def _hash_pin(pin: str, salt: str) -> str:
    return hashlib.pbkdf2_hmac("sha256", pin.encode(), salt.encode(),
                               200_000).hex()


def set_pin(name: str, pin: str) -> None:
    if len(pin.strip()) < 4:
        raise ValueError("PIN must be at least 4 characters")
    _ensure_auth_tables()
    salt = secrets.token_hex(8)
    with _lock, _conn() as c:
        c.execute("UPDATE users SET pin_hash = ? WHERE name = ?",
                  (f"{salt}${_hash_pin(pin.strip(), salt)}", name))


def verify_pin(name: str, pin: str) -> bool:
    _ensure_auth_tables()
    user = get_user(name)
    if not user or not user.get("pin_hash"):
        return False
    salt, expected = user["pin_hash"].split("$", 1)
    return secrets.compare_digest(_hash_pin(pin.strip(), salt), expected)


def has_pin(name: str) -> bool:
    _ensure_auth_tables()
    user = get_user(name)
    return bool(user and user.get("pin_hash"))


def create_session(name: str) -> str:
    _ensure_auth_tables()
    token = secrets.token_urlsafe(32)
    with _lock, _conn() as c:
        c.execute("INSERT INTO sessions (token, user_name, created_at) VALUES (?, ?, ?)",
                  (token, name, _now()))
    return token


def session_user(token: str) -> dict | None:
    if not token:
        return None
    _ensure_auth_tables()
    with _lock, _conn() as c:
        row = c.execute(
            "SELECT u.* FROM sessions s JOIN users u ON u.name = s.user_name"
            " WHERE s.token = ?", (token,)).fetchone()
    return dict(row) if row else None


def delete_session(token: str) -> None:
    _ensure_auth_tables()
    with _lock, _conn() as c:
        c.execute("DELETE FROM sessions WHERE token = ?", (token,))


# --- audit history -------------------------------------------------------


def record_audit(manifest: dict, user_name: str, kind: str = "audit") -> int:
    """Persist an audit run: metadata + the full manifest JSON, so reports
    survive server restarts. Deck bytes are never stored."""
    import json as _json

    _ensure_auth_tables()
    s = manifest["summary"]
    sev = s.get("by_severity", {})
    with _lock, _conn() as c:
        cur = c.execute(
            "INSERT INTO audits (deck, profile_id, profile_version, user_name,"
            " slides, errors, warnings, info, arabic, total, kind, manifest,"
            " created_at) VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)",
            (manifest["deck"], manifest["profile_id"], manifest["profile_version"],
             user_name, manifest["slides"], sev.get("error", 0),
             sev.get("warning", 0), sev.get("info", 0),
             s.get("arabic_flagged", 0), s.get("total", 0), kind,
             _json.dumps(manifest, ensure_ascii=False), _now()))
    return cur.lastrowid


def list_audits(limit: int = 100) -> list[dict]:
    _ensure_auth_tables()
    with _lock, _conn() as c:
        rows = c.execute(
            "SELECT id, deck, profile_id, profile_version, user_name, slides,"
            " errors, warnings, info, arabic, total, kind, created_at"
            " FROM audits ORDER BY id DESC LIMIT ?", (limit,)).fetchall()
    return [dict(r) for r in rows]


def get_audit(audit_id: int) -> dict | None:
    import json as _json

    _ensure_auth_tables()
    with _lock, _conn() as c:
        row = c.execute("SELECT * FROM audits WHERE id = ?", (audit_id,)).fetchone()
    if row is None:
        return None
    out = dict(row)
    out["manifest"] = _json.loads(out["manifest"])
    return out


def clear_pin(name: str) -> None:
    """Admin PIN reset: next sign-in sets a fresh PIN. Kills the user's
    sessions so a lost device cannot ride an old cookie."""
    _ensure_auth_tables()
    with _lock, _conn() as c:
        c.execute("UPDATE users SET pin_hash = NULL WHERE name = ?", (name,))
        c.execute("DELETE FROM sessions WHERE user_name = ?", (name,))


# --- backend selection ---------------------------------------------------
# When a DATABASE_URL is configured (Supabase on the cloud demo), override
# the SQLite implementations above with the Postgres backend. Local and test
# runs have no DATABASE_URL and keep SQLite untouched.
from .config import USING_POSTGRES as _USING_POSTGRES  # noqa: E402

if _USING_POSTGRES:  # pragma: no cover - exercised only against a live DB
    from .store_pg import (  # noqa: F401,E402
        add_comment, add_user, clear_pin, comment_counts, comments_for,
        create_session, delete_comment, delete_session, get_audit, get_user,
        has_pin, init_db, list_audits, list_users, record_audit, session_user,
        set_pin, verify_pin)
