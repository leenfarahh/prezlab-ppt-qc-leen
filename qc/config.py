"""Runtime configuration from environment variables.

Local/desktop default: SQLite, PowerPoint COM rendering, no auth gate.
Cloud demo (Render): set DATABASE_URL (Supabase), QC_AUTH_REQUIRED=1,
QC_STRICT_SESSIONS=1, QC_RENDERER=libreoffice, QC_DEMO_BANNER=<text>.
"""

import os
from pathlib import Path


def _load_dotenv():
    """Tiny .env loader (no dependency): KEY=VALUE lines from the project
    root, for local secrets like ANTHROPIC_API_KEY. Real environment
    variables always win; blank values and comments are skipped."""
    env_path = Path(__file__).resolve().parent.parent / ".env"
    if not env_path.exists():
        return
    for line in env_path.read_text(encoding="utf-8").splitlines():
        line = line.strip()
        if not line or line.startswith("#") or "=" not in line:
            continue
        key, _, value = line.partition("=")
        key, value = key.strip(), value.strip().strip('"').strip("'")
        if key and value and key not in os.environ:
            os.environ[key] = value


_load_dotenv()

# Supabase / Postgres connection string. Empty => local SQLite store.
DATABASE_URL = os.getenv("DATABASE_URL", "").strip()

# Auth enforcement (mandatory sign-in on every non-public route).
AUTH_REQUIRED = os.getenv("QC_AUTH_REQUIRED", "").strip() in ("1", "true", "yes")
STRICT_SESSIONS = os.getenv("QC_STRICT_SESSIONS", "").strip() in ("1", "true", "yes")

# Renderer: auto (try COM, else LibreOffice), com, libreoffice, or none.
RENDERER = os.getenv("QC_RENDERER", "auto").strip().lower()

# Shown as a warning strip on every page when set (e.g. the demo instance).
DEMO_BANNER = os.getenv("QC_DEMO_BANNER", "").strip()

# First-run admin: created on startup when the user table is empty, so a
# fresh cloud deploy has someone who can sign in (auth is mandatory there,
# which would otherwise lock everyone out). They set their PIN on first
# sign-in. No effect once any user exists.
BOOTSTRAP_ADMIN = os.getenv("QC_BOOTSTRAP_ADMIN", "").strip()

# Cloud demos are for non-confidential decks only; the LAN box keeps client work.
USING_POSTGRES = bool(DATABASE_URL)

def opt_out_flag(name: str) -> bool:
    """An ON-by-default switch: only an explicit falsey value disables it.

    Distinct from the opt-IN reads above (AUTH_REQUIRED and friends), where an
    unset variable means off. Named and separate so the default direction of a
    switch is visible at its definition rather than inferred from a comparison."""
    return os.getenv(name, "1").strip().lower() not in ("0", "false", "no", "off")


# Master switch for every Anthropic-backed feature (assistant triage and the
# design copilot). QC_AI=0 turns them off at the source: the UI hides them and
# the routes refuse, so no request leaves this machine even if a key is
# configured. Defaults ON; set it in a gitignored .env to disable locally
# without changing what anyone else gets.
AI_ENABLED = opt_out_flag("QC_AI")

# Assistant triage (qc/assist.py): Claude phrases the clarifying questions
# when an Anthropic API key is configured (ANTHROPIC_API_KEY); without one
# the same questions fall back to template phrasing, fully offline. Only
# finding metadata is ever sent, never slide content. Haiku is the default
# because the task is small (select and phrase a handful of questions,
# ~$0.006 per click); set QC_ASSIST_MODEL to a bigger tier if selection
# quality ever disappoints.
ASSIST_MODEL = os.getenv("QC_ASSIST_MODEL", "claude-haiku-4-5").strip()

# Design copilot (qc/copilot.py): Claude reviews rendered slide images and
# proposes layout actions; code verifies and the designer approves. Vision
# quality matters here, so the strongest tier is the default. Unlike the
# assistant, this sends slide IMAGES to the API - the UI discloses it.
COPILOT_MODEL = os.getenv("QC_COPILOT_MODEL", "claude-opus-4-8").strip()

# The judgment passes - component review (qc/components.py) and layout matching
# when applying a master (qc/layoutmatch.py) - all go through qc.llm, and this
# is the only place their provider and model are chosen. Gemini by default
# (design lead, 24/08/2026); "anthropic" is still wired, so a provider switch is
# not a one-way door.
#
# These send slide IMAGES, like the design copilot, and the UI discloses it.
# Vision quality sets the ceiling on how good the fixes can be - a wrongly
# grouped component means a fix that tears a card off its label - so the default
# is the strongest vision tier rather than the cheapest.
LLM_PROVIDER = os.getenv("QC_LLM_PROVIDER", "gemini").strip().lower()
_DEFAULT_MODELS = {"gemini": "gemini-3-pro-preview",
                   "anthropic": "claude-opus-5"}
LLM_MODEL = os.getenv(
    "QC_LLM_MODEL", _DEFAULT_MODELS.get(LLM_PROVIDER, "gemini-3-pro-preview")
).strip()

# Max upload size in MB. Raise on the desktop/LAN box for large image-heavy
# client decks; keep modest on small-RAM cloud tiers (the whole file is held
# in memory during processing, and recent decks are retained for the fix
# step). Default 250 MB.
try:
    MAX_UPLOAD_MB = int(os.getenv("QC_MAX_UPLOAD_MB", "250"))
except ValueError:
    MAX_UPLOAD_MB = 250
