"""Runtime configuration from environment variables.

Local/desktop default: SQLite, PowerPoint COM rendering, no auth gate.
Cloud demo (Render): set DATABASE_URL (Supabase), QC_AUTH_REQUIRED=1,
QC_STRICT_SESSIONS=1, QC_RENDERER=libreoffice, QC_DEMO_BANNER=<text>.
"""

import os
from pathlib import Path


def _load_dotenv():
    """Tiny .env loader (no dependency): KEY=VALUE lines from the project
    root, for local secrets like GEMINI_API_KEY. Real environment
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


# Master switch for every model-backed feature (assistant triage, the design
# copilot, component review, the chat box). QC_AI=0 turns them off at the
# source: the UI hides them and the routes refuse, so no request leaves this
# machine even if a key is configured. Defaults ON; set it in a gitignored .env
# to disable locally without changing what anyone else gets.
AI_ENABLED = opt_out_flag("QC_AI")

# GEMINI IS THE ONLY MODEL THIS BUILD TALKS TO (design lead, 31/08/2026).
#
# There was a QC_LLM_PROVIDER switch here with an Anthropic branch behind it,
# and a separate QC_ASSIST_MODEL naming a Claude tier for the triage questions.
# Both are gone. The switch bought optionality nobody used and cost a second
# schema dialect, a second refusal shape and a second key check that no real run
# ever exercised; the assist model was the last thing in the tool reaching for a
# different vendor than everything else. Everything that needs judgment now goes
# through qc.llm to Gemini, and the passes that read a slide are the only ones
# that send an image.
#
# Vision quality sets the ceiling on how good the fixes can be - a wrongly
# grouped component means a fix that tears a card off its label - so the default
# is the strongest vision tier rather than the cheapest.
#
# Gemini 3.1 Pro (design lead, 26/08/2026). Preview is the only id it answers
# on, and a "-preview" id can be retired by Google without notice, so this is
# the single line to change when it graduates or moves on; QC_LLM_MODEL
# overrides it on a host without a deploy.
DEFAULT_LLM_MODEL = "gemini-3.1-pro-preview"
LLM_MODEL = os.getenv("QC_LLM_MODEL", DEFAULT_LLM_MODEL).strip()


# A model id from a vendor this build no longer calls is the one settings
# mistake that survives the provider removal: a host still carrying
# QC_LLM_MODEL=claude-opus-5 in its .env posts that id to Gemini and gets a 404
# in every judgment pass at once, reported as "the model could not be reached"
# with nothing naming the variable responsible.
#
# Caught here, where the value is in view. Not raised: this module is imported at
# startup and a stale .env must not stop the server booting, since every page
# that is not a judgment pass still works. The pass says so when it is asked, and
# the routes print `model_note` beside the button.
def _model_note() -> str:
    """Why the configured model id will not answer, or "" when it looks fine."""
    override = os.getenv("QC_LLM_MODEL", "").strip()
    if not override or override.lower().startswith("gemini"):
        return ""
    return (f"QC_LLM_MODEL is {override!r}, which is not a Gemini model. This "
            f"build only calls Gemini; unset it to use {DEFAULT_LLM_MODEL}.")


LLM_MODEL_NOTE = _model_note()


def _positive_int(name: str, default: int) -> int:
    """An env-tunable count that refuses to be zero or junk. A 0 here would
    mean "no timeout" or "no attempts", which are the two settings nobody
    wants to reach by typo."""
    try:
        value = int(os.getenv(name, str(default)))
    except ValueError:
        return default
    return value if value > 0 else default


# How long one judgment call may take before it is abandoned, in MILLISECONDS
# (the unit google-genai's HttpOptions takes; passing seconds here would set a
# two-minute budget to two milliseconds). Generous by default: these are vision
# calls over dense slide images at MEDIA_RESOLUTION_HIGH with thinking on, and
# a slow answer is still worth more than no answer. The point is that a hung
# connection cannot pin a worker forever, not that slow calls are cut off.
LLM_TIMEOUT_MS = _positive_int("QC_LLM_TIMEOUT_MS", 120_000)

# Attempts INCLUDING the first, on the transient statuses only (408, 429, 5xx).
# The SDK's own default is 5 with delays growing to 60s, which is tuned for
# cheap calls; a run here makes one of these per distinct slide structure, so
# five attempts on a bad afternoon is a designer watching a spinner. Three
# attempts with a short ceiling retries the blip and gives up on the outage.
LLM_ATTEMPTS = _positive_int("QC_LLM_ATTEMPTS", 3)
LLM_RETRY_MAX_DELAY_SEC = 8.0

# How many judgment calls may be in flight at once.
#
# These passes ask one question PER SLIDE (or per distinct slide structure), the
# questions do not depend on each other, and each one is a vision call over a
# dense image with thinking on - seconds, not milliseconds, and almost all of it
# spent waiting on the network. Run in a row, a twenty-slide component review is
# a designer watching a spinner for several minutes; the work is embarrassingly
# parallel and was serial only because it was written a slide at a time.
#
# Four rather than twenty: the ceiling here is the provider's rate limit, and a
# burst of twenty turns into 429s that the SDK then has to sit out, which is
# slower than not having burst at all. Raise it on a host with a high quota.
LLM_CONCURRENCY = _positive_int("QC_LLM_CONCURRENCY", 4)

# Max upload size in MB. Raise on the desktop/LAN box for large image-heavy
# client decks; keep modest on small-RAM cloud tiers (the whole file is held
# in memory during processing, and recent decks are retained for the fix
# step). Default 250 MB.
try:
    MAX_UPLOAD_MB = int(os.getenv("QC_MAX_UPLOAD_MB", "250"))
except ValueError:
    MAX_UPLOAD_MB = 250
