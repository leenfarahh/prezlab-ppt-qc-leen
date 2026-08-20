"""Pilot identity resolution: session token first, legacy name cookie second.

The legacy qc_user cookie remains accepted as a fallback so existing tests
and scripts keep working; the sign-in page issues real session tokens. Both
are pilot-grade attribution; production auth is Entra ID (PRD Section 6).
"""


# When True (LAN or cloud deployment), only real session tokens count; the
# legacy qc_user name cookie is ignored because names are guessable.
from .config import STRICT_SESSIONS as _ENV_STRICT

STRICT_SESSIONS = _ENV_STRICT


def current_user(request) -> dict | None:
    from .store import get_user, session_user

    user = session_user(request.cookies.get("qc_session", ""))
    if user:
        return user
    if STRICT_SESSIONS:
        return None
    legacy = request.cookies.get("qc_user", "")
    return get_user(legacy) if legacy else None
