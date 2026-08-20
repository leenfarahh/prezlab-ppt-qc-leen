"""Master templates kept alongside profiles, so a profile can be APPLIED.

A Style Spec describes a design system; it cannot be applied as one. Restyling
a slide means handing PowerPoint real slideLayout parts to match placeholders
against, and no amount of extracted numbers substitutes for them. So when a
spec is saved as a profile, the master file it came from is saved too, and
that pair is what "use this profile to format a deck" needs.

Templates live under data/templates/, not in qc/profiles/ next to the JSON:
profiles are versioned text a design lead reviews and edits, while these are
opaque client binaries. data/ is already gitignored, which is the behaviour
client material needs by default.
"""

import hashlib
import re
from pathlib import Path

from . import store

_SAFE_ID = re.compile(r"^[A-Za-z0-9_\-:]+$")


def templates_dir() -> Path:
    # The MODULE is imported, never `from .store import DATA_DIR`: the test
    # suite repoints qc.store.DATA_DIR per test, and importing the value binds
    # a copy that no monkeypatch can reach. Getting this wrong sent test
    # templates into the real data/ directory while the isolation test still
    # passed, because it asserted on the same wrong path it was writing to.
    return Path(store.DATA_DIR) / "templates"


def _path(profile_id: str) -> Path | None:
    if not profile_id or not _SAFE_ID.match(profile_id):
        return None
    # ':' is legal in our ids (master:foo) but not in a Windows filename.
    return templates_dir() / f"{profile_id.replace(':', '__')}.pptx"


def save_master(profile_id: str, master_bytes: bytes) -> Path | None:
    path = _path(profile_id)
    if path is None:
        return None
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_bytes(master_bytes)
    return path


def load_master(profile_id: str) -> bytes | None:
    path = _path(profile_id)
    if path is None or not path.exists():
        return None
    return path.read_bytes()


def has_master(profile_id: str) -> bool:
    path = _path(profile_id)
    return path is not None and path.exists()


def master_info(profile_id: str) -> dict | None:
    """Size and digest, for showing a designer which file a profile carries."""
    path = _path(profile_id)
    if path is None or not path.exists():
        return None
    blob = path.read_bytes()
    return {"bytes": len(blob),
            "sha1": hashlib.sha1(blob).hexdigest(),
            "modified": path.stat().st_mtime}


def delete_master(profile_id: str) -> bool:
    path = _path(profile_id)
    if path is None or not path.exists():
        return False
    path.unlink()
    return True
