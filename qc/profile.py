"""Formatting profile loading (PRD Appendix A.6).

Profiles are JSON documents seeded in qc/profiles/. Modules read config via
dotted paths, e.g. profile.get("geometry.safe_zone_margins_emu.left").
"""

import json
from pathlib import Path

PROFILES_DIR = Path(__file__).resolve().parent / "profiles"


class Profile:
    def __init__(self, data: dict):
        self.id = data.get("id", "unknown")
        self.name = data.get("name", self.id)
        self.version = data.get("version", 1)
        self.config = data.get("config", {})
        self.raw = data

    @classmethod
    def load(cls, name_or_path: str) -> "Profile":
        p = Path(name_or_path)
        if not p.exists():
            p = PROFILES_DIR / f"{name_or_path}.json"
        if not p.exists():
            available = sorted(f.stem for f in PROFILES_DIR.glob("*.json"))
            raise FileNotFoundError(
                f"Profile '{name_or_path}' not found. Available: {available}")
        return cls(json.loads(p.read_text(encoding="utf-8")))

    def get(self, dotted: str, default=None):
        node = self.config
        for key in dotted.split("."):
            if not isinstance(node, dict) or key not in node:
                return default
            node = node[key]
        return node

    def module_config(self, module_key: str) -> dict:
        return self.config.get(module_key, {})
