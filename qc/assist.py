"""Assistant triage: turn recurring ambiguous findings into clarifying
questions whose answers become deterministic profile updates.

The findings the tool cannot fix are exactly the ones where it lacks INTENT
(is this dark red a brand color, or a mistake repeated 40 times?). A short
question captures that intent once; the answer updates the profile, so the
same ambiguity never gets triaged again.

Flow: aggregate(manifest) builds candidate questions deterministically from
the finding metadata; generate_questions() asks the model to select and phrase
the few worth a designer's time (falling back to template phrasing when no API
is configured, so the feature degrades gracefully to fully offline);
apply_actions() writes accepted answers into the profile with a version bump.

Confidentiality: only finding METADATA is ever sent to the model - issue counts,
hex colors, font family names, margin numbers, the profile name. Never slide
text, never images, never the deck filename. This is the one judgment pass in
the tool that sends no picture. Every action the model proposes is validated
against the locally computed candidates before it is shown or applied; the model
chooses and phrases, it does not invent values.
"""

import json
import re
import uuid
from collections import Counter

MIN_COLOR_USES = 3
MIN_FONT_USES = 3
MIN_MARGIN_BREACHES = 10
MAX_QUESTIONS = 5
_HEX = re.compile(r"^[0-9A-Fa-f]{6}$")

# structured-output schema: the model returns questions whose actions are
# drawn from a closed vocabulary; values are validated against the
# aggregates server-side regardless
_ACTION_ADD_COLOR = {
    "type": "object", "additionalProperties": False,
    "required": ["type", "hex", "name"],
    "properties": {
        "type": {"const": "add_color"},
        "hex": {"type": "string"},
        "name": {"type": "string"},
    },
}
_ACTION_ADD_FONT = {
    "type": "object", "additionalProperties": False,
    "required": ["type", "family", "role", "script"],
    "properties": {
        "type": {"const": "add_font"},
        "family": {"type": "string"},
        "role": {"type": "string"},
        "script": {"enum": ["latin", "complex"]},
    },
}
_ACTION_SET_MARGINS = {
    "type": "object", "additionalProperties": False,
    "required": ["type", "left", "top", "right", "bottom"],
    "properties": {
        "type": {"const": "set_margins"},
        "left": {"type": "integer"}, "top": {"type": "integer"},
        "right": {"type": "integer"}, "bottom": {"type": "integer"},
    },
}
QUESTIONS_SCHEMA = {
    "type": "object", "additionalProperties": False,
    "required": ["questions"],
    "properties": {
        "questions": {
            "type": "array",
            "items": {
                "type": "object", "additionalProperties": False,
                "required": ["question", "rationale", "impact", "action"],
                "properties": {
                    "question": {"type": "string"},
                    "rationale": {"type": "string"},
                    "impact": {"type": "string"},
                    "action": {"anyOf": [_ACTION_ADD_COLOR, _ACTION_ADD_FONT,
                                         _ACTION_SET_MARGINS]},
                },
            },
        },
    },
}

_SYSTEM = """You are the triage assistant inside Prezlab's internal
PowerPoint formatting QC tool. Designers audit client decks against a
formatting profile (allowed fonts, brand palette, safe-zone margins). You
receive aggregated audit findings as JSON: recurring off-palette colors,
recurring out-of-set fonts, and safe-zone pressure, plus what the profile
currently allows.

Select at most {max_q} candidates that are genuinely worth a designer's
time and phrase each as one clear yes/no question a designer can answer at
a glance. Prefer high-count recurring signals over noise; skip anything a
designer would not recognize. Write in clear, professional US English, in
a consultative tone. Do not use em dashes.

Rules for actions:
- Use ONLY values present in the candidates JSON (hex codes, family names,
  roles, scripts, margin numbers). Never invent values.
- For colors, propose a short descriptive name (e.g. "deep red").
- The impact field states plainly what accepting does, e.g. "removes 40
  recurring color errors on future audits of decks like this one".
- Fewer, better questions beat coverage. If nothing is signal, return an
  empty list."""


# ------------------------------------------------------------- aggregation

def aggregate(manifest: dict, deck_bytes: bytes | None = None,
              profile_cfg: dict | None = None) -> dict:
    """Candidate questions computed deterministically from finding metadata.
    No LLM involved; this is also the payload sent to it (metadata only)."""
    records = manifest.get("records") or []

    colors: Counter = Counter()
    color_slides: dict[str, set] = {}
    fonts: Counter = Counter()
    breaches = 0
    for r in records:
        it = r.get("issue_type")
        if it == "color_palette.off_palette_rgb" and r.get("old_value"):
            hexval = str(r["old_value"]).lstrip("#").upper()
            if _HEX.match(hexval):
                colors[hexval] += 1
                color_slides.setdefault(hexval, set()).add(r["slide_index"])
        elif it == "font.family_out_of_set" and r.get("old_value"):
            rule = r.get("profile_rule_id") or ""
            m = re.match(r"font\.roles\.(\w+)\.(latin|complex_script)", rule)
            if m:
                script = "latin" if m.group(2) == "latin" else "complex"
                fonts[(str(r["old_value"]), m.group(1), script)] += 1
        elif it == "margin_alignment.outside_safe_zone" \
                and r.get("action") == "flagged":
            # Body content only, by issue type: headings past a margin carry
            # margin_alignment.heading_past_margin instead, and a title the
            # client wants running wide is not evidence that the PROFILE's
            # margins are too tight.
            breaches += 1

    agg = {
        "profile": {
            "palette": [c.get("hex") for c in
                        (profile_cfg or {}).get("color_palette", {})
                        .get("named_colors", [])],
            "margins_emu": (profile_cfg or {}).get("geometry", {})
            .get("safe_zone_margins_emu", {}),
        },
        "colors": [
            {"hex": h, "count": n, "slides": len(color_slides[h])}
            for h, n in colors.most_common(8) if n >= MIN_COLOR_USES
        ],
        "fonts": [
            {"family": fam, "role": role, "script": script, "count": n}
            for (fam, role, script), n in fonts.most_common(8)
            if n >= MIN_FONT_USES
        ],
        "margins": None,
    }

    if breaches >= MIN_MARGIN_BREACHES and deck_bytes:
        proposal = _propose_margins(deck_bytes,
                                    agg["profile"]["margins_emu"])
        if proposal:
            agg["margins"] = {"breaches": breaches, "proposed": proposal}
    return agg


def _propose_margins(deck_bytes: bytes, current: dict) -> dict | None:
    """Margins the deck actually respects (qc.bootstrap.learn_margins),
    offered only when meaningfully tighter than the profile (i.e. the
    profile margin, not the deck, is the outlier)."""
    import io

    from pptx import Presentation

    from .bootstrap import learn_margins

    try:
        proposal = learn_margins(Presentation(io.BytesIO(deck_bytes)))
    except Exception:
        return None
    if proposal is None:
        return None
    # only worth asking when at least one side relaxes by >10%
    if not any(proposal[k] < 0.9 * current.get(k, 0) for k in proposal):
        return None
    return proposal


# ----------------------------------------------------------- question gen

def _validate(question: dict, agg: dict) -> dict | None:
    """Actions must reference values from the aggregates; anything else is
    dropped (the model chooses and phrases, it does not invent)."""
    action = question.get("action") or {}
    kind = action.get("type")
    if kind == "add_color":
        hexval = str(action.get("hex", "")).lstrip("#").upper()
        if hexval not in {c["hex"] for c in agg["colors"]}:
            return None
        action["hex"] = hexval
        action["name"] = str(action.get("name") or f"deck color {hexval}")[:40]
    elif kind == "add_font":
        key = (action.get("family"), action.get("role"), action.get("script"))
        if key not in {(f["family"], f["role"], f["script"])
                       for f in agg["fonts"]}:
            return None
    elif kind == "set_margins":
        if not agg.get("margins"):
            return None
        # the numbers are ours, never the model's
        action.update(agg["margins"]["proposed"])
    else:
        return None
    question["action"] = action
    question["id"] = uuid.uuid4().hex[:12]
    return question


def _fallback_questions(agg: dict) -> list[dict]:
    """Template phrasing when no API is configured: same questions, fully
    offline."""
    out = []
    for c in agg["colors"][:3]:
        out.append({
            "question": f"Color #{c['hex']} appears {c['count']} times across "
                        f"{c['slides']} slide(s) but is not in the profile "
                        "palette. Is it a brand color that belongs in the "
                        "profile?",
            "rationale": "Recurring use across slides usually means intent, "
                         "not error.",
            "impact": f"Accepting removes {c['count']} recurring color errors "
                      "from future audits.",
            "action": {"type": "add_color", "hex": c["hex"],
                       "name": f"deck color {c['hex']}"},
        })
    for f in agg["fonts"][:2]:
        out.append({
            "question": f"Font '{f['family']}' is used in {f['count']} runs "
                        f"for the {f['role']} role but is not in the allowed "
                        "set. Should it be allowed?",
            "rationale": "A consistently used family is usually the deck's "
                         "intended typeface.",
            "impact": f"Accepting removes {f['count']} font errors from "
                      "future audits.",
            "action": {"type": "add_font", "family": f["family"],
                       "role": f["role"], "script": f["script"]},
        })
    if agg.get("margins"):
        m = agg["margins"]
        out.append({
            "question": f"{m['breaches']} shapes sit outside the profile's "
                        "safe-zone margins, but the deck's content "
                        "consistently uses tighter margins. Relax the "
                        "profile margins to match the deck?",
            "rationale": "When most content breaches the margin, the margin "
                         "is usually what's wrong.",
            "impact": f"Accepting silences ~{m['breaches']} safe-zone "
                      "warnings on decks laid out like this one.",
            "action": {"type": "set_margins", **m["proposed"]},
        })
    return out


def api_configured() -> bool:
    """Whether the model can phrase these questions on this host.

    Delegates to qc.llm rather than reading a key itself. This module built its
    own client against a second vendor until 31/08/2026, which meant the one
    pass in the tool that sends no image was also the one pass configured
    somewhere else, answering to a different key and a different model
    variable - so a host with Gemini configured got template phrasing here and
    a model everywhere else, with nothing on the page explaining the
    difference."""
    from .llm import api_configured as _configured

    return _configured()


def _ask_model(agg: dict) -> list[dict]:
    """The aggregates, phrased. No images: this pass reads finding METADATA
    only, which is the whole reason it can run on decks whose slides must not
    leave the machine."""
    from .llm import ask_json

    answer = ask_json(system=_SYSTEM.format(max_q=MAX_QUESTIONS),
                      prompt=json.dumps(agg, sort_keys=True),
                      schema=QUESTIONS_SCHEMA)
    return answer["questions"]


def generate_questions(agg: dict) -> tuple[list[dict], str]:
    """(questions, source). The model selects and phrases when configured;
    otherwise (or on any API failure) the template fallback asks the same
    questions offline. Either way every action is validated locally."""
    if not (agg["colors"] or agg["fonts"] or agg.get("margins")):
        return [], "none"
    raw, source = None, "assistant"
    if api_configured():
        try:
            raw = _ask_model(agg)
        except Exception as exc:
            raw, source = None, f"fallback ({type(exc).__name__})"
    else:
        source = "fallback (no API key configured)"
    if raw is None:
        raw = _fallback_questions(agg)
    out = []
    for q in raw[:MAX_QUESTIONS]:
        valid = _validate(q, agg)
        if valid:
            out.append(valid)
    return out, source


# ------------------------------------------------------------------ apply

def apply_actions(pid: str, actions: list[dict], editor: str) -> dict:
    """Write accepted answers into the profile: version bump, owner stamp.
    Returns {"version": n, "applied": [summaries]}."""
    from datetime import date

    import qc.profile as profile_mod

    path = profile_mod.PROFILES_DIR / f"{pid}.json"
    data = json.loads(path.read_text(encoding="utf-8"))
    cfg = data.setdefault("config", {})
    applied = []

    for action in actions:
        kind = action.get("type")
        if kind == "add_color":
            named = cfg.setdefault("color_palette", {}) \
                .setdefault("named_colors", [])
            if not any(c.get("hex") == action["hex"] for c in named):
                named.append({"name": action["name"], "hex": action["hex"],
                              "theme_ref": None, "allowed_tints": [],
                              "allowed_shades": []})
            applied.append(f"palette + #{action['hex']} ({action['name']})")
        elif kind == "add_font":
            key = "latin" if action["script"] == "latin" else "complex_script"
            fams = cfg.setdefault("font", {}).setdefault("roles", {}) \
                .setdefault(action["role"], {}).setdefault(key, [])
            if action["family"] not in fams:
                fams.append(action["family"])
            applied.append(f"fonts + {action['family']} "
                           f"({action['role']}/{action['script']})")
        elif kind == "set_margins":
            margins = cfg.setdefault("geometry", {}) \
                .setdefault("safe_zone_margins_emu", {})
            for side in ("left", "top", "right", "bottom"):
                margins[side] = int(action[side])
            applied.append("safe-zone margins relaxed to match the deck")

    data["version"] = int(data.get("version", 1)) + 1
    data["owner"] = (f"{editor} (assistant answer, "
                     f"{date.today().strftime('%d/%m/%Y')})")
    path.write_text(json.dumps(data, indent=2, ensure_ascii=False),
                    encoding="utf-8")
    return {"version": data["version"], "applied": applied}
